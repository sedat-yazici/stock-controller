import os
import sys
import json
import time
import random
import requests
import openpyxl
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Force UTF-8 encoding on Windows console
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def log(msg):
    print(msg, flush=True)

# Load environment configuration from .env
load_dotenv(dotenv_path=os.path.join(os.getcwd(), ".env"), override=True)

# Configuration & Environment Variables
BASE_URL = os.getenv("BASE_URL", "").rstrip("/")
SERVICE_NAME = os.getenv("SERVICE_NAME", "")
LOGIN_URL = os.getenv("LOGIN_URL", f"{BASE_URL}/api/magaza/uyelik/giris?hizmetAdi={SERVICE_NAME}")
ACCOUNT_URL = os.getenv("ACCOUNT_URL", f"{BASE_URL}/api/magaza/uyelik/hesabim?hizmetAdi={SERVICE_NAME}")
USERNAME = os.getenv("USERNAME", "")
PASSWORD = os.getenv("PASSWORD", "")
CACHE_FILE = os.getenv("SESSION_CACHE_FILE", "session_cache.json")
OUTPUT_FILE = os.getenv("CATALOGUE_OUTPUT_PATH", "catalogue_export.xlsx")
DELAY_BASE = float(os.getenv("REQUEST_DELAY_SECONDS", "0.5"))
MAX_WORKERS = int(os.getenv("CRAWLER_CONCURRENCY", "8"))

DEFAULT_HEADERS = {
    "User-Agent": os.getenv("USER_AGENT", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1"
}

def safe_float(val):
    """Safely converts string or number to float, returning None on empty string or error."""
    if val is None:
        return None
    val_str = str(val).strip().replace(",", ".")
    if not val_str:
        return None
    try:
        return float(val_str)
    except (ValueError, TypeError):
        return None

def polite_sleep(base_seconds):
    """Adds a small randomized jitter to prevent algorithmic rate detection."""
    if base_seconds <= 0:
        return
    jitter = random.uniform(0.7, 1.3)
    time.sleep(base_seconds * jitter)

def convert_to_usd(amount, source_curr, rates):
    """Converts price amount from source currency to USD using portal live exchange rates."""
    if amount is None:
        return None
    source = str(source_curr or "USD").strip().upper()
    if source == "TRY": source = "TL"
    if source == "EURO": source = "EUR"
    
    if source == "USD":
        return round(float(amount), 2)
        
    source_rate = rates.get(source, 1.0)
    usd_rate = rates.get("USD", 1.0)
    if usd_rate <= 0:
        usd_rate = 1.0
        
    converted = (float(amount) * source_rate) / usd_rate
    return round(converted, 2)

def extract_exchange_rates(soup):
    """Extracts live exchange rates from HTML root attributes (e.g. data-currency-rate-*)."""
    rates = {"TL": 1.0, "KARMA": 1.0, "USD": 1.0, "EUR": 1.0}
    html_tag = soup.find("html")
    if html_tag:
        for attr, val in html_tag.attrs.items():
            if attr.startswith("data-currency-rate-"):
                c_code = attr.replace("data-currency-rate-", "").upper()
                try:
                    rates[c_code] = float(str(val).replace(",", "."))
                except ValueError:
                    pass
    return rates

def safe_save_excel(wb, output_path):
    """Saves workbook safely, handling open file lock conflicts gracefully."""
    try:
        wb.save(output_path)
        log(f"[Excel] Successfully saved results to: {output_path}")
    except PermissionError:
        backup_path = output_path.replace(".xlsx", f"_backup_{int(time.time())}.xlsx")
        log(f"[Excel] Warning: {output_path} is locked by another program (e.g. open in Excel).")
        wb.save(backup_path)
        log(f"[Excel] Saved to alternative file: {backup_path}")
    except Exception as e:
        log(f"[Excel] Error saving workbook: {e}")

def get_authenticated_session():
    """Initializes and returns an authenticated requests.Session using cache or credentials."""
    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)
    
    # 1. Try existing session cache
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cache_data = json.load(f)
            token = cache_data.get("token")
            if token:
                session.headers.update({"Authorization": f"Bearer {token}"})
                check_res = session.get(ACCOUNT_URL, timeout=10)
                if check_res.status_code == 200:
                    log("[Auth] Successfully authenticated using cached session.")
                    return session
                else:
                    log("[Auth] Cached session expired or invalid. Re-authenticating...")
                    session.headers.pop("Authorization", None)
        except Exception as e:
            log(f"[Auth] Warning reading cache: {e}")
            session.headers.pop("Authorization", None)

    # 2. Re-authenticate via API if credentials are provided in .env
    if USERNAME and PASSWORD:
        log(f"[Auth] Authenticating via API as '{USERNAME}'...")
        payload = {
            "hizmetAdi": SERVICE_NAME,
            "kullanici": USERNAME,
            "sifre": PASSWORD
        }
        try:
            res = session.post(LOGIN_URL, json=payload, timeout=15)
            if res.status_code == 200:
                token = res.json().get("token")
                if token:
                    session.headers.update({"Authorization": f"Bearer {token}"})
                    try:
                        with open(CACHE_FILE, "w", encoding="utf-8") as f:
                            json.dump({"token": token, "saved_at": time.time()}, f)
                    except Exception as e:
                        log(f"[Auth] Could not write cache file: {e}")
                    log("[Auth] Login successful.")
                    return session
            log(f"[Auth] Login failed with status code {res.status_code}")
        except Exception as e:
            log(f"[Auth] Login request error: {e}")
    else:
        log("[Auth] No login credentials found in .env. Proceeding as guest/unauthenticated...")
        
    return session

def fetch_product_stock_code_and_description(session, url):
    """Fetches a single product detail page and extracts the stock code (stok kodu)."""
    try:
        res = session.get(url, timeout=12)
        description = ""
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            # Select stock code
            strong_el = soup.select_one("div.vs-product-facts > div:nth-child(1) > strong") or soup.select_one(".vs-product-facts strong")
            if strong_el:
                stock_code = strong_el.get_text(strip=True)
                try:    
                    # Select description panel
                    desc_div = soup.find(id="product-description-panel")
                    # Remove navigation bar
                    nav_el = desc_div.find("nav")
                    if nav_el:
                        nav_el.decompose()
                    # Remove wholesaler's contact info
                    wholesaler_info = desc_div.find(id="09-toptan-fiyatlar-icin-nexpay360") or desc_div.find(id="nexpay360-guvencesi") or desc_div.find(id="nexpay360-satis-urun-bilgisi") or desc_div.find(id="nexpay360-satis-toptan-teklif")
                    if wholesaler_info:
                        for el in wholesaler_info.find_all_next():
                            if el.parent == desc_div:
                                el.decompose()
                        wholesaler_info.decompose()

                    description = desc_div.get_text()
                except Exception as e:
                    print(f"[Error] {stock_code} ürün açıklaması bulunamadı. Hata kodu: {e}")
                return stock_code, description
    except Exception:
        pass
    return "Bulunamadı", description

def extract_full_catalogue():
    """Extracts all products, converts prices to USD, scrapes stock codes, and writes trimmed Excel report."""
    session = get_authenticated_session()
    
    products = []
    seen_urls = set()
    portal_rates = {"TL": 1.0, "KARMA": 1.0, "USD": 1.0, "EUR": 1.0}
    page = 1
    max_retries = 3
    
    log(f"\n[Phase 1] Crawling catalogue listing pages from {BASE_URL}...")
    start_time = time.time()
    
    while True:
        url = f"{BASE_URL}/urunler?sayfa={page}"
        response = None
        
        for attempt in range(1, max_retries + 1):
            try:
                response = session.get(url, timeout=15)
                if response.status_code == 200:
                    break
                elif response.status_code == 429:
                    wait_time = attempt * 5
                    log(f"[Crawl] Rate limited (HTTP 429) on page {page}. Sleeping {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    time.sleep(attempt * 2)
            except Exception as e:
                time.sleep(attempt * 2)

        if not response or response.status_code != 200:
            log(f"[Crawl] Finished catalogue traversal at page {page}.")
            break
            
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Capture live exchange rates from HTML root tag
        if page == 1:
            portal_rates = extract_exchange_rates(soup)
            log(f"[Currency] Portal conversion rates loaded: USD={portal_rates.get('USD', 1.0)}, EUR={portal_rates.get('EUR', 1.0)}")
            
        cards = soup.find_all("article", class_="safir-product-card")
        if not cards:
            log(f"[Crawl] No product cards found on page {page}. Finished all available pages.")
            break
            
        page_products_count = 0
        for card in cards:
            title_a = card.select_one(".safir-card-title a")
            if not title_a:
                continue
                
            title = title_a.get_text(strip=True)
            link = title_a.get("href", "")
            if not link or link in seen_urls:
                continue
            seen_urls.add(link)
            
            # Dealer Price (Including VAT)
            dealer_el = card.select_one(".safir-card-dealer:not(.safir-dealer-net) strong")
            if dealer_el:
                raw_dealer_price = safe_float(dealer_el.get("data-vs-price-amount"))
                raw_currency = dealer_el.get("data-vs-price-currency", "USD")
            else:
                # Fallback to regular price if dealer price is absent
                reg_el = card.select_one(".safir-card-prices strong")
                raw_dealer_price = safe_float(reg_el.get("data-vs-price-amount")) if reg_el else None
                raw_currency = reg_el.get("data-vs-price-currency", "USD") if reg_el else "USD"
                
            # Convert to USD using portal live rates
            dealer_price_usd = convert_to_usd(raw_dealer_price, raw_currency, portal_rates)
            full_url = f"{BASE_URL}{link}" if link.startswith("/") else link
            
            products.append({
                "title": title,
                "dealer_price_usd": dealer_price_usd,
                "url": full_url,
                "stock_code": None,
                "description": None
            })
            page_products_count += 1
            
        log(f"  [Listing] Page {page:02d}: Extracted {page_products_count} products (Total: {len(products)})")
        page += 1
        polite_sleep(DELAY_BASE)

    total_count = len(products)
    log(f"\n[Phase 1 Complete] Found {total_count} products in {time.time() - start_time:.2f}s.")
    
    # Phase 2: Extract Stock Codes from individual product detail pages
    log(f"\n[Phase 2] Fetching stock codes for {total_count} products via {MAX_WORKERS} concurrent worker threads...")
    phase2_start = time.time()
    
    def worker_task(index, product_item):
        stock_code, description = fetch_product_stock_code_and_description(session, product_item["url"])
        return index, stock_code, description

    completed_count = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {executor.submit(worker_task, idx, p): idx for idx, p in enumerate(products)}
        for future in as_completed(future_map):
            idx, code, description = future.result()
            products[idx]["stock_code"] = code
            products[idx]["description"] = description
            completed_count += 1
            if completed_count % 50 == 0 or completed_count == total_count:
                log(f"  [Detail Progress] {completed_count}/{total_count} stock codes fetched ({completed_count / total_count * 100:.1f}%)")

    log(f"[Phase 2 Complete] All stock codes extracted in {time.time() - phase2_start:.2f}s.")

    # Phase 3: Create and format trimmed Excel workbook
    log(f"\n[Phase 3] Generating Excel workbook: {OUTPUT_FILE}...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ürün Fiyat Listesi"
    ws.views.sheetView[0].showGridLines = True
    
    # Required 4 columns
    headers = [
        "Ürün/Stok Kodu",
        "Ürün Adı",
        "Bayi Fiyatı (KDV Dahil)",
        "Ürün Linki",
        "Toptancı Açıklaması"
    ]
    ws.append(headers)
    
    # Header Styling
    header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    
    # Append Trimmed Rows
    data_font = Font(name="Segoe UI", size=10)
    for p in products:
        ws.append([
            p["stock_code"],
            p["title"],
            p["dealer_price_usd"],
            p["url"],
            p["description"]
        ])
        
    # Format Data Rows
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 20
        # Stock code
        c_code = ws.cell(row=row, column=1)
        c_code.font = data_font
        c_code.border = thin_border
        c_code.alignment = Alignment(horizontal="center", vertical="center")
        
        # Product Title
        c_title = ws.cell(row=row, column=2)
        c_title.font = data_font
        c_title.border = thin_border
        c_title.alignment = Alignment(horizontal="left", vertical="center")
        
        # Dealer Price in USD
        c_price = ws.cell(row=row, column=3)
        c_price.font = data_font
        c_price.border = thin_border
        if isinstance(c_price.value, (int, float)):
            c_price.number_format = '#,##0.00 "USD"'
            c_price.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c_price.alignment = Alignment(horizontal="center", vertical="center")
            
        # Product URL
        c_url = ws.cell(row=row, column=4)
        c_url.font = data_font
        c_url.border = thin_border
        c_url.alignment = Alignment(horizontal="left", vertical="center")
        
    # Adjust Column Widths
    ws.column_dimensions['A'].width = 22  # Ürün/Stok Kodu
    ws.column_dimensions['B'].width = 50  # Ürün Adı
    ws.column_dimensions['C'].width = 26  # Bayi Fiyatı (KDV Dahil)
    ws.column_dimensions['D'].width = 60  # Ürün Linki
    ws.column_dimensions['E'].width = 100  # Toptancı Açıklaması
    safe_save_excel(wb, OUTPUT_FILE)
    total_time = time.time() - start_time
    log(f"\n[Done] Successfully processed {total_count} products and saved to '{OUTPUT_FILE}' in {total_time:.2f}s!")

if __name__ == "__main__":
    extract_full_catalogue()
