import os
import sys
import json
import time
import random
import requests
import openpyxl
from bs4 import BeautifulSoup
from dotenv import load_dotenv
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

def extract_full_catalogue():
    """Extracts all products from the paginated catalogue and saves to an Excel file."""
    session = get_authenticated_session()
    
    products = []
    seen_urls = set()
    page = 1
    max_retries = 3
    
    log(f"\n[Crawl] Starting catalogue extraction from {BASE_URL}...")
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
                    log(f"[Crawl] HTTP {response.status_code} on page {page} (attempt {attempt}/{max_retries})")
                    time.sleep(attempt * 2)
            except Exception as e:
                log(f"[Crawl] Request error on page {page} (attempt {attempt}/{max_retries}): {e}")
                time.sleep(attempt * 2)

        if not response or response.status_code != 200:
            log(f"[Crawl] Failed to retrieve page {page} after {max_retries} attempts. Stopping pagination.")
            break
            
        soup = BeautifulSoup(response.text, "html.parser")
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
            
            # Extract category & brand chips
            chips = [s.get_text(strip=True) for s in card.select(".safir-card-chips span")]
            category = chips[0] if len(chips) > 0 else ""
            brand = chips[1] if len(chips) > 1 else ""
            
            # VAT rate from container data attribute
            vat_rate = card.get("data-vs-kdv", "20")
            
            # Retail / Regular Price
            reg_el = card.select_one(".safir-card-prices strong")
            reg_price = safe_float(reg_el.get("data-vs-price-amount")) if reg_el else None
            currency = reg_el.get("data-vs-price-currency", "USD") if reg_el else "USD"
            
            # Dealer Net Price (Excluding VAT)
            net_el = card.select_one(".safir-dealer-net strong")
            dealer_net = safe_float(net_el.get("data-vs-price-amount")) if net_el else None
            
            # Dealer Price (Including VAT)
            dealer_el = card.select_one(".safir-card-dealer:not(.safir-dealer-net) strong")
            dealer_price = safe_float(dealer_el.get("data-vs-price-amount")) if dealer_el else None
            
            # Discount rate tag (e.g. '%10 indirim')
            rate_el = card.select_one(".safir-dealer-rate")
            discount_rate = rate_el.get_text(strip=True) if rate_el else ""
            
            full_url = f"{BASE_URL}{link}" if link.startswith("/") else link
            
            products.append({
                "title": title,
                "category": category,
                "brand": brand,
                "regular_price": reg_price,
                "dealer_net_price": dealer_net,
                "dealer_price": dealer_price,
                "currency": currency,
                "vat_rate": vat_rate,
                "discount_rate": discount_rate,
                "url": full_url
            })
            page_products_count += 1
            
        log(f"  [Page {page:02d}] Extracted {page_products_count} items (Total: {len(products)})")
        page += 1
        polite_sleep(DELAY_BASE)

    elapsed = time.time() - start_time
    log(f"\n[Crawl] Crawling completed in {elapsed:.2f}s. Total items extracted: {len(products)}")

    # Create and format Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ürün Kataloğu"
    ws.views.sheetView[0].showGridLines = True
    
    headers = [
        "Ürün Adı", "Kategori", "Marka", "Satış Fiyatı", 
        "Bayi Fiyatı (KDV Hariç)", "Bayi Fiyatı (KDV Dahil)", 
        "Para Birimi", "KDV Oranı (%)", "İskonto Oranı", "Ürün Linki"
    ]
    ws.append(headers)
    
    # Header Styling
    header_fill = PatternFill(start_color="1E3D59", end_color="1E3D59", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
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
    
    # Append Rows
    for p in products:
        ws.append([
            p["title"], p["category"], p["brand"], p["regular_price"],
            p["dealer_net_price"], p["dealer_price"], p["currency"],
            p["vat_rate"], p["discount_rate"], p["url"]
        ])
        
    # Format Data Rows
    data_font = Font(name="Segoe UI", size=10)
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 20
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            # Number formatting for prices
            if col in [4, 5, 6] and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col in [7, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    # Adjust Column Widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
    ws.column_dimensions['A'].width = 45  # Ürün Adı
    ws.column_dimensions['J'].width = 50  # URL
    
    safe_save_excel(wb, OUTPUT_FILE)
    log(f"[Done] Complete catalogue saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    extract_full_catalogue()
