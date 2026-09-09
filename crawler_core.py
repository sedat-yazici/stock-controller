import os
import time
import json
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment configuration
load_dotenv(dotenv_path=os.path.join(os.getcwd(), ".env"), override=True)

def get_config():
    base_url = os.getenv("BASE_URL", "").rstrip("/")
    service_name = os.getenv("SERVICE_NAME", "")
    account_url = os.getenv("ACCOUNT_URL", f"{base_url}/api/magaza/uyelik/hesabim?hizmetAdi={service_name}")
    login_url = os.getenv("LOGIN_URL", f"{base_url}/api/magaza/uyelik/giris?hizmetAdi={service_name}")
    return {
        "base_url": base_url,
        "service_name": service_name,
        "account_url": account_url,
        "login_url": login_url
    }

def log(msg, callback=None):
    if callback:
        callback(msg)
    else:
        print(msg)

def convert_and_format_price(amount_str, source_curr, target_curr="USD", rates=None):
    """Converts price from source currency to target currency using site rates and formats as currency text."""
    if rates is None:
        rates = {"TL": 1.0, "KARMA": 1.0, "USD": 1.0}
    source = str(source_curr or "TL").strip().upper()
    if source == "TRY": source = "TL"
    if source == "EURO": source = "EUR"
    target = str(target_curr or "USD").strip().upper()
    
    try:
        raw = str(amount_str).strip()
        if "," in raw and "." in raw:
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", ".")
        num = float(raw)
    except (ValueError, TypeError):
        return None
        
    if target == "KARMA" or source == target:
        converted = num
    else:
        source_rate = rates.get(source, 1.0)
        target_rate = rates.get(target, 1.0)
        if target_rate <= 0:
            target_rate = 1.0
        converted = (num * source_rate) / target_rate
        
    # Format according to site standards (e.g. 44,40 USD or 1.340,00 USD)
    formatted = f"{converted:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{formatted}"

def clean_price(val):
    """Formats and normalizes price as a clean currency string (e.g. '44,40 USD')."""
    if val is None:
        return None
    val_str = " ".join(str(val).split()).strip()
    if val_str.lower() in ["bulunamadı", "bulunamadi", "none", "nan", ""]:
        return "Bulunamadı"
    return val_str

def check_cached_session(session, cache_file, log_callback=None):
    """Checks if a valid session exists in cache."""
    if not os.path.exists(cache_file):
        return False
    try:
        with open(cache_file, "r", encoding="utf-8") as f:
            cache = json.load(f)
        token = cache.get("token")
        if not token:
            return False
            
        session.headers.update({
            "Authorization": f"Bearer {token}",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        })
        
        config = get_config()
        res = session.get(config["account_url"], timeout=10)
        if res.status_code == 200:
            log("Reusing existing login session from cache...", log_callback)
            return True
        else:
            log("Cached session is invalid or expired. Re-authenticating...", log_callback)
            session.headers.pop("Authorization", None)
            return False
    except Exception as e:
        log(f"Failed to load cached session: {e}", log_callback)
        session.headers.pop("Authorization", None)
        return False

def save_session_cache(token, cache_file, log_callback=None):
    """Saves session token to cache."""
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump({"token": token, "saved_at": time.time()}, f)
    except Exception as e:
        log(f"Failed to write cache: {e}", log_callback)

def login(session, username, password, cache_file, log_callback=None):
    """Authenticates to the Supplier Portal API and updates session headers with Bearer token."""
    if check_cached_session(session, cache_file, log_callback):
        return True

    config = get_config()
    payload = {
        "hizmetAdi": config["service_name"],
        "kullanici": username,
        "sifre": password
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    
    log("Logging in to supplier portal...", log_callback)
    try:
        res = session.post(config["login_url"], json=payload, headers=headers, timeout=15)
        if res.status_code != 200:
            log(f"Login failed with status code {res.status_code}", log_callback)
            return False
        
        token = res.json().get("token")
        if not token:
            log("Login response did not contain an authentication token.", log_callback)
            return False
        
        save_session_cache(token, cache_file, log_callback)
        
        session.headers.update({
            "Authorization": f"Bearer {token}",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        })
        log("Login successful.", log_callback)
        return True
    except Exception as e:
        log(f"Error during login: {e}", log_callback)
        return False

def search_and_extract(session, product_code, log_callback=None):
    """Searches for a product and extracts its dealer price and stock amount."""
    config = get_config()
    base_url = config["base_url"]
    service_name = config["service_name"]
    search_url = f"{base_url}/urunler?hizmetAdi={service_name}&dilKodu=tr&arama={product_code}"
    
    try:
        response = session.get(search_url, timeout=15)
        if response.status_code != 200:
            log(f"  Search request failed with status: {response.status_code}", log_callback)
            return None, None
        
        soup = BeautifulSoup(response.text, "html.parser")
        card = soup.find("article", class_="safir-product-card")
        if not card:
            log("  No product card found.", log_callback)
            return "Bulunamadı", 0
        
        # Get details page link
        title_link = card.find("h3", class_="safir-card-title").find("a")
        if not title_link or not title_link.get("href"):
            log("  Could not find product details link.", log_callback)
            return None, None
        
        href = title_link["href"]
        detail_url = f"{base_url}{href}" if href.startswith("/") else href
        detail_response = session.get(detail_url, timeout=15)
        if detail_response.status_code != 200:
            log(f"  Failed to load detail page (status: {detail_response.status_code})", log_callback)
            return None, None
            
        detail_soup = BeautifulSoup(detail_response.text, "html.parser")
        cart_product = detail_soup.find(attrs={"data-cart-product": True})
        if not cart_product:
            log("  Could not find data-cart-product metadata block on detail page.", log_callback)
            return None, None
            
        # 1. Change the currency to USD by updating the <span data-currency-current=""> element
        for curr_span in detail_soup.find_all(attrs={"data-currency-current": True}):
            curr_span.string = "USD"
            
        # 2. Extract site conversion rates from document root dataset
        rates = {"TL": 1.0, "KARMA": 1.0, "USD": 1.0}
        html_tag = detail_soup.find("html")
        if html_tag:
            for attr, val in html_tag.attrs.items():
                if attr.startswith("data-currency-rate-"):
                    c_code = attr.replace("data-currency-rate-", "").upper()
                    try:
                        rates[c_code] = float(str(val).replace(",", "."))
                    except ValueError:
                        pass

        price = None
        stock = None         
        # 3. Transform all [data-vs-price] elements in DOM to USD
        strong = detail_soup.find("strong", attrs={"data-product-dealer-price-label": True})
        if strong:
            raw_amount = strong.get("data-vs-price-amount")
            raw_curr = strong.get("data-vs-price-currency") or "TL"
            if raw_amount:
                price = convert_and_format_price(raw_amount, raw_curr, "USD", rates)        
        if price is None:
            price = "Bulunamadı"
            
        stock = cart_product.get("data-product-stock")
        price = clean_price(price)
        if stock is not None:
            try:
                stock = int(stock)
            except ValueError:
                pass
                
        return price, stock
        
    except Exception as e:
        log(f"  Exception during search & extract: {e}", log_callback)
        return None, None
