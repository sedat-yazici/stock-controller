# Stock & Price Controller

An automated stock and price tracking crawler for a supplier portal. It fetches up-to-date dealer pricing, retail pricing, VAT rates, and inventory status, formatting results directly into styled Excel spreadsheets.

---

## Features

- **Automated Authentication & Session Caching**: Securely logs in via B2B API and reuses cached session tokens.
- **Full Catalogue Crawling**: Traverses paginated catalogue pages and extracts all products, prices, categories, and brands.
- **Search & Targeted Extraction**: Matches product codes from an input Excel file against the portal and updates prices/stock quantities.
- **GUI & CLI Modes**: Supports both a desktop GUI interface (`crawler_gui.py`) and command-line execution (`crawler.py`).
- **Currency Normalization**: Extracts live exchange rates directly from portal metadata for multi-currency conversion (USD / TRY / EUR).
- **Styled Excel Output**: Generates cleanly styled `.xlsx` reports with grid lines, color-coded headers, and auto-adjusted columns.

---

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/sedat-yazici/stock-controller.git
   cd stock-controller
   ```

2. **Install dependencies:**
   ```bash
   pip install requests beautifulsoup4 openpyxl python-dotenv
   ```

---

## Configuration

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Open `.env` and fill in your portal configuration and credentials:
   ```ini
   BASE_URL=https://your-supplier-portal.com
   SERVICE_NAME=your-supplier-portal.com
   LOGIN_URL=https://your-supplier-portal.com/api/magaza/uyelik/giris?hizmetAdi=your-supplier-portal.com
   ACCOUNT_URL=https://your-supplier-portal.com/api/magaza/uyelik/hesabim?hizmetAdi=your-supplier-portal.com
   USERNAME=your_username
   PASSWORD=your_password
   INPUT_EXCEL_PATH=input.xlsx
   OUTPUT_EXCEL_PATH=results.xlsx
   ```

> **Note:** The `.env` file and `session_cache.json` are excluded from version control via `.gitignore`. Never commit credentials.

---

## Usage

### Desktop GUI
Launch the visual control interface:
```bash
python crawler_gui.py
```

### Targeted Product Code Matching (CLI)
Process an input Excel file (`input.xlsx`) and output results:
```bash
python crawler.py
```

### Full Catalogue Extractor
Crawl and extract the entire portal catalogue to Excel:
```bash
python ultimate_crawler.py
```

---

## License

MIT License. For authorized B2B portal integration and data management only.
