# Stock & Price Controller

A bespoke B2B data synchronization and automation tool designed to extract, monitor, and normalize wholesale dealer pricing, inventory levels, and product catalogues from supplier portals.

> **Project Note**: This repository serves as a personal codebase for version control, continuous iteration, and a showcase of custom data scraping, automation, and desktop GUI engineering.

---

## Technical Overview & Capabilities

This project was built to replace manual stock and price checking workflows by providing automated, high-throughput extraction pipelines integrated with styled Excel reporting.

### Key Architectural Highlights

- **Session Management & Token Caching**: Implements an automated authentication handshake that caches active Bearer tokens locally, validating token health prior to requests to eliminate redundant login roundtrips.
- **High-Throughput SSR Extraction**: Traverses server-side rendered (SSR) catalogue listings and product microdata (`data-vs-*`, `data-cart-product` attributes) using `BeautifulSoup` and `requests`, processing hundreds of products in seconds without the overhead of headless browser automation.
- **Dynamic Multi-Currency Normalization**: Extracts live exchange rates directly from portal root markup (`data-currency-rate-*`), calculating accurate real-time currency conversions across USD, EUR, and TRY.
- **Dual Interface Design (Desktop GUI & CLI)**:
  - **Desktop GUI Application (`crawler_gui.py`)**: Built with Tkinter, featuring real-time streaming execution logs, file pickers, thread-safe asynchronous worker execution, and status indicators.
  - **Headless CLI Pipelines (`crawler.py`, `ultimate_crawler.py`)**: Lightweight scripts designed for rapid execution, targeted batch matching against input spreadsheets, or full-catalogue export routines.
- **Resilient Excel Reporting**: Custom `openpyxl` integration that enforces corporate styling, price number formatting, auto-fitting column dimensions, and conflict handling for locked files.
- **Defensive Crawling Safeguards**: Incorporates randomized delay jitter, modern browser fingerprint headers, and exponential backoff retry logic.
