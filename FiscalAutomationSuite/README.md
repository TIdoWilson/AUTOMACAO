# Fiscal Automation Suite

A comprehensive suite for automating Brazilian fiscal and accounting tasks, including parcelamento processing, PDF emission, and web scraping from tax authority portals.

## Features

- **Authentication**: OAuth2 authentication with Serpro Integra Contador API
- **Parcelamento Consultation**: Query available parcels for Simples Nacional
- **PDF Emission**: Emit and download PDF guides for Simples Nacional
- **PGFN Processing**: Automate PGFN parcelamento downloads via Sispar website
- **Excel Processing**: Read and process parcelamento lists from Excel files

## Requirements

- Python 3.8+
- Dependencies: requests, requests-pkcs12, pandas, openpyxl, python-dotenv, selenium (for PGFN)

## Installation

1. Install dependencies:
   ```bash
   pip install requests requests-pkcs12 pandas openpyxl python-dotenv selenium
   ```

2. Configure `.env` file with required credentials:
   - SERPRO_CONSUMER_KEY
   - SERPRO_CONSUMER_SECRET
   - SERPRO_PFX_PATH
   - SERPRO_PFX_PASSWORD
   - CNPJ_CONTRATANTE

## Usage

### Modes

- `python fiscal_automation.py` - Test authentication
- `python fiscal_automation.py consultar` - Consult available parcels
- `python fiscal_automation.py emitir` - Emit PDFs for parcels
- `python fiscal_automation.py pgfn` - Process PGFN via Sispar
- `python fiscal_automation.py processar` - Process Excel list for downloads

### Files

- `LISTA PARCELAMENTOS.xlsx` - Input Excel with parcelamento data
- `PDFs/` - Output directory for downloaded PDFs
- `arquivos/api_teste/` - API test outputs
- `fiscal_automation.log` - Main log file
- `charges.log` - Billing log

## Notes

- PGFN processing requires Selenium and Chrome WebDriver
- Simples Nacional operations use Serpro Integra Contador API
- Ensure all environment variables are set correctly

## License

Proprietary - Internal use only