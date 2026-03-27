# AI Coding Guidelines for Fiscal Automation Suite

## Project Overview
This is a multi-language suite of automation scripts for Brazilian fiscal and accounting tasks, including web scraping from tax authority portals (e-CAC, NFSE), document processing (PDF, Excel, Word), desktop UI automation, and report generation. Scripts are organized in folders by functionality, with a shared launcher interface for Python components. Includes both Python and C# implementations for performance-critical tasks.

## Architecture
- **Python Components**: Modular scripts in dedicated folders (e.g., `ECAC leitor/`, `auto_ata/`, `BAIXADOR NFSE NACIONAL/`), shared `menu_interface.py` launcher, tkinter-based UIs
- **C# Components**: Performance-optimized refactors of Python scripts using .NET, FlaUI for desktop automation, EPPlus/ClosedXML for Excel, QuestPDF for PDF generation
- **Data Flow**: Excel inputs → web automation/file processing → templated Excel/Word/PDF outputs
- **Structural Decisions**: Separate scripts for specific tasks, shared utilities, cross-language implementations for speed vs. flexibility trade-offs

## Key Patterns
- **Web Automation (Python)**: Use Playwright to connect to existing Chrome via CDP (`p.chromium.connect_over_cdp("http://127.0.0.1:9222")`). Handle anti-bot detection with regex patterns for denial/bot messages, automatic retries (up to 3 per doc), and 33-second delays.
- **Desktop UI Automation (C#)**: Use FlaUI with UIA3 for automating Windows applications (e.g., IOB accounting software). Implement --dump mode for UI element discovery, --run mode for execution.
- **Document Formatting**: Normalize CPF/CNPJ with `only_digits()` and `format_doc()`. Use Brazilian Portuguese accents in processing. Handle currency in centavos (integers) for precision.
- **Excel Handling**: Load templates with `openpyxl` (Python) or EPPlus/ClosedXML (C#), replace placeholders like `{{CNPJ}}`/`{{TOTAL}}`, copy row styles and merged cells. Output files named with timestamps.
- **Error Handling**: Custom exceptions (`SkipDoc`, `BotDetected`). Log to files and debug dumps (screenshots/HTML) in `_debug_*` folders.
- **PDF Generation (C#)**: Use QuestPDF with Community license for report generation from processed data.
- **Dependencies**: Python: `openpyxl`, `playwright`, `docx`, `pypdf2`. C#: EPPlus, ClosedXML, QuestPDF, FlaUI. Build with PyInstaller (Python) or dotnet (C#).

## Workflows
- **Python Build**: Use PyInstaller for executable creation (if .spec files exist, use `pyinstaller --onefile script.spec`; otherwise, run directly with `python script.py`). Execute via `menu_interface.py` launcher or directly. The launcher (`menu_interface.py`) provides a GUI to list scripts, run them, check/install dependencies automatically, and manage virtual environments.
- **C# Build**: Use `dotnet restore`, `dotnet build`, `dotnet run -- <args>` for compilation and execution.
- **Run**: For web scripts, ensure Chrome running on port 9222. For desktop automation, target application must be open or specify exe path.
- **Debug**: Check debug folders for screenshots/logs. Use `dump_debug()` (Python) or --dump flag (C#) for UI captures. Logs in `logs/` or `.launcher_logs/`.
- **Test**: Manual testing; validate outputs against templates. No automated tests present.
- **Dependency Management**: Python dependencies are auto-detected from imports and installed via pip in virtual environments. C# uses NuGet packages referenced in .csproj files.

## Conventions
- **Naming**: Portuguese variable names (e.g., `remetente`, `assunto`). File names with spaces and special chars. C# uses PascalCase for classes, camelCase for variables.
- **Imports**: Standard library first, then third-party. Use `from pathlib import Path`.
- **Code Style**: Functions for utilities, main() entry point. Inline comments in Portuguese. C# uses async/await sparingly, focuses on synchronous processing for reliability.
- **File Structure**: Input Excel in script folder, templates alongside, outputs in subfolders (e.g., `atas_geradas/`, `caixa_postal_ecac_*.xlsx`, `PDFs Prontos/`). C# projects use standard .NET structure with bin/obj folders.

## Integration Points
- **External Systems**: Brazilian tax portals (receita.fazenda.gov.br), Chrome browser, Windows desktop applications (IOB, etc.).
- **Cross-Language**: Python scripts often have C# equivalents for performance (e.g., `ajuste_diario_gfbr`).
- **APIs**: No external APIs; direct web scraping with selectors for dynamic content, UI automation for desktop apps.

Reference key files: `ECAC leitor/ECAC_LEITOR V2 (FINAL).py` (web scraping), `auto_ata/geraa_ata_auto_v3.py` (document generation), `menu_interface.py` (launcher), `FiscalAutomationSuite/README.md` (API automation suite), `../C#/ajuste_diario_gfbr_csharp/Program.cs` (C# Excel processing), `../C#/Downloader XLSX 721 + 720/Program.cs` (UI automation).