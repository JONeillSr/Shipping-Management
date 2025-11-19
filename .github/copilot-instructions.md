This repository contains a Windows-focused PowerShell-based Logistics Automation Suite (email generation, PDF invoice parsing, recipient management, and quote tracking).

Keep guidance short and actionable with pointers to concrete files and patterns in this repo.

1) Big picture
- Primary components (see `logistics_suite_readme.md`):
  - `Logistics-Automation-Suite.ps1` — GUI launcher and orchestration hub.
  - `Logistics-Config-GUI.ps1` — configuration GUI and template manager; calls the parser with `Invoke-PDFInvoiceParser`.
  - `Generic-PDF-Invoice-Parser.ps1` — the universal, vendor-aware PDF→structured-data parser used by the GUI and wrappers.
  - `Generate-LogisticsEmail.ps1` — consumes CSV + JSON config to create HTML emails and PDFs (depends on `Convert-HTMLtoPDF.ps1`).
  - `Freight-Recipient-Manager.ps1` — manages `Data/FreightRecipients.json` and provides recipient selection for emails.

2) How code is executed and typical workflows
- Scripts are run directly in Windows PowerShell (PowerShell 5.1+). Typical developer/user commands:
  - Create starter templates: `.\\Create-StarterTemplates.ps1`
  - Launch suite: `.\Logistics-Automation-Suite.ps1`
  - Parse a PDF from the GUI: `Logistics-Config-GUI.ps1` → Import from PDF button calls `Generic-PDF-Invoice-Parser.ps1 -ReturnObject`.
  - Generate emails from CSV: `.\Generate-LogisticsEmail.ps1 -CSVPath .\auction.csv -ImageDirectory .\Images -ConfigPath .\Templates\MyTemplate.json`

3) Important integration points & external dependencies
- PDF text extraction: `Generic-PDF-Invoice-Parser.ps1` prefers `pdftotext` (xpdf-tools) and falls back to iTextSharp (NuGet). If neither is available the script warns. Mention these in changes that affect parsing.
- PDF conversion: `Convert-HTMLtoPDF.ps1` (helper) supports Foxit/Edge/Chrome; `Generate-LogisticsEmail.ps1` loads it if present. Avoid changing HTML here-strings without testing conversion.
- PowerShell modules: `Generate-LogisticsEmail.ps1` requires `PSWritePDF` and `ImportExcel` (scripts install them if missing). Keep module usage minimal and prefer the existing checks.
- Data files: `Data/FreightRecipients.json`, `Data/AuctionQuotes.json`, and `Templates/*.json`. Scripts expect these JSON files and create directories if missing — preserve shapes when editing.

4) Patterns & conventions to follow
- Windows/PowerShell-first: GUI uses System.Windows.Forms; expect synchronous Windows UI flows. Don't convert to cross-platform frameworks unless you update all launchers.
- Parameter style: scripts use advanced functions with CmdletBinding and named parameters (e.g., `-PDFPath`, `-ReturnObject`). Follow the same style for new scripts.
- JSON templates/configs: use `ConvertTo-Json -Depth 10` and read with `ConvertFrom-Json`. Preserve property names and nesting (see `Get-AuctionConfig` and `Save-TemplateFile`).
- Logging: `Generate-LogisticsEmail.ps1` uses `Write-JTLSLog`. If you add logs, use this function or follow its timestamped format.
- Error handling: functions catch exceptions and return $null or write warnings; GUI layers show friendly messages. Keep user-facing messages concise and color-coded where appropriate.

5) Small, concrete examples to reference
- To call the parser programmatically and get structured data:
  & `Generic-PDF-Invoice-Parser.ps1 -PDFPath 'C:\Invoices\inv.pdf' -ReturnObject`
- To save a new template from the GUI code path:
  The GUI calls `Save-TemplateFile -Config $config -TemplateName 'Brolyn_Oct09_2025'` which writes JSON to `./Templates/`.
- To ensure PDF conversion works in CI/manual runs: test `Convert-HTMLtoPDF.ps1` with an HTML file and verify expected output in `.\Output\`.

6) What to watch for when modifying code
- HTML generation uses PowerShell here-strings heavily. Small changes (escaping `<`, `@`, `"`) can break PDF conversion or Outlook attachment behavior — run full manual end-to-end tests (generate HTML -> convert PDF -> open Outlook draft).
- OneDrive paths and usernames with apostrophes are handled carefully in the launcher (see recent changelog). Preserve quoting logic in helper calls.
- `Generic-PDF-Invoice-Parser.ps1` contains vendor-specific regexes and learned patterns (`Data/InvoicePatterns.json`); if you add vendors, add patterns there and ensure `Find-InvoiceVendor` picks them up.

7) Quick debugging tips for contributors
- Reproduce UI workflows manually (launch `Logistics-Automation-Suite.ps1`) to exercise integration points.
- Use `-ReturnObject` on the parser to get a structured object for unit tests: `.\Generic-PDF-Invoice-Parser.ps1 -PDFPath .\test.pdf -ReturnObject | ConvertTo-Json -Depth 10`
- Check logs in `.\Logs\` (script-created file names: `LogisticsEmail_*.log`).
- For parsing issues, confirm `pdftotext` is installed or test with iTextSharp path under `%USERPROFILE%\.nuget\packages\itextsharp`.

8) Files to reference while coding
- `logistics_suite_readme.md` — user-facing architecture and workflows
- `Generic-PDF-Invoice-Parser.ps1` — vendor parsing logic and patterns
- `Logistics-Config-GUI.ps1` — template manager and parser integration
- `Generate-LogisticsEmail.ps1` — HTML/PDF generation, modules, and logging
- `Freight-Recipient-Manager.ps1` — data shape for recipients

If anything here is unclear or you want more detail (examples, test snippets, or additions), tell me which area to expand and I will iterate.
