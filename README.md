# Semrush-style Site Audit Dashboard (starter)

## Run on Windows

1. Extract this ZIP into a new folder.
2. Open that folder in File Explorer.
3. Click the address bar, type `cmd`, press Enter.
4. Run:

   `py -m pip install -r requirements.txt`

5. Then run:

   `py -m streamlit run app.py`

6. If a browser tab does not open, visit `http://localhost:8501`.
7. Enter a site you own or are authorized to audit and click **Start audit**.
8. Use **Refresh audit** to refresh the log/report and download the Excel file.

## Excel output

The workbook has a single worksheet named `1 - All` and preserves the 75 column headers and order from the example workbook you supplied.

Some fields are exact measurements; others are estimates or placeholders. Pixel width is approximated, transfer size is currently approximated from downloaded response body size, carbon values and spelling/grammar counts are blank, and JS-link metrics are not collected. Validate these before relying on them for client deliverables.

## Important limitations

- This is a local prototype, not a complete Semrush clone.
- It does not render JavaScript.
- It does not separately request every discovered link to verify its status.
- Redirect handling, crawl depth, duplicate detection and some metrics need further testing before production use.
- It respects robots.txt and uses a minimum request delay of 0.5 seconds.
- Use only on websites you own or have permission to crawl.
- Do not expose this app publicly without authentication, quotas, robust SSRF protection and deployment hardening.
