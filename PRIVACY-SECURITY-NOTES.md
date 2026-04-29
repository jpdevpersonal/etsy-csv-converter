# Privacy And Security Notes

This tool keeps CSV conversion local to the browser, but the live page will not be a zero-network page once Google Ads are enabled.

What the page can honestly claim:
- CSV contents are processed locally in the browser.
- CSV contents are not uploaded for conversion by this tool.
- The app does not use cookies, localStorage, sessionStorage, IndexedDB, fetch, XHR, or WebSocket APIs in the repo-local code.

What the page cannot claim once ads are enabled:
- "No data is sent anywhere."
- "Nothing is shared with any third party."
- "Completely private" without qualification.

Why:
- Google Ads introduces third-party script, iframe, and network activity.
- Google may receive standard page request, browser, device, IP, and cookie-related data even though CSV contents remain local.

Deployment guidance:
- Keep the page on HTTPS only.
- Keep CSV-processing language specific: say the CSV stays local and is not uploaded.
- Use the headers in [_headers.example](_headers.example) or transpose those exact values into your CDN/server config.
- Re-test the final CSP in the browser after adding the actual ad snippet, because Google Ads can add domains over time.

Link guidance:
- External links in the page use `referrerpolicy="no-referrer"` and `rel="noreferrer"` to avoid leaking the tool URL when navigating out.
- If you want referral analytics between this page and `simplebiztoolkit.com`, remove the per-link no-referrer settings intentionally rather than by accident.

Hardening still worth doing later:
- Move inline JSON-LD and inline `style` attributes out of the HTML so the CSP can drop `'unsafe-inline'`.
- If you know the final host, convert the example headers into host-native config instead of relying on documentation alone.