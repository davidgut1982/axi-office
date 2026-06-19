---
"@axi-office/outlook": patch
---

Outlook calendar: normalise bare dates to full ISO 8601 datetimes in cal-view.

- calViewCommand now accepts either full ISO 8601 datetimes (passed through unchanged) or bare YYYY-MM-DD dates (expanded: start → T00:00:00Z, end → T23:59:59Z).
- Invalid strings (neither a bare date nor an ISO datetime) throw an AxiError with a clear format hint.
- Updated help text and added tests for bare-date expansion, mixed input, and invalid input.
