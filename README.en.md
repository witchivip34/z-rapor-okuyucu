# z-rapor-okuyucu (English summary)

Turkish cash registers print a **Z report** at the end of each day: department totals, VAT breakdown,
payment split (cash/card), voids and running (cumulative) totals. Accounting offices receive these on
paper every month and type them in by hand.

This tool:

1. **Reads** photos of Z reports with a vision-language model (Anthropic Messages API by default, or the
   `claude -p` CLI). Each photo is tiled into 9 enlarged overlapping crops and read **twice
   independently**; disagreements become warnings. The model is told never to guess (unreadable = `null`).
2. **Checks** every figure with 7 arithmetic/consistency rules: department sum = net sales = fiscal total =
   VAT breakdown; payments = net sales; VAT formula per rate; bank breakdown = card; counts; the
   **cumulative chain** (today's running total − yesterday's = today's sales), which also lets the code
   *prove and fix* a misread digit; Z-number/date continuity.
3. **Writes** a colour-coded Excel workbook (red = doesn't add up, yellow = unusual/estimated, blue =
   corrected from the chain) and import files for the Luca accounting software (balance-sheet journal
   template and simplified-ledger CSV). *Luca import has not yet been verified against a live system.*

Also included: an expense-receipt list generator and a password-protected mobile upload page (FastAPI).

```bash
pip install -e ".[web,test]"
z-rapor dogrula ornekler/okuma_ornek.json -o calisma/ornek.xlsx   # no API key needed
export ANTHROPIC_API_KEY=...                                       # or put it in .env
z-rapor oku ornekler/z_ornek.png -o calisma/demo.xlsx
python -m pytest -q
```

All sample data is synthetic ("ÖRNEK LOKANTA", tax ID 1234567890). The user interface, messages and
main documentation are in Turkish: see [README.md](README.md). License: MIT © 2026 Ahmet Said Kara.
