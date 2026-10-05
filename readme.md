# pdf2fidavista

Converts a Swedbank (Latvia) account statement PDF ("Konta Izraksts") into a FiDAViSta XML
statement, in the same shape as Swedbank's own FiDAViSta export (version 1.01, WINDOWS-1257
encoded).

## Usage

```
pip install pdfplumber

# Inspect the PDF text (one line per printed row)
python pdf2fidavista.py statement.pdf --dump-text

# Convert (bank, client, IBAN, period and balances are read from the PDF)
python pdf2fidavista.py statement.pdf -o statement.xml
```

Anything the script can't find in the PDF can be passed on the command line: `--iban`, `--ccy`,
`--open-bal`, `--close-bal`, `--start`, `--end`, `--bank-name`, `--bank-id`, `--client-name`,
`--client-id`, `--from`. Use `--version 1.2` to write the FiDAViSta 1.2 namespace instead of 1.01.

## How it works

Swedbank prints each transaction as a block of rows (date, document number, counterparty, account;
then archive number, operation code, BIC and amount; then payment details). Debit vs credit is only
visible from which column the amount sits in, so the parser works on word coordinates
(`pdfplumber.extract_words`) rather than plain text. Column boundaries are taken from the table
header on each page, with fallbacks in the CONFIG section of the script.

Swedbank operation codes are mapped to FiDAViSta type codes (`INB` → `INP`, `IZP`/`PRV` → `OUTP`,
everything else → `OTHR`), and the archive number is used as `BankRef`.

After converting, the script checks the credit and debit turnovers and that
opening balance + credits − debits = closing balance against the totals printed in the PDF.
A mismatch means some rows weren't parsed correctly; run `--dump-text` and compare.

## Limitations

- **Scanned PDFs** contain no text; run OCR first (e.g. `ocrmypdf`).
- **Other banks** use different layouts; the parser targets Swedbank's statement format only.
- Validating the output against the official XSD from Finance Latvia is a good idea.
