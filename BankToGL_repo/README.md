# BankToGL

Turns a monthly bank statement into a draft D365 Business Central journal import
file (Gen. Journal Line layout), using an editable rules file so no code
changes are needed month to month.

## Files

| File | What it is |
|---|---|
| `app.py` | The screen you see in the browser (built with Streamlit) |
| `bank_to_gl.py` | The engine: reads the statement, applies the rules, writes the import file |
| `bc_template.py` | The exact 176-column Business Central header, captured from a real export |
| `GL_Rules_template.xlsx` | Starting point for your rules/settings file - copy it to `GL_Rules.xlsx` and fill in your own account numbers |
| `requirements.txt` | The Python packages the app needs |
| `START_APP.bat` | Double-click to install requirements and launch the app (Windows) |

`GL_Rules.xlsx` itself, and any statement or output file, are **not** committed
(see `.gitignore`) - they hold real account numbers and transaction data.

## Run it locally

```
pip install -r requirements.txt
cp GL_Rules_template.xlsx GL_Rules.xlsx   # then edit GL_Rules.xlsx with your own settings
streamlit run app.py
```

## How the pieces fit together

```
app.py  ───calls (subprocess)──▶  bank_to_gl.py
                                       │
                                       ├─ reads the bank statement (xlsx/csv)
                                       ├─ reads GL_Rules.xlsx:
                                       │     Rules            - keyword → account
                                       │     Payroll Schedule - employee × type → account, amount filled monthly
                                       │     Chart of Accounts - valid account list
                                       │     Settings          - bank account code, clearing accounts, etc.
                                       ├─ imports bc_template.py for the output header
                                       └─ writes:
                                             GL_Import_..._DRAFT.xlsx   (import this into D365)
                                             Review_....xlsx            (checks + flagged lines)
```

`app.py` never talks to `GL_Rules.xlsx` or the statement directly - it only
uploads files, shells out to `bank_to_gl.py`, and shows the two output files.
All the logic lives in `bank_to_gl.py`.
