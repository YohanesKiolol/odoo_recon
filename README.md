# Odoo ↔ Bank Reconciliation Script

Compares bank transactions (Excel) against Odoo exported transactions (Excel) by amount. Outputs a color-coded Excel report.

## Setup

```bash
cd odo_automation

# Create virtual environment (one-time)
python3 -m venv .venv

# Activate it (do this every time before running)
source .venv/bin/activate

# Install dependencies (one-time)
pip install -r requirements.txt

# Copy and fill in your config
cp .env.example .env
# Edit .env with your column names and file paths
```

## Usage

1. Export your Odoo transactions to Excel (Accounting → Journal Items → Export)
2. Place both files in `input/`
3. Set the column names in `.env`
4. Run:

```bash
python main.py
```

Or override file paths directly:

```bash
python main.py --bank input/my_bank.xlsx --odo input/my_odoo.xlsx
```

## Output

`output/reconciliation_YYYYMMDD_HHMMSS.xlsx` with 3 sheets:

| Sheet | Content |
|-------|---------|
| **Semua Transaksi** | All transactions with status |
| **Selisih** | Discrepancies only |
| **Legenda** | Status color legend |

### Status Labels

| Status | Meaning |
|--------|---------|
| `Done` | Amount found in both Bank and Odoo ✅ |
| `Cuma ada di Bank` | Amount only in Bank, not in Odoo ⚠️ |
| `Cuma ada di ODO` | Amount only in Odoo, not in Bank ⚠️ |

## Config (.env)

| Key | Description |
|-----|-------------|
| `BANK_EXCEL_PATH` | Path to bank Excel file |
| `BANK_AMOUNT_COLUMN` | Column name for amount in bank Excel |
| `BANK_DATE_COLUMN` | (Optional) Date column — display only |
| `BANK_DESC_COLUMN` | (Optional) Description column — display only |
| `ODO_EXCEL_PATH` | Path to Odoo Excel file |
| `ODO_AMOUNT_COLUMN` | Column name for amount in Odoo Excel |
| `ODO_DATE_COLUMN` | (Optional) Date column — display only |
| `ODO_DESC_COLUMN` | (Optional) Description column — display only |
| `DECIMAL_SEPARATOR` | `auto` (recommended), `comma`, or `dot` |
| `OUTPUT_DIR` | Output folder (default: `output`) |

## Amount Format Handling

Script auto-detects and handles:
- `1.500.000` — Indonesian thousands (dot)
- `1.500.000,50` — Indonesian with decimal comma
- `1,500,000` — US thousands (comma)
- `1,500,000.50` — US with decimal dot
- `Rp 1.500.000` — with currency prefix
- `1500000` — plain number

## Test

```bash
python test_amount_utils.py
```

## File Structure

```
odo_automation/
├── main.py              # Entry point
├── config.py            # Reads .env
├── excel_reader.py      # Read input Excels
├── reconciler.py        # Comparison logic (multiset)
├── excel_writer.py      # Write output Excel
├── amount_utils.py      # Amount parsing/normalization
├── test_amount_utils.py # Self-test for parser
├── requirements.txt
├── .env.example         # Template — copy to .env
├── input/               # Put your Excel files here
└── output/              # Reports written here
```
