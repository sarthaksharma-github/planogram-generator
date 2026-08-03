# Home Depot Planogram Generator

A Streamlit web application that generates complete bay-level planogram layouts for every store from an Excel workbook.

## Features

- **Automatic POG format detection** — handles 3 format variants and trailing description text
- **LFT bays placed first**, followed by POG bays
- **Robust column detection** — case-insensitive, with positional fallback
- **Duplicate "Facings" column support** — correctly identifies Stock vs Display facing columns
- **Invalid facing values clamped to 1** — prevents expansion bugs from CF or other columns
- **Notes rules** — stable-partition logic to move specific SKUs to the end
- **Validation sheet** — all issues logged to Excel output without stopping execution

## Supported POG Formats

| Input | Result |
|-------|--------|
| `12 Bay - 99` | 99 × 12 |
| `12 Bay - 99 - Floor Tile Test - Str 1602 - June 26` | 99 × 12 (description stripped) |
| `7 BAY - 99,99,99,99,99,99,87` | explicit list |
| `10 Bay - 99,...,87 - Floor Tile Test` | explicit list, last description stripped |
| `10 Bay - 8-99,2-51` | 99×8 + 51×2 (multiplier shorthand) |

## Input Workbook Structure

### Sheet: `Store List`
| Store | Current Store POG | Current LFT | Notes |
|-------|-------------------|-------------|-------|
| 6349 | 8 Bay - 99 - Floor Tile... | 1 Bay - 51 - LFT... | Normal flow |

### Sheet: `Stock SKUs and Displays`
| Store | Stock SKU | Stock Description | ... | Facings | Display SKU | Display Description | ... | Facings | CF |
|-------|-----------|-------------------|-----|---------|-------------|---------------------|-----|---------|-----|

### Sheet: `Special Order Boards`
| Store | SO SKU | Description | ... | Facings | ... | CF |
|-------|--------|-------------|-----|---------|-----|-----|

## Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Cloud

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set **Main file path** to `app.py`
5. Click **Deploy**

## Project Structure

```
planogram_app/
├── app.py                  # Streamlit UI
├── requirements.txt
├── .gitignore
├── README.md
└── planogram/
    ├── __init__.py
    ├── config.py           # BAY_RULES, NOTES_RULES, constants
    ├── models.py           # SKURecord, BayRule dataclasses
    ├── logger.py           # PlanogramLogger (deduplicating)
    ├── parser.py           # POG string parser + self-tests
    ├── loader.py           # Excel loading, column detection
    ├── allocator.py        # Bay allocation + orchestrator
    └── writer.py           # Excel output (BytesIO)
```

## Extending

### Add a new bay type
Edit `planogram/config.py`, add one row to `BAY_RULES`:
```python
BAY_RULES["63"] = BayRule(display=5, so=1, stock=5)
```

### Add a new Notes rule
Edit `planogram/config.py`, append one dict to `NOTES_RULES`:
```python
{"trigger": "move calacatta to end", "keywords": ["calacatta"]}
```
