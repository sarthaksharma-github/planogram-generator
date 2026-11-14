# HD Planogram Generator

A high-performance Streamlit web application styled with the ChaiUI design system that generates bay-level planogram layouts, provides automated compliance auditing, and generates store SKU mapping sheets from Excel workbooks.

## Features

- **ChaiUI design system**: Restrained dark theme (#0a0a0a), warm hairline borders, cream highlight accents (#ffedd5), Manrope and Montserrat typography, zero box-shadows, and asymmetric corner action buttons.
- **Animated loading screen**: Smooth 2.4-second fade-in and fade-out transition displaying the application name on initial session launch.
- **Automatic POG format detection**: Handles multiplier shorthand, explicit lists, single repeated sizes, and trailing description text.
- **LFT bays placed first**, followed by standard POG bays.
- **Robust column detection**: Case-insensitive column matching with positional fallbacks.
- **Facing clamp safety**: Clamps facing values to 1 or 2 to prevent expansion issues.
- **Automated audit verification**: Validates end-position placement rules (for example, Baja and Vigo) with compliance rate analytics.
- **Store SKU generator**: Generates store SKU mapping sheets for active board sets.

## Supported POG Formats

| Input | Result |
|-------|--------|
| `12 Bay - 99` | 99 x 12 |
| `12 Bay - 99 - Floor Tile Test - Str 1602 - June 26` | 99 x 12 (description stripped) |
| `7 BAY - 99,99,99,99,99,99,87` | Explicit list |
| `10 Bay - 99,...,87 - Floor Tile Test` | Explicit list, last description stripped |
| `10 Bay - 8-99,2-51` | 99x8 + 51x2 (multiplier shorthand) |

## Workflows

1. **Generate planogram and run audit**: Full pipeline allocating bays into `Generated Planogram`, logging issues into `Validation`, running compliance audit into `Audit Report`, and downloading the complete multi-sheet workbook.
2. **Audit existing planogram**: Standalone audit mode for existing planogram sheets with CSV export.
3. **Generate store SKU sheet**: Standalone generator mapping SO Board Sets and SKUs into a new `Store SKU` sheet.

## Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Running Tests

```bash
python -m pytest tests/test_planogram.py -v
```
