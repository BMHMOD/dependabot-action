# VBA Macro: Fill Yard Files from STOCK

## Overview
This VBA macro automatically fills **Internal Yard** and **External Yard** Excel files with container data from the **STOCK.xlsx** file based on:
- Container Mode: Import, Export, Storage
- Block: M, A, B, C, D, H, F, Y777, S22, etc.
- Full/Empty status (FE): F or E
- Container Length: 20, 40, 45 feet

## Files Included
1. **FillYardsFromStock.vba** - Basic version (simple, easy to understand)
2. **FillYardsFromStock_Enhanced.vba** - Enhanced version ⭐ **RECOMMENDED**
   - Faster processing
   - Progress bar
   - Better error handling
   - Optimized for large files
3. **تعليمات_الاستخدام.md** - Arabic instructions (detailed)
4. **README.md** - English instructions (this file)

## Quick Start

### Step 1: Import the Code
1. Open a new Excel workbook
2. Press `Alt + F11` to open VBA Editor
3. Click `Insert` → `Module`
4. Copy all code from `FillYardsFromStock_Enhanced.vba`
5. Paste into the white window
6. Save as **Excel Macro-Enabled Workbook** (*.xlsm)

### Step 2: Run the Macro
1. Press `Alt + F8` to open Macro list
2. Select `FillYardsFromStock_Enhanced`
3. Click `Run`
4. Select your three files when prompted:
   - STOCK.xlsx
   - internal yard.xlsx
   - external yard.xlsx
5. Wait for processing to complete

## Customization

### Adding Internal Yard Blocks
In the code, find this section:
```vba
dictInternal.Add "M", 6
dictInternal.Add "A", 9
```
- First value: Block name from STOCK file
- Second value: Row number in Internal Yard file where this block starts

To add a new block:
```vba
dictInternal.Add "NEWBLOCK", 56
```

### Adding External Yard Areas
Find this section:
```vba
dictExternal.Add "التجارية", Array(6, "S444|S068|S032")
```
- First value: Yard name (for reference)
- Second value (Array):
  - First element: Starting row number
  - Second element: List of Areas/Blocks (separated by |)

To add a new yard:
```vba
dictExternal.Add "NewYard", Array(16, "S100|S200|AREA3")
```

## Expected File Structure

### STOCK.xlsx
Must contain these columns:
- **Column F (6)**: Area
- **Column G (7)**: Block
- **Column J (10)**: Container Length (20, 40, 45)
- **Column M (13)**: FE (F=Full, E=Empty)
- **Column P (16)**: Mode (Import, Export, Storage)

### Internal Yard Structure
Each block has 3 rows:
- Row 1: Import
- Row 2: Export
- Row 3: Transshipment/Storage

Columns:
- **C**: 20F count
- **D**: 40F count
- **E**: 20E count
- **F**: 40E count
- **G**: 45 count

### External Yard Structure
Each yard has 2 rows:
- Row 1: Import
- Row 2: Export

Columns:
- **C**: 20F count
- **D**: 40F count
- **E**: 20E count
- **F**: 40E count

## Troubleshooting

### "File not selected" error
- Make sure you select a valid file when prompted

### Wrong results
- Check block/yard mappings in code
- Verify row numbers are correct
- Ensure STOCK file column positions match expected structure

### Slow performance
- Enhanced version is optimized for speed
- For very large files (>50,000 rows), expect a few minutes

## Important Notes
1. ✅ **Backup your files** before running
2. ⚠️ The macro **clears old data** and refills from scratch
3. 📊 Excel formulas in yard files will recalculate automatically
4. 🚀 Optimized to process thousands of records quickly

## Support
- Review comments in the VBA code (in Arabic)
- Test on small sample data first
- Modify code sections as needed for your specific requirements

## Example Output

**Internal Yard - Block M:**
```
Row 6 (Import):   20F=150, 40F=200, 20E=50, 40E=75
Row 7 (Export):   20F=100, 40F=180, 20E=30, 40E=60
Row 8 (Storage):  20F=20,  40F=40,  20E=10, 40E=15
```

**External Yard - التجارية:**
```
Row 6 (Import):   20F=80, 40F=120, 20E=25, 40E=40
Row 7 (Export):   20F=60, 40F=90,  20E=15, 40E=30
```

---

**Version**: 1.0  
**Last Updated**: October 2025  
**Language**: VBA (Visual Basic for Applications)  
**Compatibility**: Excel 2010 and later

Good luck! 🚀
