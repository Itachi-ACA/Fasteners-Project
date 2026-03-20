
INDOTECH HARDWARE AUTOMATION PLATFORM
======================================

1) HOW TO USE
--------------
1. Click Select Hardware PDF.
2. Choose PDF.
3. Select output folder.
4. Excel & PDF auto generated.

2) IF APPLICATION FAILS
-----------------------
- Close already opened Excel file.
- Ensure PDF is not scanned image.
- Ensure Excel is installed.

3) WEIGHT TABLE
---------------
Only HDG coating values are active.
MAGNI, BLACK, SS, HT will show N/A until
their values are added to WEIGHT_TABLE.

To add MAGNI rates later, open the code and add:
  "M16X50-MAGNI": 27,
under the WEIGHT_TABLE section.

4) ROUNDING RULE
----------------
Odd lengths (X25, X35, X45...) are automatically
rounded up to nearest 10 (X30, X40, X50...).
Output displays the rounded size.

======================================
IndoTech Internal Software
======================================
