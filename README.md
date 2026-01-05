# excel-delimiter-text-functions
Delimiter-based LEFT / RIGHT / MID functions for Excel (VBA Add-in)

# Excel Delimiter-Based Text Functions (VBA Add-In)

Small Excel VBA add-in providing delimiter-aware alternatives to `LEFT`, `RIGHT`, and `MID`.

## Functions
- `TextLeft(text, delimiter, n)` → text before the Nth delimiter  
- `TextRight(text, delimiter, n)` → text after the Nth delimiter  
- `TextMid(text, delimiter, n1, n2)` → text between the N1th and N2th delimiter  

Counts are always from the left. Delimiters may include spaces.

## Examples
=TextLeft("A - B - C - D", "-", 2)   → A - B  
=TextRight("A - B - C - D", "-", 3)  → D  
=TextMid("ISO - 25A1 - 12345 - P3", "-", 2, 3) → 12345

## Usage
Import the VBA module, save as `.xlam`, load once.  
Functions appear in Excel formula auto-suggestions.
