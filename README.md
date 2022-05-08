# node-red-buffer-xlsx
A simple node-red node for converting JSON to a buffered xlsx. This package uses the node package better-xlsx.

## Usage
To ensure the node works correctly the JSON has to be formatted in a certain way. All styling options are optional.

### Sheet Formatting
The JSON can consist of multiple sheets, these are the properties each sheet can have:
- sheet_name
- sheet_styling
- header_styling
- columns_styling
- rows

  **sheet_name** (required):
  This is the name the sheet shall receive at the bottom.

  **sheet_styling** (optional):
  Styling for the entire sheet, see **Styling Formats**

  **header_styling** (optional):
  Styling for the header row, see **Styling Formats**

  **columns_styling** (optional):
  columns_styling should be an array of object with the following info:
  - index
  - column_styling

    *index*: The column you want to apply the styling to starting at 0

    *column_styling*: See **Styling Formats**

  **rows** (required):
  See **Row Formatting**.
  
 **Sheet Example**:
 ```
 [
  {
    "sheet_name": "sheet1",
    "sheet_styling": {
      "hAlign": "center",
      "fSize": "11"
    },
    "header_styling": {
      "fBold": true
    },
    "columns_styling": [
      {
        "index": 0,
        "column_styling": {
          "fItalic": true
        }
      }
    ],
    "rows": [
      {
        // See # Row Formatting
      }
    ]
  }
]
```

### Row Formatting
The JSON can consist of multiple rows, these are the properties each row can have:
- row_styling
- cells

  **row_styling** (optional):
  Styling for the entire row, see **Styling Formats**

  **cells** (required):
  See **Cell Formatting**.

**Row Example**:
```
"rows": [
  {
    "row_styling": {
      "hAlign": "left"
     },
    "cells": [
      {
        // See # Cells Formatting
      }
    ]
  }
]
```

### Cell Formatting
The JSON can consist of multiple cells, these are the properties each cell can have:
- cell_value
- cell_styling

  **cell_value** (optional):
  Value that will be displayed in the cell

  **cell_styling** (optional):
  Styling for single cell, see **Styling Formats**.

**Cell Example**
```
"cells": [
  {
    "cell_value": "Hello there!",
    "cell_styling": {
      "fColor": "ffa2917d"
    }
  }
]
```

### Styling Formats
Each of the properties ending with styling can include the following optional properties:
- pattern_type
- fgColor
- bgColor
- hAlign
- vAlign
- indent
- shrinkToFit
- textRotation
- wrapText
- fSize
- fName
- fFamily
- fCharset
- fColor
- fBold
- fItalic
- fUnderline
- cell_format
- cell_formula

  **pattern_type**:
  Specifies if cell background should be filled or not with the following options:
  - solid
  - none

  **fgColor**:
  Decides the foreground colour of the cell, all colours should be in HEX RBGA format without the #.

  **bgColor**:
  Deicdes the background colour of the cell, all colours should be in HEX RGBA format without the #.

  **hAlign**:
  Horizontal Alignment of cell values with the following options:
  - general
  - center
  - left
  - right

  **vAlign**:
  Vertical Alignment of cell values with the following options:
  - general
  - top
  - bottom
  - center

  **indent**:
  Decides indent of cell, should be a numeric value.

  **shrinkToFit**:
  Decides whether the cell value should shrink to fit the cell, should be a boolean.

  **textRotation**:
  Decides the rotation of the cell value, should be a numeric value between 0 and 255.

  **wrapText**:
  Decides whether the cell value should wrap to fit the cell, should be a boolean.

  **fSize**:
  Decides the cell value's font size, should be a numeric value.

  **fName**:
  Decides the cell value's font family, should be a font family.
  
  **fFamily**:
  - TBD

  **fCharset**:
  Decides the cell value's charset, should be a charset.

  **fColor**:
  Decides the cell value's colour, all colours should be in HEX RGBA format without the #.

  **fBold**:
  Decides whether the cell's value is Bold or not, should be a boolean.

  **fItalic**:
  Decides whether the cell's value is Italic or not, should be a boolean.

  **fUnderline**:
  Decides whether the cell's value is underlined or not, should be a boolean.

  **cell_format**:
  Decides number formatting for the cell with following options as example:
  ```
  "0" // integer index to built in formats
  "0.00%" // string matching a built-in format
  "0.0%" // string specifying a custom format
  "0.00%;\\(0.00%\\);\\-;@" // string specifying a custom format, escaping special characters
  "m/dd/yy" // string a date format using Excel's format notation
  ```

  **cell_formula**:
  Decides the cell's formula, see excel formulas for more.
  
  **Borders**:
  Decides the cell's borders formatting using the following parameters:
  - all
  - top
  - right
  - bottom
  - left

  Each of them can have the following:
  
  - style 
  - bColor

   **style**:
   Style can consist of the following types of border styles:
   ```
   thin
   medium
   thick
   dotted
   hair
   dashed
   mediumDashed
   dashDot
   mediumDashDot
   dashDotDot
   mediumDashDotDot
   slantDashDot
   ```
   
   **bColor**:
   Decides the colour of the borders, all colours should be in HEX RGBA format without the #.

**Styling Priorities**:
Styling is done in the following order:
1. Cell Styling
2. Header Styling
3. Column Styling
4. Row Styling
5. Sheet Styling

Unless Style Merging is selected the top priority will always be the only style chosen. If Style Merging is selected however the node will attempt to merge styles while still following the priority.

**Full Styling Example**:
```
{
  "pattern_type": "solid",
  "fgColor": "ffa2917d",
  "bgColor": "43ff64d9",
  "hAlign": "center",
  "vAlign": "top",
  "indent": "3",
  "shrinkToFit": false,
  "textRotation": "155",
  "wrapText": true,
  "fSize": "11",
  "fName": "Calibri",
  "fFamily": "Calibri",
  "fCharset": "UTF-8",
  "fColor": "9b0f64d9",
  "fBold": true,
  "fItalic": false,
  "fUnderline": false,
  "cell_format": "m/dd/yy",
  "cell_formula": "A1 - C2",
  "borders": {
    "top": {
      "style": "medium",
      "bColor": "009b9bd9"
    },
    "right": {
      "style": "thin",
      "bColor": "9b0f64d9"
    },
    "bottom": {
      "style": "thick",
      "bColor": "9b0f64d9"
    },
    "left": {
      "style": "dashDotDot",
      "bColor": "009b9bd9"
    }
  }
}
```

## Full Examples
Here some examples of full JSON

### Simple JSON
```
[
  {
    "sheet_name": "SimpleSheet",
    "sheet_styling": {
      "pattern_type": "solid",
      "fgColor": "ffffffff",
      "bgColor": "FE0000",
      "hAlign": "left",
      "vAlign": "left",
      "borders": {
       "all": {
        "style": "thick",
        "bColor": "009b9bd9"
       }
      }
    },
    "header_styling": {
      "pattern_type": "solid",
      "fgColor": "ffffffff",
      "bgColor": "ffe4e2de",
      "hAlign": "center",
      "vAlign": "center",
      "fBold": true
    },
    "columns_styling": [
      {
        "index": 0,
        "column_styling": {
          "hAlign": "right",
          "vAlign": "center",
          "fBold": true
        }
      },
      {
        "index": 4,
        "column_styling": {
          "hAlign": "right",
          "vAlign": "top",
          "fItalic": true
        }
      }
    ],
    "rows": [
      {
        "row_styling": {
          "pattern_type": "solid",
          "fgColor": "ffffffff",
          "bgColor": "ffe4e2de",
          "hAlign": "center",
          "vAlign": "center"
        },
        "cells": [
          {
            "cell_value": "ID",
            "cell_formula": "",
            "cell_format": "",
            "cell_styling": {
              "fBold": false
            }
          },
          {
            "cell_value": "Name",
            "cell_styling": {
              "fBold": false
            }
          },
          {
            "cell_value": "Company",
            "cell_styling": {
              "fBold": false
            }
          },
          {
            "cell_value": "Location",
            "cell_formula": "",
            "cell_format": "",
            "cell_styling": {
              "fBold": false
            }
          },
          {
            "cell_value": "Price"
          },
          {
            "cell_value": "Amount"
          }
        ]
      },
      {
        "row_styling": {
          "pattern_type": "solid",
          "fgColor": "ffffffff",
          "bgColor": "ffe4e2de",
          "hAlign": "center",
          "vAlign": "center"
        },
        "cells": [
          {
            "cell_value": "238126"
          },
          {
            "cell_value": "James. B."
          },
          {
            "cell_value": "Makers.BV",
            "cell_styling": {
              "pattern_type": "solid",
              "bgColor": "430f6480",
              "fColor": "FF000000"
            }
          },
          {
            "cell_styling": {
              "fBold": true,
              "fgColor": "43ff6480"
            }
          },
          {
            "cell_value": "152,00",
            "cell_format": "$0,00"
          },
          {
            "cell_value": "16,00",
            "cell_formula": "A2 - E2"
          }
        ]
      }
    ]
  }
]
```
### Complex JSON
- To be made

# TODO:
- [ ] Testing
- [ ] Complex JSON
- [ ] Ability for more complex features
- [X] Border Support

