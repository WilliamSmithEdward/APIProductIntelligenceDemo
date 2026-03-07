# APIProductIntelligenceDemo

> **Excel + VBA demo** — pulls live product data & nested reviews from
> [dummyjson.com](https://dummyjson.com) into refreshable Excel Tables,
> with a real-time dashboard and a category dropdown filter.  
> Powered by the included **ModernJsonInVBA** library.  
> **Zero external dependencies.**

---

## What it does

| Feature | Detail |
|---|---|
| Live API call | `GET https://dummyjson.com/products?limit=100` |
| Nested data | Reviews array inside every product object is flattened into its own table |
| Refreshable tables | Two formatted Excel Tables: `tblProducts` and `tblReviews` |
| Dashboard | Summary stats (product count, avg price, avg rating, review count, category count) |
| Category dropdown | Data Validation list — filters the product table in real time |
| Auto-refresh | Button on the dashboard triggers a new API call |

---

## Repository layout

```
src/
  ModernJsonInVBA.bas       Pure-VBA JSON parser (the core library)
  ProductDataFetcher.bas    API call + sheet population
  DashboardController.bas   Dashboard stats, dropdown, filtered table
  WorkbookSetup.bas         Workbook init + "Refresh Data" button macro
  ThisWorkbook.cls          Workbook_Open and Workbook_SheetChange events

scripts/
  build.ps1                 PowerShell script to assemble the .xlsm file
```

---

## Quick start

### Option A — Build from source (recommended)

Prerequisites:
- Windows with Microsoft Excel installed (2016 or later).
- In Excel ▸ File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸
  Macro Settings: **enable "Trust access to the VBA project object model"**.

```powershell
# From the repository root:
.\scripts\build.ps1
```

This creates `APIProductIntelligenceDemo.xlsm` in the repo root.  
Open it in Excel — data loads automatically on `Workbook_Open`.

### Option B — Manual import

1. Open a new Excel workbook and save it as a **macro-enabled workbook** (`.xlsm`).
2. Open the VBA editor (`Alt + F11`).
3. Import the four `.bas` files from `src/` via **File ▸ Import File**.
4. Copy the body of `src/ThisWorkbook.cls` into the `ThisWorkbook` code module.
5. Press `F5` or call `InitializeWorkbook` from the Immediate Window.

---

## How it works

### 1 — ModernJsonInVBA (JSON parser)

A hand-written recursive-descent parser — no `ScriptControl`, no COM object,
no third-party DLL.  Handles the full JSON spec including Unicode escapes.

```vba
' Parse any JSON string → Dictionary / Collection / primitive
Dim root As Object
Set root = ParseJson(jsonStr)

' Safe helper functions (never throw on missing keys)
Dim title    As String  : title    = GetString(product, "title")
Dim price    As Double  : price    = GetNumber(product, "price")
Dim reviews  As Collection : Set reviews = GetArray(product, "reviews")

' Dot + bracket path navigation
Dim firstReview As String
firstReview = GetNestedValue(root, "products[0].reviews[0].comment")
```

### 2 — ProductDataFetcher (API → Excel Tables)

```
https://dummyjson.com/products?limit=100
        │
        ▼ MSXML2.XMLHTTP.6.0
   raw JSON string
        │
        ▼ ParseJson()
   root Dictionary
        │
        ├─ products  (Collection of 100 Dictionaries)
        │       ▼
        │   tblProducts  (Products sheet)
        │
        └─ products[n].reviews  (nested Collection)
                ▼
            tblReviews  (Reviews sheet)
```

### 3 — DashboardController (real-time dashboard)

The Dashboard sheet is built programmatically:

```
Row 1   Title banner
Row 3   "Live Statistics" label
Row 4   | Products: 100 | Avg Price: $… | Avg Rating: 4.xx | Reviews: … | Categories: … |
Row 6   "Filter by Category:" label
Row 7   [  All Categories  ▼ ]   ← Data Validation dropdown
Row 9   Section heading
Row 10  Column headers (blue bar)
Row 11+ Filtered product rows
```

Selecting a different category in the dropdown triggers `Workbook_SheetChange`
→ `OnCategoryChange` → `UpdateFilteredTable` — the table rebuilds instantly
without a new API call.

---

## Modules reference

### `ModernJsonInVBA.bas`

| Function | Signature | Returns |
|---|---|---|
| `ParseJson` | `(jsonStr As String) As Variant` | Dictionary / Collection / primitive |
| `GetNestedValue` | `(node, path As String) As Variant` | Any nested value via `"a.b[0].c"` |
| `GetString` | `(dict, key [, default]) As String` | String (safe) |
| `GetNumber` | `(dict, key [, default]) As Double` | Double (safe) |
| `GetArray` | `(dict, key) As Collection` | Collection or Nothing |
| `GetObject` | `(dict, key) As Object` | Dictionary or Nothing |

### `ProductDataFetcher.bas`

| Sub | Description |
|---|---|
| `RefreshProductData` | Full refresh: fetch → parse → write sheets → update dashboard |

### `DashboardController.bas`

| Sub | Description |
|---|---|
| `SetupDashboard` | One-time layout (called from `InitializeWorkbook`) |
| `RefreshDashboard` | Recalculate stats + dropdown + table (called after data load) |
| `OnCategoryChange` | Redraw filtered table (called from `Workbook_SheetChange`) |

### `WorkbookSetup.bas`

| Sub | Description |
|---|---|
| `InitializeWorkbook` | Called from `Workbook_Open`; sets up & loads |
| `RefreshAll` | Assigned to the "Refresh Data" button |

---

## dummyjson.com API

The demo uses the free, no-auth-required [dummyjson.com](https://dummyjson.com) API.

Sample response shape:

```json
{
  "products": [
    {
      "id": 1,
      "title": "Essence Mascara Lash Princess",
      "category": "beauty",
      "price": 9.99,
      "rating": 4.94,
      "stock": 5,
      "brand": "Essence",
      "description": "...",
      "thumbnail": "https://cdn.dummyjson.com/...",
      "reviews": [
        {
          "rating": 2,
          "comment": "Very unhappy...",
          "date": "2024-05-23T08:56:21.618Z",
          "reviewerName": "John Doe",
          "reviewerEmail": "john.doe@x.com"
        }
      ]
    }
  ],
  "total": 194,
  "skip": 0,
  "limit": 100
}
```

---

## Requirements

- Windows (MSXML2.XMLHTTP is a Windows component)
- Microsoft Excel 2016 or later (32-bit or 64-bit)
- Active internet connection for the API call
- Macros enabled in Excel

---

## License

MIT
