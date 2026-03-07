Attribute VB_Name = "ProductDataFetcher"
' ===========================================================================
' ProductDataFetcher  -  Pull live product data from dummyjson.com
' ===========================================================================
' Calls:  https://dummyjson.com/products?limit=100
' Populates two Excel Tables:
'   tblProducts  (Products sheet)  - one row per product
'   tblReviews   (Reviews  sheet)  - one row per nested review
'
' Entry point:  RefreshProductData
' ===========================================================================

Option Explicit

Private Const API_URL As String = _
    "https://dummyjson.com/products?limit=100" & _
    "&select=id,title,category,price,rating,stock,brand,description,thumbnail,reviews"

Private Const PRODUCTS_SHEET As String = "Products"
Private Const REVIEWS_SHEET  As String = "Reviews"
Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_REVIEWS  As String = "tblReviews"

' Products table column positions
Private Const PCOL_ID          As Long = 1
Private Const PCOL_TITLE       As Long = 2
Private Const PCOL_CATEGORY    As Long = 3
Private Const PCOL_PRICE       As Long = 4
Private Const PCOL_RATING      As Long = 5
Private Const PCOL_STOCK       As Long = 6
Private Const PCOL_BRAND       As Long = 7
Private Const PCOL_DESCRIPTION As Long = 8
Private Const PCOL_THUMBNAIL   As Long = 9

' Reviews table column positions
Private Const RCOL_PRODUCT_ID As Long = 1
Private Const RCOL_REVIEWER   As Long = 2
Private Const RCOL_RATING     As Long = 3
Private Const RCOL_COMMENT    As Long = 4
Private Const RCOL_DATE       As Long = 5

' ===========================================================================
' PUBLIC
' ===========================================================================

' ---------------------------------------------------------------------------
' RefreshProductData
' Main entry point.  Fetch, parse and write all data, then refresh dashboard.
' ---------------------------------------------------------------------------
Public Sub RefreshProductData()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Connecting to dummyjson.com..."

    ' -- 1. Fetch JSON from API ------------------------------------------
    Dim jsonStr As String
    jsonStr = HttpGet(API_URL)

    If jsonStr = "" Then
        MsgBox "Could not retrieve data from the API." & Chr(10) & _
               "Please check your internet connection and try again.", _
               vbExclamation, "Connection Error"
        GoTo Cleanup
    End If

    Application.StatusBar = "Parsing JSON response..."

    ' -- 2. Parse JSON -------------------------------------------------------
    Dim root As Object
    Set root = ParseJson(jsonStr)

    Dim products As Collection
    Set products = GetArray(root, "products")

    If products Is Nothing Then
        MsgBox "The API returned an unexpected response format.", _
               vbExclamation, "Parse Error"
        GoTo Cleanup
    End If

    ' -- 3. Write to sheets --------------------------------------------------
    Application.StatusBar = "Writing " & products.Count & " products..."
    WriteProductsSheet products

    Application.StatusBar = "Writing reviews..."
    WriteReviewsSheet products

    ' -- 4. Refresh dashboard ------------------------------------------------
    Application.StatusBar = "Refreshing dashboard..."
    RefreshDashboard

    Application.StatusBar = "Ready  |  " & products.Count & " products loaded from dummyjson.com"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "RefreshProductData"
End Sub

' ===========================================================================
' PRIVATE  -  Sheet writers
' ===========================================================================

Private Sub WriteProductsSheet(ByVal products As Collection)
    Dim ws As Worksheet
    Set ws = EnsureSheet(PRODUCTS_SHEET)
    ClearSheet ws

    ' Headers
    ws.Cells(1, PCOL_ID).Value          = "ID"
    ws.Cells(1, PCOL_TITLE).Value       = "Title"
    ws.Cells(1, PCOL_CATEGORY).Value    = "Category"
    ws.Cells(1, PCOL_PRICE).Value       = "Price"
    ws.Cells(1, PCOL_RATING).Value      = "Rating"
    ws.Cells(1, PCOL_STOCK).Value       = "Stock"
    ws.Cells(1, PCOL_BRAND).Value       = "Brand"
    ws.Cells(1, PCOL_DESCRIPTION).Value = "Description"
    ws.Cells(1, PCOL_THUMBNAIL).Value   = "Thumbnail URL"

    ' Data rows
    Dim row As Long
    row = 2
    Dim prod As Variant
    For Each prod In products
        ws.Cells(row, PCOL_ID).Value          = GetNumber(prod, "id")
        ws.Cells(row, PCOL_TITLE).Value       = GetString(prod, "title")
        ws.Cells(row, PCOL_CATEGORY).Value    = GetString(prod, "category")
        ws.Cells(row, PCOL_PRICE).Value       = GetNumber(prod, "price")
        ws.Cells(row, PCOL_RATING).Value      = GetNumber(prod, "rating")
        ws.Cells(row, PCOL_STOCK).Value       = GetNumber(prod, "stock")
        ws.Cells(row, PCOL_BRAND).Value       = GetString(prod, "brand")
        ws.Cells(row, PCOL_DESCRIPTION).Value = GetString(prod, "description")
        ws.Cells(row, PCOL_THUMBNAIL).Value   = GetString(prod, "thumbnail")
        row = row + 1
    Next prod

    ' Format as Excel Table
    MakeTable ws, TABLE_PRODUCTS, row - 1, PCOL_THUMBNAIL

    ' Column formatting
    ws.Columns(PCOL_PRICE).NumberFormat  = "$#,##0.00"
    ws.Columns(PCOL_RATING).NumberFormat = "0.00"
    ws.Columns.AutoFit
End Sub

Private Sub WriteReviewsSheet(ByVal products As Collection)
    Dim ws As Worksheet
    Set ws = EnsureSheet(REVIEWS_SHEET)
    ClearSheet ws

    ' Headers
    ws.Cells(1, RCOL_PRODUCT_ID).Value = "Product ID"
    ws.Cells(1, RCOL_REVIEWER).Value   = "Reviewer"
    ws.Cells(1, RCOL_RATING).Value     = "Rating"
    ws.Cells(1, RCOL_COMMENT).Value    = "Comment"
    ws.Cells(1, RCOL_DATE).Value       = "Date"

    Dim row As Long
    row = 2

    Dim prod As Variant
    For Each prod In products
        Dim productId As Long
        productId = CLng(GetNumber(prod, "id"))

        Dim reviews As Collection
        Set reviews = GetArray(prod, "reviews")
        If reviews Is Nothing Then GoTo NextProduct

        Dim rev As Variant
        For Each rev In reviews
            ws.Cells(row, RCOL_PRODUCT_ID).Value = productId
            ws.Cells(row, RCOL_REVIEWER).Value   = GetString(rev, "reviewerName")
            ws.Cells(row, RCOL_RATING).Value     = GetNumber(rev, "rating")
            ws.Cells(row, RCOL_COMMENT).Value    = GetString(rev, "comment")
            ws.Cells(row, RCOL_DATE).Value       = GetString(rev, "date")
            row = row + 1
        Next rev

NextProduct:
    Next prod

    If row > 2 Then MakeTable ws, TABLE_REVIEWS, row - 1, RCOL_DATE
    ws.Columns.AutoFit
End Sub

' ===========================================================================
' PRIVATE  -  HTTP
' ===========================================================================

Private Function HttpGet(ByVal url As String) As String
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "User-Agent", "APIProductIntelligenceDemo/1.0"
    http.send
    If http.Status = 200 Then HttpGet = http.responseText
    Exit Function
Fail:
    HttpGet = ""
End Function

' ===========================================================================
' PRIVATE  -  Sheet / table helpers
' ===========================================================================

' Return the named sheet, creating it if it doesn't exist
Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

' Remove all ListObjects then clear all cells
Private Sub ClearSheet(ByVal ws As Worksheet)
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        tbl.Unlist
    Next tbl
    ws.Cells.Clear
End Sub

' Format a used range as an Excel Table with a blue style
Private Sub MakeTable(ByVal ws As Worksheet, ByVal tableName As String, _
                      ByVal lastDataRow As Long, ByVal lastCol As Long)
    If lastDataRow < 1 Then Exit Sub

    ' Drop any previous table with the same name
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not tbl Is Nothing Then tbl.Delete

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastDataRow, lastCol))
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name        = tableName
    tbl.TableStyle  = "TableStyleMedium2"
End Sub
