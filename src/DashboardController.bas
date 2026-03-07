Attribute VB_Name = "DashboardController"
' ===========================================================================
' DashboardController  -  Real-time dashboard with category dropdown
' ===========================================================================
' Manages the "Dashboard" sheet which shows:
'   Row 1       Title banner
'   Row 3-4     Summary statistics (product count, avg price, avg rating,
'               total reviews, category count)
'   Row 6-7     Category filter – Data Validation dropdown
'   Row 9+      Filtered products table (updates on dropdown change)
'
' Public entry points
' -------------------
'   SetupDashboard     – called once from WorkbookSetup.InitializeWorkbook
'   RefreshDashboard   – called after every data refresh
'   OnCategoryChange   – called from Workbook_SheetChange when A7 changes
' ===========================================================================

Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const PRODUCTS_SHEET  As String = "Products"
Private Const REVIEWS_SHEET   As String = "Reviews"

' Fixed row positions on the Dashboard sheet
Private Const ROW_TITLE        As Long = 1
Private Const ROW_STATS_HDR    As Long = 3
Private Const ROW_STATS        As Long = 4
Private Const ROW_FILTER_LABEL As Long = 6
Private Const ROW_FILTER       As Long = 7   ' <- dropdown cell lives in A7
Private Const ROW_TABLE_HDR    As Long = 9

' ===========================================================================
' PUBLIC
' ===========================================================================

' ---------------------------------------------------------------------------
' SetupDashboard
' One-time layout initialisation.  Safe to call again (it clears & redraws).
' ---------------------------------------------------------------------------
Public Sub SetupDashboard()
    Dim ws As Worksheet
    Set ws = EnsureDashboardSheet()
    ws.Cells.Clear

    ' ---- Title ----
    With ws.Range("A1")
        .Value      = "Product Intelligence Dashboard"
        .Font.Size  = 20
        .Font.Bold  = True
        .Font.Color = RGB(31, 73, 125)
    End With
    ws.Range("A1:J1").Merge

    ' ---- Stats header ----
    With ws.Cells(ROW_STATS_HDR, 1)
        .Value = "Live Statistics"
        .Font.Bold  = True
        .Font.Size  = 11
        .Font.Color = RGB(31, 73, 125)
    End With

    ' Stat labels (every 2 columns)
    ws.Cells(ROW_STATS, 1).Value  = "Products:"
    ws.Cells(ROW_STATS, 3).Value  = "Avg Price:"
    ws.Cells(ROW_STATS, 5).Value  = "Avg Rating:"
    ws.Cells(ROW_STATS, 7).Value  = "Reviews:"
    ws.Cells(ROW_STATS, 9).Value  = "Categories:"

    Dim lbl As Variant
    For Each lbl In Array(1, 3, 5, 7, 9)
        With ws.Cells(ROW_STATS, CLng(lbl))
            .Font.Bold  = True
            .Font.Color = RGB(64, 64, 64)
        End With
    Next lbl

    ' ---- Filter label ----
    With ws.Cells(ROW_FILTER_LABEL, 1)
        .Value     = "Filter by Category:"
        .Font.Bold = True
    End With

    ' ---- Dropdown placeholder (real list built by UpdateCategoryDropdown) ----
    ws.Cells(ROW_FILTER, 1).Value = "All Categories"
    With ws.Cells(ROW_FILTER, 1)
        .Interior.Color   = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Bold        = True
    End With

    ' ---- Table heading placeholder ----
    With ws.Cells(ROW_TABLE_HDR, 1)
        .Value     = "Products (click Refresh Data to load)"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(31, 73, 125)
    End With

    ws.Columns.AutoFit
End Sub

' ---------------------------------------------------------------------------
' RefreshDashboard
' Recalculate stats and redraw the filtered table.  Called after data load.
' ---------------------------------------------------------------------------
Public Sub RefreshDashboard()
    On Error GoTo ErrHandler

    Dim ws     As Worksheet
    Dim wsProd As Worksheet
    Dim wsRev  As Worksheet

    Set ws = EnsureDashboardSheet()

    On Error Resume Next
    Set wsProd = ThisWorkbook.Sheets(PRODUCTS_SHEET)
    Set wsRev  = ThisWorkbook.Sheets(REVIEWS_SHEET)
    On Error GoTo ErrHandler

    If wsProd Is Nothing Then Exit Sub
    If wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row <= 1 Then Exit Sub

    UpdateStats          ws, wsProd, wsRev
    UpdateCategoryDropdown ws, wsProd
    UpdateFilteredTable  ws, wsProd, GetSelectedCategory(ws)
    ws.Columns.AutoFit
    Exit Sub

ErrHandler:
    ' Dashboard not fully ready yet – silently ignore
End Sub

' ---------------------------------------------------------------------------
' OnCategoryChange
' Fired by Workbook_SheetChange when cell A7 on Dashboard changes.
' ---------------------------------------------------------------------------
Public Sub OnCategoryChange()
    Dim ws     As Worksheet
    Dim wsProd As Worksheet

    Set ws = EnsureDashboardSheet()

    On Error Resume Next
    Set wsProd = ThisWorkbook.Sheets(PRODUCTS_SHEET)
    On Error GoTo 0
    If wsProd Is Nothing Then Exit Sub

    UpdateFilteredTable ws, wsProd, GetSelectedCategory(ws)
    ws.Columns.AutoFit
End Sub

' ===========================================================================
' PRIVATE  -  Stats
' ===========================================================================

Private Sub UpdateStats(ByVal wsDash As Worksheet, _
                        ByVal wsProd As Worksheet, _
                        ByVal wsRev  As Worksheet)
    Dim lastRow As Long
    lastRow = wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then Exit Sub

    ' -- Counts --
    Dim totalProducts As Long
    totalProducts = lastRow - 1

    Dim avgPrice As Double
    avgPrice = WorksheetFunction.Average(wsProd.Range("D2:D" & lastRow))

    Dim avgRating As Double
    avgRating = WorksheetFunction.Average(wsProd.Range("E2:E" & lastRow))

    Dim totalReviews As Long
    If Not wsRev Is Nothing Then
        Dim revLast As Long
        revLast = wsRev.Cells(wsRev.Rows.Count, 1).End(xlUp).Row
        totalReviews = IIf(revLast > 1, revLast - 1, 0)
    End If

    Dim catCount As Long
    catCount = CountUniqueValues(wsProd, 3)

    ' -- Write stat values --
    WriteStatValue wsDash, ROW_STATS, 2,  totalProducts, "0"
    WriteStatValue wsDash, ROW_STATS, 4,  avgPrice,      "$#,##0.00"
    WriteStatValue wsDash, ROW_STATS, 6,  avgRating,     "0.00"
    WriteStatValue wsDash, ROW_STATS, 8,  totalReviews,  "0"
    WriteStatValue wsDash, ROW_STATS, 10, catCount,      "0"
End Sub

Private Sub WriteStatValue(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                            ByVal val As Variant, ByVal fmt As String)
    With ws.Cells(r, c)
        .Value         = val
        .NumberFormat  = fmt
        .Font.Bold     = True
        .Font.Size     = 14
        .Font.Color    = RGB(0, 112, 192)
    End With
End Sub

' ===========================================================================
' PRIVATE  -  Category dropdown
' ===========================================================================

Private Sub UpdateCategoryDropdown(ByVal wsDash As Worksheet, _
                                   ByVal wsProd  As Worksheet)
    Dim lastRow As Long
    lastRow = wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then Exit Sub

    ' Collect unique categories in order
    Dim seen   As Object
    Dim cats() As String
    Dim count  As Long
    Set seen = CreateObject("Scripting.Dictionary")
    count = 0

    Dim i As Long
    For i = 2 To lastRow
        Dim cat As String
        cat = Trim(wsProd.Cells(i, 3).Value)
        If cat <> "" And Not seen.Exists(cat) Then
            seen.Add cat, 1
            ReDim Preserve cats(0 To count)
            cats(count) = cat
            count = count + 1
        End If
    Next i

    ' Build comma-delimited list for Data Validation
    Dim listStr As String
    listStr = "All Categories"
    Dim j As Long
    For j = 0 To count - 1
        listStr = listStr & "," & cats(j)
    Next j

    ' Apply Data Validation drop-down to A7
    With wsDash.Cells(ROW_FILTER, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listStr
        .ShowInput  = True
        .ShowError  = False
    End With
End Sub

' ===========================================================================
' PRIVATE  -  Filtered products table
' ===========================================================================

Private Sub UpdateFilteredTable(ByVal wsDash As Worksheet, _
                                 ByVal wsProd  As Worksheet, _
                                 ByVal catFilter As String)
    ' Clear everything from row ROW_TABLE_HDR downward
    Dim clearStart As Long
    clearStart = ROW_TABLE_HDR
    With wsDash.Range(wsDash.Cells(clearStart, 1), _
                      wsDash.Cells(wsDash.Rows.Count, 10))
        .Clear
    End With

    ' ---- Section heading ----
    Dim headingText As String
    If catFilter = "All Categories" Then
        headingText = "All Products"
    Else
        headingText = "Category: " & catFilter
    End If
    With wsDash.Cells(ROW_TABLE_HDR, 1)
        .Value     = headingText
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(31, 73, 125)
    End With

    ' ---- Column headers (row ROW_TABLE_HDR + 1) ----
    Dim hdrRow As Long
    hdrRow = ROW_TABLE_HDR + 1
    wsDash.Cells(hdrRow, 1).Value = "ID"
    wsDash.Cells(hdrRow, 2).Value = "Title"
    wsDash.Cells(hdrRow, 3).Value = "Category"
    wsDash.Cells(hdrRow, 4).Value = "Price"
    wsDash.Cells(hdrRow, 5).Value = "Rating"
    wsDash.Cells(hdrRow, 6).Value = "Stock"
    wsDash.Cells(hdrRow, 7).Value = "Brand"

    With wsDash.Range(wsDash.Cells(hdrRow, 1), wsDash.Cells(hdrRow, 7))
        .Font.Bold     = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color    = RGB(255, 255, 255)
    End With

    ' ---- Data rows ----
    Dim lastProd As Long
    lastProd = wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row

    Dim destRow As Long
    destRow = hdrRow + 1

    Dim k As Long
    For k = 2 To lastProd
        Dim prodCat As String
        prodCat = wsProd.Cells(k, 3).Value

        If catFilter = "All Categories" Or prodCat = catFilter Then
            wsDash.Cells(destRow, 1).Value = wsProd.Cells(k, 1).Value  ' ID
            wsDash.Cells(destRow, 2).Value = wsProd.Cells(k, 2).Value  ' Title
            wsDash.Cells(destRow, 3).Value = wsProd.Cells(k, 3).Value  ' Category
            wsDash.Cells(destRow, 4).Value = wsProd.Cells(k, 4).Value  ' Price
            wsDash.Cells(destRow, 5).Value = wsProd.Cells(k, 5).Value  ' Rating
            wsDash.Cells(destRow, 6).Value = wsProd.Cells(k, 6).Value  ' Stock
            wsDash.Cells(destRow, 7).Value = wsProd.Cells(k, 7).Value  ' Brand

            ' Alternating row shading (1st data row = destRow - hdrRow == 1 -> shaded)
            If (destRow - hdrRow) Mod 2 = 1 Then
                wsDash.Range(wsDash.Cells(destRow, 1), _
                             wsDash.Cells(destRow, 7)).Interior.Color = RGB(220, 230, 241)
            End If

            destRow = destRow + 1
        End If
    Next k

    ' Format numeric columns
    If destRow > hdrRow + 1 Then
        wsDash.Range(wsDash.Cells(hdrRow + 1, 4), _
                     wsDash.Cells(destRow - 1, 4)).NumberFormat = "$#,##0.00"
        wsDash.Range(wsDash.Cells(hdrRow + 1, 5), _
                     wsDash.Cells(destRow - 1, 5)).NumberFormat = "0.00"
    End If
End Sub

' ===========================================================================
' PRIVATE  -  Utilities
' ===========================================================================

Private Function EnsureDashboardSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = DASHBOARD_SHEET
    End If
    Set EnsureDashboardSheet = ws
End Function

Private Function GetSelectedCategory(ByVal ws As Worksheet) As String
    GetSelectedCategory = CStr(ws.Cells(ROW_FILTER, 1).Value)
End Function

Private Function CountUniqueValues(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow <= 1 Then Exit Function

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lastRow
        Dim v As String
        v = Trim(ws.Cells(i, col).Value)
        If v <> "" Then seen(v) = 1
    Next i
    CountUniqueValues = seen.Count
End Function
