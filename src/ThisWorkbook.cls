VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ===========================================================================
' ThisWorkbook  -  Workbook-level event handlers
' ===========================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Workbook_Open
' Runs automatically when the file is opened.
' Fetches live data and sets up the dashboard.
' ---------------------------------------------------------------------------
Private Sub Workbook_Open()
    InitializeWorkbook
End Sub

' ---------------------------------------------------------------------------
' Workbook_SheetChange
' Detects changes to the Category dropdown on the Dashboard sheet (cell A7)
' and immediately redraws the filtered products table.
' ---------------------------------------------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Dashboard" Then Exit Sub
    If Target.Address <> "$A$7" Then Exit Sub

    ' Prevent re-entrancy while we update the table
    Application.EnableEvents = False
    On Error GoTo Done

    OnCategoryChange

Done:
    Application.EnableEvents = True
End Sub
