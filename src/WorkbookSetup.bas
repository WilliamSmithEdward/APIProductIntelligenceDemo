Attribute VB_Name = "WorkbookSetup"
' ===========================================================================
' WorkbookSetup  -  Workbook initialisation and public macro entry points
' ===========================================================================
' Called from ThisWorkbook.Workbook_Open  on first launch.
' Also exposes the public macros that can be assigned to buttons.
' ===========================================================================

Option Explicit

' ---------------------------------------------------------------------------
' InitializeWorkbook
' Sets up the dashboard layout then fetches live data.
' Called automatically from Workbook_Open.
' ---------------------------------------------------------------------------
Public Sub InitializeWorkbook()
    Application.ScreenUpdating = False

    ' Build the dashboard shell (layout, labels, placeholder dropdown)
    SetupDashboard

    ' Pull live data from dummyjson.com  (also calls RefreshDashboard)
    RefreshProductData

    ' Land on the dashboard
    ThisWorkbook.Sheets("Dashboard").Activate

    Application.ScreenUpdating = True

    MsgBox "Product Intelligence Dashboard is ready!" & Chr(10) & Chr(10) & _
           "• Use the Category dropdown to filter products." & Chr(10) & _
           "• Click Refresh Data to pull the latest data from the API.", _
           vbInformation, "APIProductIntelligenceDemo"
End Sub

' ---------------------------------------------------------------------------
' RefreshAll
' Public macro assigned to the Refresh Data button on the dashboard.
' ---------------------------------------------------------------------------
Public Sub RefreshAll()
    RefreshProductData
End Sub
