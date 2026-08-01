Option Compare Database
Option Explicit

Public Function CalculateReorderPoint(ByVal intCurrentStock As Integer, ByVal intDailySales As Integer) As Integer
    ' Calculate when to reorder based on 7-day lead time
    CalculateReorderPoint = intDailySales * 7
End Function

Public Function GetStockStatus(ByVal intQty As Integer) As String
    Select Case intQty
        Case Is < 5:  GetStockStatus = "CRITICAL"
        Case Is < 10: GetStockStatus = "LOW"
        Case Is < 50: GetStockStatus = "NORMAL"
        Case Else:    GetStockStatus = "OVERSTOCKED"
    End Select
End Function
Public Sub ExportToExcel(ByVal strQueryName As String)
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
        strQueryName, CurrentProject.Path & "\" & strQueryName & ".xlsx", True
    MsgBox "Exported to " & CurrentProject.Path & "\" & strQueryName & ".xlsx"
End Sub