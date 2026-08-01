Option Compare Database
Option Explicit

Public Function FormatCurrency2(ByVal dblAmount As Double) As String
    FormatCurrency2 = Format(dblAmount, "$#,##0.00")
End Function

Public Function IsLowStock(ByVal intQty As Integer) As Boolean
    IsLowStock = (intQty < 10)
End Function

Public Sub RefreshMainForm()
    If CurrentProject.AllForms("frmProducts").IsLoaded Then
        Forms("frmProducts").Requery
    End If
End Sub
