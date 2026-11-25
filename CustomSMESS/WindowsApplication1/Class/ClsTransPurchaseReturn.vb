Option Explicit On
Option Strict Off

Public Class ClsTransPurchaseReturn
    'fungsi cetak custom pada smess
    Public Function Cetak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPO As Double, _
    ByVal vQuickPR As Boolean, ByVal vPrimarySupplier As Double) As Boolean
        Return True
    End Function
End Class
