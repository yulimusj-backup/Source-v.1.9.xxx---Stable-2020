Option Explicit On
Option Strict Off
Public Class ClsTransaksiAOP
    Public Function ApakahBolehEdit_Custom(ByVal vPrimary_AOP_SMESS As Double) As Boolean
        Return True
    End Function

    Public Function ApakahBolehHapus_Custom(ByVal vPrimary_AOP_SMESS As Double) As Boolean
        Return True
    End Function
    Public Sub CetakCustomAOP(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryAOP As Double)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
