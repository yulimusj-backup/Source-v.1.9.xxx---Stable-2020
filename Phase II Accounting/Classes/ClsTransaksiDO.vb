Option Explicit On
Option Strict Off

Public Class ClsTransaksiDO
    Public Function GetNoInvoice(ByVal vPrimaryDO As Double, ByVal DTA1 As DseTransaksiDOTableAdapters.DataTable1TableAdapter, _
    ByVal TBL1 As DseTransaksiDO.DataTable1DataTable) As String
        Dim row As DseTransaksiDO.DataTable1Row

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            DTA1.Fill(TBL1, vPrimaryDO)
            If TBL1.Rows.Count <> 0 Then
                row = TBL1.Rows(0)
                Return row.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
