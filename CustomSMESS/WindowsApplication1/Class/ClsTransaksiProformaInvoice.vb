Public Class ClsTransaksiProformaInvoice
    Public Function CetakProformaInvoice_Custom(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimarysalesInvoice As Double, ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer, ByVal vJabatanProforma As String) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
