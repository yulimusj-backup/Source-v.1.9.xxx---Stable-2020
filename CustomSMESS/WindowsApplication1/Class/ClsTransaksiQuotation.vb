Public Class ClsTransaksiQuotation
    Public Function CetakQuotation(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Cetak Versi Detail
    Public Function CetakQuotationVersiDetail(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Double
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Cetak Versi Bahasa Indonesia Pendek
    Public Function CetakQuotationVersiIndoPendek(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Cetak Versi Bahasa English Pendek
    Public Function CetakQuotationVersiEnglishPendek(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Cetak Versi Bahasa English Panjang
    Public Function CetakQuotationVersiEnglishPanjang(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''Cetak Versi Bahasa Indonesia Panjang
    Public Function CetakQuotationVersiIndoPanjang(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
