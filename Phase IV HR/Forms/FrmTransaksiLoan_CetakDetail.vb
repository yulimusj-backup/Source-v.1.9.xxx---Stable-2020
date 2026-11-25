Public Class FrmTransaksiLoan_CetakDetail
    Public PrimaryLoan As Double = -1
    Public Type As String = ""
    Public ObjLoan As ClsTransaksiLoan
    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ObjLoan.CetakDetail(Me.PrimaryLoan, Me.crvReport, Me.Type)
            'If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
            '    ObjCust.CetakJurnal(Me.PrimaryJurnal, Me.crvReport)
            'Else
            '    Me.ObjLoan.CetakDetail(Me.PrimaryJurnal, Me.crvReport, Me.DS)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class