Public Class FrmTransaksiJournal_Cetak_Sales
    Public PrimaryJurnal As Double = -1
    Public vDGV As DataGridView
    Public vFrm As FrmTransaksiJournal_Sales

    Dim Obj As New ClsTransaksiJournal
    Dim DS As New DseTransaksiJournal
    Dim Report As New CrpCetakJurnal
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCust As New CustomSMESS.ClsTransaksiJournal

    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                ObjCust.CetakJournal_DetailInvoice_Sales(Me.PrimaryJurnal, Me.crvReport, Me.vDGV, Me.vFrm.ColReceivementOriginal.Index, _
                Me.vFrm.ColCurrencyBalance.Index, Me.vFrm.ColInvoiceNo.Index, Me.vFrm.ColInvoiceDate.Index, Me.vFrm.ColExhangeRateBalance.Index)
            Else
                Obj.CetakJournal_DetailInvoice_Sales(Me.PrimaryJurnal, Me.crvReport, Me.DS, vDGV, vFrm)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class