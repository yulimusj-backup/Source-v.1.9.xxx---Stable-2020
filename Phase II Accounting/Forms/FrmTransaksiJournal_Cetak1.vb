Option Explicit On
Option Strict Off

Public Class FrmTransaksiJournal_Cetak1
    Public PrimaryJurnal As Double = -1
    Dim Obj As New ClsTransaksiJournal
    Dim DS As New DseTransaksiJournal
    Dim Report As New CrpCetakJurnal
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
    Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping

    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            TipeJurnal2 = Me.Obj.GetTipeJournal2(PrimaryJurnal)
            If TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank Then
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                    ObjCust.CetakJurnal_Cash_n_Bank1(Me.PrimaryJurnal, Me.crvReport)
                Else
                    Obj.CetakJournal(Me.PrimaryJurnal, Me.crvReport, Me.DS)
                End If
            Else
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                    ObjCust.CetakJurnal(Me.PrimaryJurnal, Me.crvReport)
                Else
                    Obj.CetakJournal(Me.PrimaryJurnal, Me.crvReport, Me.DS)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class