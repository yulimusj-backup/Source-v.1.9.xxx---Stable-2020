Option Explicit On
Option Strict Off

Public Class FrmTransaksiCreditNote_Cetak
    Dim Obj As New ClsTransaksiJournal
    Dim DS As New DseTransaksiJournal
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Public TipeTransaksi As New ClsTransaksiCreditNote.EnumTipeTransaksi
    Public TipeKartu As New ClsTransaksiCreditNote.EnumTipeKartu
    Public PrimaryCurrAkunAR As Double = -1
    Public PrimaryCurrAkunReff As Double = -1
    Public PrimaryJournal As Double = -1
    Public PrimaryCustSupp As Double = -1
    Dim ObjCust As New CustomSMESS.ClsTransaksiJournal

    Dim ObjCDNote As New ClsTransaksiCreditNote
    Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping

    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                If Me.TipeKartu = ClsTransaksiCreditNote.EnumTipeKartu.Customer Then
                    Me.Text = "Customer's Credit Note"
                Else
                    Me.Text = "Supplier's Credit Note"
                End If
            Else
                If Me.TipeKartu = ClsTransaksiCreditNote.EnumTipeKartu.Customer Then
                    Me.Text = "Customer's Debit Note"
                Else
                    Me.Text = "Supplier's Debit Note"
                End If
            End If

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint) = "1" Then
                ObjCust.CetakCDNote(Me.PrimaryJournal, Me.crvReport, Me.TipeTransaksi, Me.TipeKartu, Me.PrimaryCustSupp, Me.PrimaryCurrAkunAR, Me.PrimaryCurrAkunReff)
            Else
                ObjCDNote.CetakCDNote(Me.PrimaryJournal, Me.crvReport, Me.TipeTransaksi, Me.TipeKartu, Me.PrimaryCustSupp, Me.PrimaryCurrAkunAR, Me.PrimaryCurrAkunReff)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class