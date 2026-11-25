Option Explicit On
Option Strict Off

Public Class FrmTransaksiSO_ReminderAR
    Public vPrimaryCustomer As Double
    Public vDate As Date = Today
    Public vCreditLimit As Decimal = 0
    Public vTipeReminder As EnumTipeReminder = EnumTipeReminder.AR_Exceeds_Limit
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Tolerance_AROverdue As Decimal = 0
    Public Tolerance_ARCreditLimit As Decimal = 0
    Public FormAcuan As EnumFormAcuan = EnumFormAcuan.SO
    Public vUserCode As String = ""
    Public ObjOtorisasiDO_ExceedCreditLimit As Core_Module.ClsMasterTipeOtorisasi
    Public ObjOtorisasiDO_OverdueInvoice As Core_Module.ClsMasterTipeOtorisasi

    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    Public Enum EnumFormAcuan As Integer
        SO
        _DO
    End Enum

    Public Enum EnumTipeReminder As Integer
        AR_Exceeds_Limit
        AR_Overdue
    End Enum

    Private Sub FrmTransaksiSO_ReminderAR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.vTipeReminder = EnumTipeReminder.AR_Overdue Then
                Me.ObjSO.GetAR_Detail_OverDue(Me.vPrimaryCustomer, Me.vDate, Me)
                Me.Label1.Text = "This customer has overdue account receivable on: " + Format(Me.vDate, "dd MMM yyyy")
            Else
                Me.ObjSO.GetAR_Detail_ExceedsLimit(Me.vPrimaryCustomer, Me)
                Me.Label1.Text = "This customer exceeds credit limit: IDR. " + _
                Me.ObjCD.DigitGrouping(vCreditLimit, Me.ObjCurrency.GetDecimal_DefaultCurrency) + _
                " - Tolerance: IDR " + Me.ObjCD.DigitGrouping(Me.Tolerance_ARCreditLimit, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If Me.FormAcuan = EnumFormAcuan._DO Then
                If Me.vTipeReminder = EnumTipeReminder.AR_Exceeds_Limit Then
                    If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_Override_ARCreditLimit") = 0 Then
                        Me.btnYes.Enabled = False
                    End If
                Else
                    If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_Override_ARDue") = 0 Then
                        Me.btnYes.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click
        Dim FrmOtorisasi As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.FormAcuan = EnumFormAcuan._DO Then
                If Me.vTipeReminder = EnumTipeReminder.AR_Exceeds_Limit Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.vUserCode, "FrmTransaksiDO_Override_ARCreditLimit") = False Then
                        FrmOtorisasi.AccessName = "FrmTransaksiDO_Override_ARCreditLimit"
                        FrmOtorisasi.Text = "Approval : DO Exceed credit limit"
                        If FrmOtorisasi.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Me.DialogResult = Windows.Forms.DialogResult.Yes
                            'simpan data otorisasi ke objek
                            With Me.ObjOtorisasiDO_ExceedCreditLimit
                                .Primary = 1
                                .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmOtorisasi.cboUserCode.Text)
                                .PrimaryDokumenAcuan = -1
                                .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.vUserCode)
                                .PrimaryTipeApproval = Core_Module.ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Exceed_Credit_Limit
                                .Reason = FrmOtorisasi.TxtReason.Text
                                .NoDocumentHist = ""
                            End With
                            Me.Close()
                        End If
                    Else
                        Me.DialogResult = Windows.Forms.DialogResult.Yes
                        Me.Close()
                    End If
                ElseIf Me.vTipeReminder = EnumTipeReminder.AR_Overdue Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.vUserCode, "FrmTransaksiDO_Override_ARDue") = False Then
                        FrmOtorisasi.AccessName = "FrmTransaksiDO_Override_ARDue"
                        FrmOtorisasi.Text = "Approval : DO overdue invoice"
                        If FrmOtorisasi.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Me.DialogResult = Windows.Forms.DialogResult.Yes
                            'simpan data otorisasi ke object
                            With Me.ObjOtorisasiDO_OverdueInvoice
                                .Primary = 1
                                .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmOtorisasi.cboUserCode.Text)
                                .PrimaryDokumenAcuan = -1
                                .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.vUserCode)
                                .PrimaryTipeApproval = Core_Module.ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_Overdue_Invoice
                                .Reason = FrmOtorisasi.TxtReason.Text
                                .NoDocumentHist = ""
                            End With
                            Me.Close()
                        End If
                    Else
                        Me.DialogResult = Windows.Forms.DialogResult.Yes
                        Me.Close()
                    End If
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Yes
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try



    End Sub

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNo.Click
        Me.DialogResult = Windows.Forms.DialogResult.No
        Me.Close()
    End Sub
End Class