Option Explicit On
Option Strict Off
Imports Core_Module, AccModule

Public Class FrmTrans_Session
    'data yang tidak boleh diedit pada session jika sudah ada transaksi pos yang merujuk pada transaksi ini
    '1. kode cashier
    '2. kode cabang
    '3. ??
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public DefPrec As Double = 0
    Public RoundingTipe As Decimal
    Public RoundingValue As Decimal

    Dim Obj As New ClsTrans_Session
    Dim ObjCD As New ClsCekData
    Dim ObjBranch As New ClsMast_Cabang
    Dim ObjAcc As New ClsMasterAkun
    Dim ObjUser As New ClsUser
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjItem As New ClsTrans_Session_Item

    Dim Baru As Boolean
    Dim buffer As String = ""
    Dim DefaultClock As DateTime = "1950/01/01" + " " + Now.Hour.ToString + ":" + Now.Minute.ToString + ":" + Now.Second.ToString
    Dim TBLItemPaymentDetails As New DseTrans_Session.DGVPaymentDetailsDataTable

    Private Sub FrmTrans_Session_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DefPrec = ObjCur.GetDefaultPrecission(ObjGeneralSet.GetPrimaryCurr())
            If ObjGeneralSet.GetAllowRounding = 0 Then
                Me.RoundingTipe = ClsCekData.EnumTipeRounding.Middlepoint
                Me.RoundingValue = 1
            ElseIf ObjGeneralSet.GetAllowRounding = 1 Then
                Me.RoundingTipe = ObjGeneralSet.GetRoundingTipe
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.A_Seribu Then
                    Me.RoundingValue = 1000
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.B_LimaRatus Then
                    Me.RoundingValue = 500
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.C_Seratus Then
                    Me.RoundingValue = 100
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.D_LimaPuluh Then
                    Me.RoundingValue = 50
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.E_Sepuluh Then
                    Me.RoundingValue = 10
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.F_Lima Then
                    Me.RoundingValue = 5
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.G_Satu Then
                    Me.RoundingValue = 1
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.H_NolKomaLima Then
                    Me.RoundingValue = 0.5
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.I_NolKomaSatu Then
                    Me.RoundingValue = 0.1
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.J_NolKomaNolLima Then
                    Me.RoundingValue = 0.05
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.K_NolKomaNolSatu Then
                    Me.RoundingValue = 0.01
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.L_NolKomaNolNolLima Then
                    Me.RoundingValue = 0.005
                End If
                If ObjGeneralSet.GetRoundingvalue = ClsTrans_POS_GeneralSetting.EnumRoundingValue.M_NolKomaNolNolSatu Then
                    Me.RoundingValue = 0.001
                End If
            End If
            '-----
            Me.cboStatus.SelectedIndex = 0
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            Me.dgvAccount_OpSes.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAccount_CloseSes.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_DetailTrans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_cancelled_Trans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_Payment_Sum.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_PaymentD.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_PayOutS.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgv_CancelledItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVPendingTrans.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboStatus.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tco1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tbpDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tbpDetailPayment.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tbpPayout_Dropcash.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tbpCancellation.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.tcoTrans_Details.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage5.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage6.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage7.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)

            Me.CancelButton = Me.btnCancel
            Me.chk_PaymentS.Enabled = True
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tco1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tbpDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tbpDetailPayment.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tbpPayout_Dropcash.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tbpCancellation.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tcoTrans_Details.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage6.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage7.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc_Payment.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage5.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.chk_PaymentS.Checked = False
            Me.BackColorDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabControl2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage3.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage5.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage6.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage7.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tco1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tbpDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tbpDetailPayment.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tbpPayout_Dropcash.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tbpCancellation.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tcoTrans_Details.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabControl2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage3.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage5.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage6.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage7.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tco1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tbpDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tbpDetailPayment.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tbpPayout_Dropcash.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tbpCancellation.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tcoTrans_Details.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc_Payment.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabControl1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabControl2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage3.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage4.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage5.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage6.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage7.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tco1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tbpDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tbpDetailPayment.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tbpPayout_Dropcash.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tbpCancellation.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tcoTrans_Details.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc_Payment.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc_Payment.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc_Payment.Panel2.Controls, True, True, True)

            Me.TxtPrimary.Text = "-1"
            Me.TxtCash_Beginning.Text = "0"
            Me.dtpSessionPeriod.Value = Now
            Me.DtpSessionThrough.Value = Now
            Me.TxtPrimaryCashier.Text = "-1"
            Me.TxtPrimaryBranch.Text = "-1"
            Me.TxtPrimaryBeginingCashAccount.Text = "-1"
            Me.TxtPrimaryCashierCashAccount.Text = "-1"
            Me.TxtPrimaryBalanceCashAccount.Text = "-1"
            Me.Txt_TransDetail_TotalTransQty.Text = "0"
            Me.TxtSalesItemQty.Text = "0"
            Me.TxtLoginPeriod.Text = "00:00:00"
            Me.TxtInputPeriod.Text = "00:00:00"
            Me.TxtAvgPeriodTrans.Text = "00:00:00"
            Me.TxtAvgPeriodItem.Text = "00:00:00"
            Me.TxtCash_Beginning.Text = "0"
            Me.TxtCash_Sales.Text = "0"
            Me.TxtDropCash.Text = "0"
            Me.TxtPayouts.Text = "0"
            Me.TxtBalance.Text = "0"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ListPrimaryAkun As New List(Of Double)
        Dim ijk As New FrmTrans_Session_OpenClose
        Dim StatusLama As ClsTrans_Session.EnumStatusSession
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Session ID")
            KoleksiTextBox.Add(Me.TxtSessionID)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            If Me.TxtPrimaryBranch.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input branch code", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan branch code", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtKodeBranch.Focus()
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.Baru = False Then
                If Me.Obj.StartDate.Month = Me.dtpSessionPeriod.Value.Month And _
                Me.Obj.StartDate.Year = Me.dtpSessionPeriod.Value.Year And _
                Me.TxtPrimaryBranch.Text = Me.Obj.PrimaryBranch Then
                    Me.TxtSessionID.Text = Me.Obj.SessionID
                Else
                    Dim NoTransaksiLama As String = Obj.GetSessionID(Me.TxtPrimary.Text)
                    If NoTransaksiLama <> Me.TxtSessionID.Text Then
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpSessionPeriod.Value = Me.Obj.StartDate
                            Me.DtpSessionThrough.Value = Me.Obj.EndDate
                            Me.TxtPrimaryBranch.Text = Me.Obj.PrimaryBranch
                            Me.TxtKodeBranch.Text = ObjBranch.GetCode(Me.TxtPrimaryBranch.Text)
                            Me.TxtSessionID.Text = Me.Obj.SessionID
                            Return False
                        End If
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If Bahasa = EnumBahasa.Inggris Then
                        If MsgBox("Other user have modified this data, reload latest data?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me, Me.TBLItemPaymentDetails)
                            Me.SalinKeField()
                        End If
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        If MsgBox("Beberapa data telah dimodifikasi, apakah mau mengambil kembali data terakhir?", _
                        MsgBoxStyle.YesNo, "Konfirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me, Me.TBLItemPaymentDetails)
                            Me.SalinKeField()
                        End If
                    End If
                    Return False
                End If
            End If

            'Desc harus diisi
            If String.IsNullOrEmpty(Me.TxtSessionID.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input session id", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan session id", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtSessionID.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtCashierCode.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input cashier name", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan cashier name", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCode.Focus()
                Return False
            End If

            If Me.TxtCash_Beginning.Text = "0" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input beginning cash", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan beginning cash", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCash_Beginning.Focus()
                Return False
            End If
            If Me.TxtPrimaryBeginingCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input beginning cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan beginning cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBeginingCashAccountCode.Focus()
                Return False
            End If
            If Me.TxtPrimaryCashierCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input cashier cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan cashier cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCashAccountCode.Focus()
                Return False
            End If
            If Me.TxtPrimaryBalanceCashAccount.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input balance cash account", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan balance cash account", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBalanceCashAccountCode.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtRemarks.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
                    Return False
                End If
            End If

            'Cek Bilangan Positif
            If ObjCD.CekBilanganPositif(Me.TxtCash_Beginning.Text) = False Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCash_Beginning.Focus()
                Return False
            End If

            'Cek kode akun
            If ListPrimaryAkun.Contains(Me.TxtPrimaryCashierCashAccount.Text) = False Then
                ListPrimaryAkun.Add(Me.TxtPrimaryCashierCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Cashier cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Cashier cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtCashierCashAccountCode.Focus()
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.TxtPrimaryBeginingCashAccount.Text) = False Then
                ''ListPrimaryAkun.Add(Me.TxtPrimaryBeginingCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Beginning cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Beginning cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBeginingCashAccountCode.Focus()
                Return False
            End If

            If ListPrimaryAkun.Contains(Me.TxtPrimaryBalanceCashAccount.Text) = False Then
                ''ListPrimaryAkun.Add(Me.TxtPrimaryBalanceCashAccount.Text)
            Else
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Balance cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Balance cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtBalanceCashAccountCode.Focus()
                Return False
            End If

            'Cek pada saat open close session 
            ''MsgBox("Session tidak bisa ditutup jika masih ada transaksi pos yang tidak memiliki transaksi payment.")
            If Me.Baru = False Then
                If Obj.GetJmlPOSBelumPaid(Me.TxtPrimary.Text) > 0 Then
                    If Obj.GetJumlahUntukValidasiSession(Me.TxtPrimary.Text) > 0 Then
                        'Cek Status Session Sebelum di save
                        StatusLama = Obj.GetStatus(Me.TxtPrimary.Text)
                        If StatusLama <> Me.cboStatus.SelectedIndex Then
                            If Me.cboStatus.SelectedIndex = ClsTrans_Session.EnumStatusSession.Open Then
                                MsgBox("Session tidak bisa dibuka jika data session sudah ditransfer ke server pusat.", MsgBoxStyle.Information, "Pending Belum Dikerjain")
                            Else
                                ijk.UserCode = Me.UserCode
                                ijk.BackColorAll = Me.BackColorAll
                                ijk.Bahasa = Me.Bahasa
                                ijk.mParentFormSession = Me
                                ijk.mPrimarySession = Me.TxtPrimary.Text
                                ijk.btnPayment.Focus()
                                ijk.ShowDialog()

                                'load ulang data
                                Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, Me, Me.TBLItemPaymentDetails)
                                Me.SalinKeField()
                                Me.cboStatus.SelectedIndex = 1
                                If ijk.mrOK = False Then
                                    MsgBox("There is some transactions POS unpaid, you can't closed this transaction", MsgBoxStyle.Information, "Information")
                                    Me.cboStatus.SelectedIndex = 0
                                    Me.DGVPendingTrans.Rows.Clear()
                                    Me.ObjItem.IsiDGVPendingPOS(Me.TxtPrimary.Text, Me.DGVPendingTrans, Me, Me.ObjCD)
                                    Return False
                                End If
                            End If
                        End If
                        '-----
                    End If
                End If
            End If

            If Me.DtpSessionThrough.Value < Me.dtpSessionPeriod.Value Then
                MsgBox("Please cek session period and session through", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If CDec(Me.TxtBalance.Text) < 0 Then
                MsgBox("Transaction session cannot saved, because balance cannot be negative ", MsgBoxStyle.Information, "Information")
                Me.TxtCash_Beginning.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Me.DGVPendingTrans.Rows.Clear()
            Me.ObjItem.IsiDGVPendingPOS(Me.TxtPrimary.Text, Me.DGVPendingTrans, Me, Me.ObjCD)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ListPrimaryAkun = Nothing
            ijk.Dispose()
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .SessionID = Me.TxtSessionID.Text
                .StartDate = Me.dtpSessionPeriod.Value.Date + " " + Me.dtpSessionPeriod.Value.Hour.ToString + ":" + Me.dtpSessionPeriod.Value.Minute.ToString + ":" + Me.dtpSessionPeriod.Value.Second.ToString
                .EndDate = Me.DtpSessionThrough.Value.Date + " " + Me.DtpSessionThrough.Value.Hour.ToString + ":" + Me.DtpSessionThrough.Value.Minute.ToString + ":" + Me.DtpSessionThrough.Value.Second.ToString
                .Remarks = Me.TxtRemarks.Text
                .Status = Me.cboStatus.SelectedIndex
                .CashierCode = Me.TxtCashierCode.Text
                .PrimaryBranch = Me.TxtPrimaryBranch.Text
                .CashSales = Me.TxtCash_Sales.Text
                .DropCash = Me.TxtDropCash.Text
                .Payout = Me.TxtPayouts.Text
                .Balance = Me.TxtBalance.Text
                .BeginningCash = Me.TxtCash_Beginning.Text
                .PrimaryBeginningCA = Me.TxtPrimaryBeginingCashAccount.Text
                .PrimaryCashierCA = Me.TxtPrimaryCashierCashAccount.Text
                .PrimaryBalanceCA = Me.TxtPrimaryBalanceCashAccount.Text
                'berkaitan dengan transfer data
                .PrimaryTransServer = -1
                .PrimaryJurnalOpenSession = -1
                .PrimaryJurnalCloseSession = -1
                .StatusTransfer = ClsTrans_Session.EnumStatusTransfer.BelumTransfer
                .StartTransferTime = DefaultClock
                .EndTransferTime = DefaultClock
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtSessionID.Text = .SessionID
                Me.dtpSessionPeriod.Value = Format(.StartDate, "dd MMM yyyy") + " " + .StartDate.Hour.ToString + ":" + .StartDate.Minute.ToString
                Me.DtpSessionThrough.Value = Format(.EndDate, "dd MMM yyyy") + " " + .EndDate.Hour.ToString + ":" + .EndDate.Minute.ToString
                Me.TxtRemarks.Text = .Remarks
                If .Status = 0 Then
                    Me.cboStatus.SelectedIndex = 0
                Else
                    Me.cboStatus.SelectedIndex = 1
                End If
                Me.TxtPrimaryCashier.Text = "-1"
                Me.TxtCashierCode.Text = .CashierCode
                Me.TxtCashierName.Text = ObjUser.GetNama(.CashierCode)
                Me.TxtPrimaryBranch.Text = .PrimaryBranch
                Me.TxtKodeBranch.Text = ObjBranch.GetCode(.PrimaryBranch)
                Me.TxtKodeBranchDesc.Text = ObjBranch.GetDesc(.PrimaryBranch)
                Me.TxtCash_Beginning.Text = ObjCD.DigitGrouping(.BeginningCash, DefPrec)
                Me.TxtPrimaryBeginingCashAccount.Text = .PrimaryBeginningCA
                Me.TxtBeginingCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBeginningCA)
                Me.TxtBeginingCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Me.TxtPrimaryCashierCashAccount.Text = .PrimaryCashierCA
                Me.TxtCashierCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryCashierCA)
                Me.TxtCashierCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Me.TxtPrimaryBalanceCashAccount.Text = .PrimaryBalanceCA
                Me.TxtBalanceCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBalanceCA)
                Me.TxtBalanceCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBalanceCashAccountCode.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtSessionID.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.TxtSessionID.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, Me, Me.TBLItemPaymentDetails)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_PaymentS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_PaymentS.CheckedChanged
        Try
            If Me.chk_PaymentS.Checked = True Then
                Me.spc_Payment.Panel2Collapsed = False
            Else
                Me.spc_Payment.Panel2Collapsed = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeFieldMasterSession()
        Dim ijk As New FrmMast_Session_Cari
        Dim ObjMasterSession As New ClsMast_Session
        Dim i As New FrmMast_Session
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If ObjMasterSession.CariDetail(ijk.mPrimary, Me.ObjCD, i) = True Then
                    With ObjMasterSession
                        Me.dtpSessionPeriod.Value = Format(Now, "dd MMM yyyy") + " " + Now.Hour.ToString + ":" + Now.Minute.ToString
                        Me.DtpSessionThrough.Value = Format(Now, "dd MMM yyyy") + " " + Now.Hour.ToString + ":" + Now.Minute.ToString
                        Me.TxtPrimaryBranch.Text = .PrimaryBranch
                        Me.TxtKodeBranch.Text = ObjBranch.GetCode(.PrimaryBranch)
                        Me.TxtKodeBranchDesc.Text = ObjBranch.GetDesc(.PrimaryBranch)
                        Me.TxtSessionID.Text = Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                        Me.TxtPrimaryCashier.Text = "-1"
                        Me.TxtCashierCode.Text = .CashierCode
                        Me.TxtCashierName.Text = ObjUser.GetNama(.CashierCode)
                        Me.TxtCash_Beginning.Text = ObjCD.DigitGrouping(.BeginningCash, DefPrec)
                        Me.TxtPrimaryBeginingCashAccount.Text = .PrimaryBeginningCashAccount
                        Me.TxtBeginingCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBeginningCashAccount)
                        Me.TxtBeginingCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                        Me.TxtPrimaryCashierCashAccount.Text = .PrimaryCashierCashAccount
                        Me.TxtCashierCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryCashierCashAccount)
                        Me.TxtCashierCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtCashierCashAccountCode.Text)
                        Me.TxtPrimaryBalanceCashAccount.Text = .PrimaryBalanceCashAccount
                        Me.TxtBalanceCashAccountCode.Text = ObjAcc.GetAccountCode(.PrimaryBalanceCashAccount)
                        Me.TxtBalanceCashAccountDesc.Text = ObjAcc.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                        Me.TxtRemarks.Text = .Remarks
                    End With
                    Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                    Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                    Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.dtpSessionPeriod.Focus()
                Else
                    Me.Baru = True
                    Obj.ResetObject(Me)
                    Me.AktifkanDetail()
                    Me.KosongkanField()
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                    Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                    Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                    Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.dtpSessionPeriod.Focus()
                End If
            Else
                Me.Baru = True
                Obj.ResetObject(Me)
                Me.AktifkanDetail()
                Me.KosongkanField()
                Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                Me.dtpSessionPeriod.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
            i.Dispose()
            ObjMasterSession = Nothing
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim Hasil As MsgBoxResult
        Try
            'jika new, maka tawarkan untuk copy dari master session yang sudah ada
            'MsgBox("tawarkan untuk copy data dari master session yang sudah ada sebelumnya.")
            Hasil = MsgBox("Do you want to input data from the master session?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
                Me.Baru = True
                Obj.ResetObject(Me)
                Me.AktifkanDetail()
                Me.KosongkanField()
                Me.SalinKeFieldMasterSession()
                Me.cboStatus.Enabled = False
                Me.cboStatus.SelectedIndex = 0
            Else
                Me.Baru = True
                Obj.ResetObject(Me)
                Me.AktifkanDetail()
                Me.KosongkanField()
                Me.cboStatus.Enabled = False
                Me.cboStatus.SelectedIndex = 0
                Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                Me.dtpSessionPeriod.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmTransaksiSessionCari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me, Me.TBLItemPaymentDetails) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtSessionID.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSessionID.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True, Bahasa) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    Me.dtpSessionPeriod.Focus()
                End If
            End If

            'jika sudah ada data transaksi pos yang merujuk pada transaksi ini, ada data2 yang tidak bisa diedit:
            '1. data cabang --> karena akan mempengaruhi perhitungan diskon
            '2. data cashier code
            '3.  
            'jika status transfer, sudah mulai transfer, maka tidak boleh diedit, terhitung 10 menit sejak dimulainya proses transfer.
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtSessionID.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSessionID.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                 Bahasa, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                        Bahasa, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.dtpSessionPeriod.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeBranch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtKodeBranch.KeyDown
        Dim objCode As New ClsMast_Cabang
        Dim i As New FrmMast_Cabang_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtKodeBranch.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtKodeBranch.Text) = True Then
                        Me.TxtPrimaryBranch.Text = objCode.GetPrimary(Me.TxtKodeBranch.Text)
                        Me.TxtKodeBranch.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                        Me.TxtKodeBranchDesc.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                        Me.TxtRemarks.Focus()
                        '---
                        If Me.TxtPrimaryBranch.Text <> "-1" Then
                            If Me.Baru = True Then
                                Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                            Else
                                If Me.TxtPrimaryBranch.Text = Obj.GetPrimaryBranch(Me.TxtPrimary.Text) Then
                                    Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                                Else
                                    If Me.TxtPrimaryBranch.Text <> Obj.PrimaryBranch Then
                                        Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                                    Else
                                        Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                                    End If
                                End If
                            End If
                        Else
                            If Me.Baru = False Then
                                Me.TxtSessionID.Text = Obj.GetSessionID(Me.TxtPrimary.Text)
                            Else
                                Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                            End If
                        End If
                    Else
                        Me.BtnFindBranch.PerformClick()
                    End If
                Else
                    Me.BtnFindBranch.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtKodeBranch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtKodeBranch.Leave
        Dim objCode As New ClsMast_Cabang
        Dim i As New FrmMast_Cabang_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtKodeBranch.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtKodeBranch.Text) = True Then
                    Me.TxtPrimaryBranch.Text = objCode.GetPrimary(Me.TxtKodeBranch.Text)
                    Me.TxtKodeBranch.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                    Me.TxtKodeBranchDesc.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                Else
                    If Me.TxtKodeBranch.Text <> "" Then
                        i.txtCari.Text = Me.TxtKodeBranch.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        Me.TxtPrimaryBranch.Text = i.mPrimary
                        Me.TxtKodeBranch.Text = objCode.GetCode(Me.TxtPrimaryBranch.Text)
                        Me.TxtKodeBranchDesc.Text = objCode.GetDesc(Me.TxtPrimaryBranch.Text)
                    Else
                        Me.TxtKodeBranch.Clear()
                        Me.TxtKodeBranchDesc.Clear()
                        Me.TxtPrimaryBranch.Text = "-1"
                        Me.TxtKodeBranch.Focus()
                    End If
                End If
            Else
                Me.TxtKodeBranch.Clear()
                Me.TxtKodeBranchDesc.Clear()
                Me.TxtPrimaryBranch.Text = "-1"
            End If
            '---
            If Me.TxtPrimaryBranch.Text <> "-1" Then
                If Me.Baru = True Then
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                Else
                    If Me.TxtPrimaryBranch.Text = Obj.GetPrimaryBranch(Me.TxtPrimary.Text) Then
                        Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                    Else
                        If Me.TxtPrimaryBranch.Text <> Obj.PrimaryBranch Then
                            Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                        Else
                            Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
            Else
                If Me.Baru = False Then
                    Me.TxtSessionID.Text = Obj.GetSessionID(Me.TxtPrimary.Text)
                Else
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBranchCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindBranch.Click
        Dim i As New FrmMast_Cabang_Cari
        Dim ObjCode As New ClsMast_Cabang
        i.BackColorAll = Me.BackColorAll
        Try
            'kalo cabangnya di edit, maka hanya bagian cabangnya aja yang berubah
            If Me.TxtKodeBranch.Text <> "" Then
                i.txtCari.Text = Me.TxtKodeBranch.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then
                Me.TxtPrimaryBranch.Text = i.mPrimary
                Me.TxtKodeBranch.Text = ObjCode.GetCode(Me.TxtPrimaryBranch.Text)
                Me.TxtKodeBranchDesc.Text = ObjCode.GetDesc(Me.TxtPrimaryBranch.Text)
                Me.TxtKodeBranch.Focus()
            Else
                Me.TxtKodeBranch.Clear()
                Me.TxtKodeBranchDesc.Clear()
                Me.TxtPrimaryBranch.Text = "-1"
                Me.TxtKodeBranch.Focus()
            End If
            '---
            If Me.TxtPrimaryBranch.Text <> "-1" Then
                If Me.Baru = True Then
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                Else
                    If Me.TxtPrimaryBranch.Text = Obj.GetPrimaryBranch(Me.TxtPrimary.Text) Then
                        Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                    Else
                        If Me.TxtPrimaryBranch.Text <> Obj.PrimaryBranch Then
                            Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                        Else
                            Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
            Else
                If Me.Baru = False Then
                    Me.TxtSessionID.Text = Obj.GetSessionID(Me.TxtPrimary.Text)
                Else
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBeginingCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
       Handles TxtBeginingCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtBeginingCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtBeginingCashAccountCode.Text) = True Then
                        Me.TxtPrimaryBeginingCashAccount.Text = objCode.GetPrimary(Me.TxtBeginingCashAccountCode.Text)
                        Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                        Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                        Me.TxtCashierCashAccountCode.Focus()
                    Else
                        Me.btnFindBeginingCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindBeginingCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBeginingCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtBeginingCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBeginingCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtBeginingCashAccountCode.Text) = True Then
                    Me.TxtPrimaryBeginingCashAccount.Text = objCode.GetPrimary(Me.TxtBeginingCashAccountCode.Text)
                    Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                    Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Else
                    If Me.TxtBeginingCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtBeginingCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryBeginingCashAccount.Text = i.mPrimaryAkun
                        Me.TxtBeginingCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                        Me.TxtBeginingCashAccountDesc.Text = objCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                    Else
                        Me.TxtBeginingCashAccountCode.Clear()
                        Me.TxtBeginingCashAccountDesc.Clear()
                        Me.TxtPrimaryBeginingCashAccount.Text = "-1"
                        Me.TxtBeginingCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtBeginingCashAccountCode.Clear()
                Me.TxtBeginingCashAccountDesc.Clear()
                Me.TxtPrimaryBeginingCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBeginningCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindBeginingCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBeginingCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtBeginingCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryBeginingCashAccount.Text = i.mPrimaryAkun
                Me.TxtBeginingCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryBeginingCashAccount.Text)
                Me.TxtBeginingCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtBeginingCashAccountCode.Text)
                Me.TxtCashierCashAccountCode.Focus()
            Else
                Me.TxtBeginingCashAccountCode.Clear()
                Me.TxtBeginingCashAccountDesc.Clear()
                Me.TxtPrimaryBeginingCashAccount.Text = "-1"
                Me.TxtBeginingCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtCashierCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashierCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCashierCashAccountCode.Text) = True Then
                        Me.TxtPrimaryCashierCashAccount.Text = objCode.GetPrimary(Me.TxtCashierCashAccountCode.Text)
                        Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                        Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                        Me.TxtBalanceCashAccountCode.Focus()
                    Else
                        Me.btnFindCashierCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindCashierCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashierCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCashierCashAccountCode.Text) = True Then
                    Me.TxtPrimaryCashierCashAccount.Text = objCode.GetPrimary(Me.TxtCashierCashAccountCode.Text)
                    Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                    Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Else
                    If Me.TxtCashierCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCashierCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCashierCashAccount.Text = i.mPrimaryAkun
                        Me.TxtCashierCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                        Me.TxtCashierCashAccountDesc.Text = objCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                    Else
                        Me.TxtCashierCashAccountCode.Clear()
                        Me.TxtCashierCashAccountDesc.Clear()
                        Me.TxtPrimaryCashierCashAccount.Text = "-1"
                        Me.TxtCashierCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashierCashAccountCode.Clear()
                Me.TxtCashierCashAccountDesc.Clear()
                Me.TxtPrimaryCashierCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindCashierCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindCashierCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCashierCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCashierCashAccount.Text = i.mPrimaryAkun
                Me.TxtCashierCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryCashierCashAccount.Text)
                Me.TxtCashierCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtCashierCashAccountCode.Text)
                Me.TxtBalanceCashAccountCode.Focus()
            Else
                Me.TxtCashierCashAccountCode.Clear()
                Me.TxtCashierCashAccountDesc.Clear()
                Me.TxtPrimaryCashierCashAccount.Text = "-1"
                Me.TxtCashierCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBalanceCashAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtBalanceCashAccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtBalanceCashAccountCode.Text) = True Then
                        Me.TxtPrimaryBalanceCashAccount.Text = objCode.GetPrimary(Me.TxtBalanceCashAccountCode.Text)
                        Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                        Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.btnFindBalanceCashAccount.PerformClick()
                    End If
                Else
                    Me.btnFindBalanceCashAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtBalanceCashAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBalanceCashAccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtBalanceCashAccountCode.Text) = True Then
                    Me.TxtPrimaryBalanceCashAccount.Text = objCode.GetPrimary(Me.TxtBalanceCashAccountCode.Text)
                    Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                    Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                Else
                    If Me.TxtBalanceCashAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtBalanceCashAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryBalanceCashAccount.Text = i.mPrimaryAkun
                        Me.TxtBalanceCashAccountCode.Text = objCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                        Me.TxtBalanceCashAccountDesc.Text = objCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                    Else
                        Me.TxtBalanceCashAccountCode.Clear()
                        Me.TxtBalanceCashAccountDesc.Clear()
                        Me.TxtPrimaryBalanceCashAccount.Text = "-1"
                        Me.TxtBalanceCashAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtBalanceCashAccountCode.Clear()
                Me.TxtBalanceCashAccountDesc.Clear()
                Me.TxtPrimaryBalanceCashAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindBalanceCashAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindBalanceCashAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtBalanceCashAccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtBalanceCashAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryBalanceCashAccount.Text = i.mPrimaryAkun
                Me.TxtBalanceCashAccountCode.Text = ObjCode.GetAccountCode(Me.TxtPrimaryBalanceCashAccount.Text)
                Me.TxtBalanceCashAccountDesc.Text = ObjCode.GetDescription(Me.TxtBalanceCashAccountCode.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtBalanceCashAccountCode.Clear()
                Me.TxtBalanceCashAccountDesc.Clear()
                Me.TxtPrimaryBalanceCashAccount.Text = "-1"
                Me.TxtBalanceCashAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCash_Beginning_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.Enter
        Try
            ObjCD.Number_Enter(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCash_Beginning.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCash_Beginning.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtCash_Beginning, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCash_Beginning_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.Leave
        Try
            ObjCD.Number_Leave(Me.TxtCash_Beginning, sender, e, Me.DefPrec, ObjCD)
            Me.Obj.PerhitunganBalance(Me.TxtPrimary.Text, CDec(Me.TxtCash_Beginning.Text), Me, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCashierCode.KeyDown
        Dim objCode As New ClsUser
        Dim i As New FrmCariMasterUser
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCashierCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCashierCode.Text) = True Then
                        Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                        Me.TxtKodeBranch.Focus()
                    Else
                        Me.btnFindCashier.PerformClick()
                    End If
                Else
                    Me.btnFindCashier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCashierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCashierCode.Leave
        Dim objCode As New ClsUser
        Dim i As New FrmCariMasterUser
        'i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCashierCode.Text) = True Then
                    Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                Else
                    If Me.TxtCashierCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCashierCode.Text
                    End If

                    i.ShowDialog()
                    If String.IsNullOrEmpty(i.Kodeuser) = False Then
                        Me.TxtCashierCode.Text = i.Kodeuser
                        Me.TxtCashierName.Text = objCode.GetNama(Me.TxtCashierCode.Text)
                    Else
                        Me.TxtCashierCode.Clear()
                        Me.TxtCashierName.Clear()
                        Me.TxtPrimaryCashier.Text = "-1"
                        Me.TxtCashierCode.Focus()
                    End If
                End If
            Else
                Me.TxtCashierCode.Clear()
                Me.TxtCashierName.Clear()
                Me.TxtPrimaryCashier.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindCashier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCashier.Click
        Dim i As New FrmCariMasterUser
        Dim ObjCode As New ClsUser
        'i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCashierCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtCashierCode.Text
            End If

            i.ShowDialog()
            If String.IsNullOrEmpty(i.Kodeuser) = False Then
                Me.TxtCashierCode.Text = i.Kodeuser
                Me.TxtCashierName.Text = ObjCode.GetNama(Me.TxtCashierCode.Text)
                Me.TxtKodeBranch.Focus()
            Else
                Me.TxtCashierCode.Clear()
                Me.TxtCashierName.Clear()
                Me.TxtPrimaryCashier.Text = "-1"
                Me.TxtCashierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dtpSessionPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpSessionPeriod.Leave
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Coding BackColor
    Private Sub TxtSessionID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSessionID.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSessionID.Enabled = True Then
                    If Me.TxtSessionID.ReadOnly = False Then
                        TxtSessionID.BackColor = Me.BackColorAll
                    Else
                        TxtSessionID.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSessionID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSessionID.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtSessionID.Enabled = True Then
                    If Me.TxtSessionID.ReadOnly = False Then
                        TxtSessionID.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtSessionID.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCode.Enabled = True Then
                    If Me.TxtCashierCode.ReadOnly = False Then
                        TxtCashierCode.BackColor = Me.BackColorAll
                    Else
                        TxtCashierCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCashierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCode.Enabled = True Then
                    If Me.TxtCashierCode.ReadOnly = False Then
                        TxtCashierCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCashierCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                Else
                    TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeBranch.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtKodeBranch.Enabled = True Then
                    If Me.TxtKodeBranch.ReadOnly = False Then
                        Me.TxtKodeBranch.BackColor = Me.BackColorAll
                    Else
                        Me.TxtKodeBranch.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeBranch.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtKodeBranch.Enabled = True Then
                    If Me.TxtKodeBranch.ReadOnly = False Then
                        Me.TxtKodeBranch.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtKodeBranch.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCash_Beginning.Enabled = True Then
                    If Me.TxtCash_Beginning.ReadOnly = False Then
                        Me.TxtCash_Beginning.BackColor = Me.BackColorAll
                    Else
                        Me.TxtCash_Beginning.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCash_Beginning.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCash_Beginning.Enabled = True Then
                    If Me.TxtCash_Beginning.ReadOnly = False Then
                        Me.TxtCash_Beginning.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtCash_Beginning.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginingCashAccountCode.Enabled = True Then
                    If Me.TxtBeginingCashAccountCode.ReadOnly = False Then
                        Me.TxtBeginingCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginingCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginingCashAccountCode.Enabled = True Then
                    If Me.TxtBeginingCashAccountCode.ReadOnly = False Then
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBeginingCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCashAccountCode.Enabled = True Then
                    If Me.TxtCashierCashAccountCode.ReadOnly = False Then
                        Me.TxtCashierCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTerminalCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCashierCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtCashierCashAccountCode.Enabled = True Then
                    If Me.TxtCashierCashAccountCode.ReadOnly = False Then
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtCashierCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalanceCashAccountCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalanceCashAccountCode.Enabled = True Then
                    If Me.TxtBalanceCashAccountCode.ReadOnly = False Then
                        Me.TxtBalanceCashAccountCode.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalanceCashAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalanceCashAccountCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBalanceCashAccountCode.Enabled = True Then
                    If Me.TxtBalanceCashAccountCode.ReadOnly = False Then
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBalanceCashAccountCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Baru = True Then
                    Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                Else
                    If Me.dtpSessionPeriod.Value.Month = Obj.GetStartDate(Me.TxtPrimary.Text).Month And _
                    Me.dtpSessionPeriod.Value.Year = Obj.GetStartDate(Me.TxtPrimary.Text).Year Then
                        Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                    Else
                        If Me.dtpSessionPeriod.Value.Month <> Me.Obj.StartDate.Month Or _
                        Me.dtpSessionPeriod.Value.Year <> Me.Obj.StartDate.Year Then
                            Me.TxtSessionID.Text = Me.Obj.GetNoBaru(Me.dtpSessionPeriod.Value, Me.TxtPrimaryBranch.Text)
                        Else
                            Me.TxtSessionID.Text = Me.Obj.GetSessionID(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_Payment_Sum_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_Payment_Sum.CellClick
        Try
            If Me.dgv_Payment_Sum.Rows.Count > 0 Then
                ObjItem.IsiDGVPaymentDetails(Me.dgv_Payment_Sum.SelectedRows(0).Cells(Me.Col_PaymentS_PrimaryPayment.Index).Value, _
                Me, Me.ObjCD, Me.TBLItemPaymentDetails)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDGVPendingTrans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVPendingTrans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.Col_PD_ShowDetails.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDGVPendingTrans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVPendingTrans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_PD_ShowDetails.Index Then
                    iFrm.mPrimaryPOS = Me.DGVPendingTrans.Item(Me.Col_PD_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub

    Private Sub dgv_DetailTrans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv_DetailTrans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.col_TransD_ShowDetail.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_DetailTrans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_DetailTrans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.col_TransD_ShowDetail.Index Then
                    iFrm.mPrimaryPOS = Me.dgv_DetailTrans.Item(Me.Col_TransD_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub

    Private Sub dgv_cancelled_Trans_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgv_cancelled_Trans.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.Col_CancelledTrans_ShowDetails.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgv_cancelled_Trans_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_cancelled_Trans.CellClick
        Dim iFrm As New FrmTrans_Session_View
        Try
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.Col_CancelledTrans_ShowDetails.Index Then
                    iFrm.mPrimaryPOS = Me.dgv_cancelled_Trans.Item(Me.Col_CancelledTrans_PrimaryPOS.Index, e.RowIndex).Value
                    iFrm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            iFrm.Dispose()
        End Try
    End Sub

End Class