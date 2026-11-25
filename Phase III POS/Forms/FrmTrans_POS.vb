Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public Obj As New ClsTrans_POS
    Public ObjItem As New ClsTrans_POS_Item
    Public DefPrec As Double = 0
    Public PrimaryCurr As Double = -1
    Public RoundingTipe As Decimal
    Public RoundingValue As Decimal
    Public mFormPassword As FrmTrans_POS_ConfirmPassword
    Public IndexDGVDeletingRow As Integer = -1
    Public mrCloseFind As Boolean = False

    Dim FontItem_3Baris As New Font("Microsoft Sans Serif", 20, FontStyle.Regular)
    Dim FontItem_2Baris As New Font("Microsoft Sans Serif", 30, FontStyle.Regular)
    Dim FontSubTotal As New Font("Microsoft Sans Serif", 36, FontStyle.Regular)
    Dim ObjMember As New ClsMast_Member
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjDiskon As New ClsMast_Diskon
    Dim ObjSession As New ClsTrans_Session
    Dim ObjPOS As New ClsTrans_POS
    Dim ObjPOSItem As New ClsTrans_POS_Item
    Dim ObjCD As New ClsCekData
    Dim ObjProd As New clsMasterProduct
    Dim ObjLokasi As New clsMasterStorageArea
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim ObjUnit As New ClsMasterUnit
    Dim ObjUser As New ClsUser
    Dim ObjPayment As New ClsMast_PaymentOption
    Dim ObjVoid As New ClsTrans_POS_Void
    Dim MunculkanMsgBoxMember As Boolean = False

    Public TBLDiskon As New DseMast_Diskon.tblmaster_discount2DataTable
    Public TBLNext As New DseTrans_POS.TBLNextDataTable
    Dim mDeletingRow As Boolean = True
    Dim PrimarySession As Double = -1
    Dim PrimaryBranch As Double = -1
    Dim Baru As Boolean = False
    Dim Buffer As String = ""
    Dim AllowToModifyUnitPrice As Boolean = False
    Dim AllowToModifyDiscountRate As Boolean = False
    Dim InputProductDetails As Boolean = False
    'Dim PengurangHeight As Integer = 30
    Dim PengurangHeight As Integer = 112


    'format penomoran POS: XXXXXXX/POS/KodeCabang/Bulan/Tahun
    '--> Reset per tahun
    '--> Urutan Per Cabang
    Private Sub FrmTrans_POS_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'Check No Transkasi terakhir ada item nya (kalo ga ada hapus no itu)
            If ObjItem.GetJmlByPrimaryTransaksi(Me.TxtPrimary.Text) <= 0 Then
                'Ambil Primary Max, lalu dihapus 
                Obj.Hapus(Me.TxtPrimary.Text)
                'Hapus Jika ada Viod
                ObjVoid.HapusByPrimaryPOS(Me.TxtPrimary.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Hasil As MsgBoxResult
        Try
            Hasil = MsgBox("Exit POS ?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.No Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboButton.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc2.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.spc2.Controls, True, True, True, True, True, True, True, True)
            'Me.CancelButton = Me.btnCancel
            'Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            'set dgv
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToAddRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboButton.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.spc1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnClose
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToAddRows = False
            Me.KosongkanFieldItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.gboButton.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc1.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc2.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.spc2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.gboButton.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.spc2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kosongkan Field ------> ini digunakan pada saat form Load 
    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.gboButton.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.spc2.Panel2.Controls, True, True, True)
            Me.dtpTanggalTransaksi.Value = Now
            Me.DtpExpired.Value = Now
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimarySession.Text = "-1"
            Me.TxtPrimaryMember.Text = "-1"
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.TxtQty.Text = "1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'perhatikan bahwa dalam transksi ini, semua data akan langsung diupdate setiap kali ada item yang masuk ke dalam dgv item
    'ketika buat transaksi baru, maka langsung bikin transaksi baru dalam tabel database, walaupun tidak ada item-nya
    Private Sub FrmTrans_POS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.ClearSelection()
            If ObjGeneralSet.AllowToCancel() = 1 Then
                Me.dgvItem.AllowUserToDeleteRows = False
            Else
                Me.dgvItem.AllowUserToDeleteRows = True
            End If
            '-----
            Me.Obj = New ClsTrans_POS
            Me.Baru = True
            Me.mrCloseFind = False
            Me.PbxPaid.Visible = False
            Me.PbxViod.Visible = False
            Me.KosongkanField()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            '-----
            Me.PrimaryCurr = ObjGeneralSet.GetPrimaryCurr()
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

            Me.DefPrec = ObjCur.GetDecimal_DefaultCurrency_POS()
            '------
            'If ObjSession.GetJmlByCashierCode(Me.UserCode) > 1 Then
            FrmTrans_POS_Payment_Load.mCashierCode = Me.UserCode
            FrmTrans_POS_Payment_Load.ShowDialog()
            If FrmTrans_POS_Payment_Load.mExit = True Then
                FrmTrans_POS_Payment_Load.Close()
                '----NonAktifkan Detail----
                Me.NonAktifkanDetail()
                Me.KeyPreview = True
                Me.btnClose.Enabled = True
            Else
                PrimarySession = FrmTrans_POS_Payment_Load.mPrimarySession
                Me.TxtPrimarySession.Text = PrimarySession
                'If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Open Then
                Me.TxtPrimary.Text = "-1"
                Me.TxtIdNumber.Text = ObjPOS.GetNoBaru(Now, Me.TxtPrimarySession.Text)
                'Else
                '----NonAktifkan Detail----
                'Me.NonAktifkanDetail()
                'Me.KeyPreview = True
                'Me.btnFind.Enabled = True
                'Me.btnClose.Enabled = True
                'Me.btnNext.Enabled = False
                'Me.btnFind.PerformClick()
                'Me.btnFind.Focus()
                'End If
            End If
            'ElseIf ObjSession.GetJmlByCashierCode(Me.UserCode) = 1 Then
            'Me.PrimarySession = ObjSession.GetPrimaryByCashierCode(Me.UserCode)
            'Me.TxtPrimarySession.Text = Me.PrimarySession
            'If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Open Then
            'Me.TxtPrimary.Text = "-1"
            'Me.TxtIdNumber.Text = ObjPOS.GetNoBaru(Me.dtpTanggalTransaksi.Value, Me.TxtPrimarySession.Text)
            ' Else
            '----NonAktifkan Detail----
            'Me.NonAktifkanDetail()
            'Me.KeyPreview = True
            'Me.btnClose.Enabled = True
            'Me.btnFind.Enabled = True
            'Me.btnNext.Enabled = False
            'Me.btnFind.PerformClick()
            'Me.btnFind.Focus()
            'End If
            'Else
            'MsgBox("jika user ybs, memiliki 2 session yang aktif pada saat yang bersama, maka user ybs harus memilih session yang akan dirujuk pada transaksi pos yang diberikan.")
            'MsgBox("jika user ybs, tidak memiliki session yang aktif, maka user ybs akan langsung diberikan informasi dan di keluarkan ke form utama, jika punya hak akses ke form lain, atau dikeluarkan dari aplikasi jika hanya memiliki akses untuk form ini.")
            'End If

            If Me.AllowToModifyDiscountRate = False And Me.AllowToModifyUnitPrice = False And Me.InputProductDetails = False Then
                'Me.spc2.Panel2Collapsed = True
                Me.TxtUnitPrice.TabStop = False
                Me.TxtUnitPrice.Visible = False
                Me.TxtUnitPrice.ReadOnly = True
                Me.TxtNettPrice.TabStop = False
                Me.TxtNettPrice.Visible = False
                Me.TxtNettPrice.ReadOnly = True
                Me.TxtDiscString.TabStop = False
                Me.TxtDiscString.Visible = False
                Me.TxtDiscString.ReadOnly = True
                Me.TxtDiscAmount.TabStop = False
                Me.TxtDiscAmount.Visible = False
                Me.TxtDiscAmount.ReadOnly = True
                Me.LblUnitPrice.Visible = False

                Me.TxtCodeNo.Parent = Me.spc2.Panel2
                Me.LblCodeNo.Parent = Me.spc2.Panel2
                Me.btnFindCodeNo.Parent = Me.spc2.Panel2

                Me.TxtCodeNo.Top = Me.TxtUnitPrice.Top
                Me.LblCodeNo.Top = Me.LblUnitPrice.Top
                Me.btnFindCodeNo.Top = Me.TxtUnitPrice.Top

                Me.spc2.Height -= Me.PengurangHeight
                Me.dgvItem.Height += Me.PengurangHeight
                Me.dgvItem.Top -= Me.PengurangHeight
            End If

            Me.Timer1.Interval = 2000
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.TxtDisplay.Font = Me.FontItem_3Baris
            Me.TxtDisplay.Text = "KY-011 - Tempat Dupa Kecil @IDR 50.000" + vbCrLf + "Discount Member 1% IDR 500" + vbCrLf + _
            "3 pcs x IDR 49.500 = IDR 148.500"
            'Timer1.enable =false 
            'Timer1.enable =true
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Me.TxtDisplay.Font = Me.FontSubTotal
            Me.TxtDisplay.Text = "Total: IDR. 150.000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            Me.TxtDisplay.Font = Me.FontItem_2Baris
            Me.TxtDisplay.Text = "KY-011 - Tempat Dupa Kecil @IDR 50.000" + vbCrLf + _
            "3 pcs x IDR 49.500 = IDR 148.500"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPayment.Click
        Dim ij As New FrmTrans_POS_Payment
        Try
            If Me.dgvItem.Rows.Count > 0 Then
                'cek apakah sudah dibayar
                If Obj.CekApakahSudahPaid(Me.TxtPrimary.Text) = False Then
                    If Obj.GetIsVoid(Me.TxtPrimary.Text) = 0 Then
                        'Single Payment
                        ij.ParentForm_1 = Me
                        ij.Obj = Me.Obj
                        ij.cboPaymentTipe.Focus()
                        ij.ShowDialog()
                    Else
                        MsgBox("This transaction has been void", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                Else
                    MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnPayOutDropCash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPayOutDropCash.Click
        Dim ij As New FrmTrans_POS_PayoutDropCash_Input
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            ij.Parent_Form = Me
            ij.cboPayout_DropCash.Focus()
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
            ''MsgBox("Jika user hanya memiliki hak akses 1 saja ke form ini saja, maka ketika log out, ybs juga langsung keluar dari program. Tidak hanya ke main form saja.")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTrans_POS_Cari
        Try
            ij.mPrimarySession = Me.TxtPrimarySession.Text
            ij.ShowDialog()
            ij.Dispose()
            If ij.mPrimary <> -1 Then
                If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                    Me.NonAktifkanDetail()
                    Me.PbxViod.Visible = False
                    Me.PbxPaid.Visible = False
                    Obj.ResetObject(Me.dgvItem)
                    Me.KosongkanField()
                    Me.Obj.CariDetail(ij.mPrimary, Me.dgvItem, ObjCD, Me)
                    Me.SalinKeFieldUntukFind()
                    Me.btnClose.Enabled = True
                    Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                    '----
                    Me.btnFind.Enabled = True
                    Me.btnNext.Enabled = False
                    Me.KeyPreview = True
                    Me.btnFind.Focus()
                    Exit Sub
                End If
                If Obj.GetJmlPayment(ij.mPrimary) > 0 Then
                    Me.NonAktifkanDetail()
                    Me.PbxPaid.Visible = True
                    Me.PbxPaid.BringToFront()
                    Obj.ResetObject(Me.dgvItem)
                    Me.KosongkanField()
                    Me.Obj.CariDetail(ij.mPrimary, Me.dgvItem, ObjCD, Me)
                    Me.SalinKeFieldUntukFind()
                    Me.btnClose.Enabled = True
                    '----
                    Me.btnFind.Enabled = True
                    Me.btnNew.Enabled = True
                    Me.btnNext.Enabled = True
                    Me.KeyPreview = True
                    '----
                    If ObjGeneralSet.GetAllowUserToReprintReceipt() > 0 Then
                        If ObjGeneralSet.GetUseAuthorizationToReprintReceipt() > 0 Then
                            Dim FormPass As New FrmTrans_POS_ConfirmPassword
                            Me.mFormPassword = FormPass
                            Me.mFormPassword.cboUser.Focus()
                            Me.mFormPassword.ShowDialog()
                            If FormPass.mCancel = True Then
                                Me.btnReprint.Enabled = False
                            Else
                                Me.btnReprint.Enabled = True
                            End If
                            ij.Dispose()
                        Else
                            Me.btnReprint.Enabled = True
                        End If
                    Else
                        Me.btnReprint.Enabled = False
                    End If
                    Me.dgvItem.Focus()
                    Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                ElseIf Obj.GetIsVoid(ij.mPrimary) > 0 Then
                    Me.NonAktifkanDetail()
                    Me.PbxViod.Visible = True
                    Me.PbxViod.BringToFront()
                    Obj.ResetObject(Me.dgvItem)
                    Me.KosongkanField()
                    Me.Obj.CariDetail(ij.mPrimary, Me.dgvItem, ObjCD, Me)
                    Me.SalinKeFieldUntukFind()
                    Me.btnClose.Enabled = True
                    Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                    '----
                    Me.btnFind.Enabled = True
                    Me.btnNew.Enabled = True
                    Me.btnNext.Enabled = True
                    Me.KeyPreview = True
                    Me.dgvItem.Focus()
                Else
                    Me.AktifkanDetail()
                    Me.PbxPaid.Visible = False
                    Me.PbxViod.Visible = False
                    If ij.mrClose = False Then
                        If ij.mPrimary <> -1 Then
                            Obj.ResetObject(Me.dgvItem)
                            Me.KosongkanField()
                            Me.Obj.CariDetail(ij.mPrimary, Me.dgvItem, ObjCD, Me)
                            Me.SalinKeFieldUntukFind()
                            Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                        End If
                    End If
                    Me.TxtQty.Focus()
                    Me.TxtQty.SelectAll()
                End If
            Me.Baru = False
                If ij.mrClose = False Then
                    Me.mrCloseFind = False
                Else
                    Me.mrCloseFind = True
                End If
            End If
            ''MsgBox("data yang sudah memiliki item payment, tidak bisa diedit.")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.GotFocus
        Try
            'focus pada baris pertama yang terlihat
            If Me.dgvItem.Rows.Count > 0 Then
                Me.dgvItem.ClearSelection()
                Me.dgvItem.Rows(Me.dgvItem.FirstDisplayedCell.RowIndex).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.LostFocus
        Try
            Me.dgvItem.ClearSelection()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Dim Row As DataGridViewRow
        Try
            ObjCD.SetNumber(Me.dgvItem)
            If mDeletingRow = True Then
                For Each Row In Me.dgvItem.Rows
                    Me.SalinKeObjekUntukUpdateDataItem(Row)
                Next
            End If
            Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim Hasil As New MsgBoxResult
        Dim i As New FrmTrans_POS_VoidQty
        Dim TandaHapusLangung As Boolean = False
        Dim Row As DataGridViewRow
        Dim ijReason As New FrmTrans_POS_VoidQty1
        Dim KodeUserDelete As String = ""
        Dim PrimaryUserDelete As Double = -1
        Dim RowDiskon As DseMast_Diskon.tblmaster_discount2Row
        Try
            Me.mDeletingRow = True
            Me.IndexDGVDeletingRow = e.Row.Index

            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If

            If Obj.CekApakahSudahPaid(Me.TxtPrimary.Text) = True Then
                MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If

            If Obj.GetIsVoid(Me.TxtPrimary.Text) = 1 Then
                MsgBox("This transaction has been void", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If

            'jika allow user to cancel item di uncheck, maka tidak boleh hapus 
            If ObjGeneralSet.GetAllowToCancelItem <= 0 Then
                e.Cancel = True
                Exit Sub
            End If

            'yang bisa dihapuskan ada item dari tranasksi yang sessionnya masih open, jika session tidak open tidak bisa lagi dihapuskan
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) > 0 Then
                Me.mDeletingRow = False
                e.Cancel = True
                Exit Sub
            End If

            'jika pada general setting user tidak perlu otorisasi, maka bisa langsung dihapuskan
            'jika perlu otorisasi, maka keluarkan form otorisasi terlebih dahulu
            If ObjGeneralSet.GetAllowToCancelItem > 0 Then
                If ObjGeneralSet.GetUseAuthorAllowCancelItem() > 0 Then
                    Dim ij As New FrmTrans_POS_ConfirmPassword
                    Me.mFormPassword = ij
                    Me.mFormPassword.cboUser.Focus()
                    Me.mFormPassword.ShowDialog()
                    If ij.mCancel = True Then
                        Me.mDeletingRow = False
                        e.Cancel = True
                        Exit Sub
                    End If
                    KodeUserDelete = ij.cboUser.Text
                    ij.Dispose()
                End If
            Else
                e.Cancel = True
                Exit Sub
            End If

            'jika qty > 1, maka keluarkan form untuk konfirmasi qty terlebih dahulu
            If e.Row.Cells(Me.ColQty.Index).Value > 1 Then
                ''Hasil = MsgBox("Qty = " + e.Row.Cells(Me.ColQty.Index).Value.ToString + " , delete item ?", MsgBoxStyle.YesNo, "Confirm")
                i.Parent_Form = Me
                i.Text = "Enter Void Qty (Max Qty : " + e.Row.Cells(Me.ColQty.Index).Value.ToString + " )"
                i.TxtQty.Text = ObjCD.DigitGrouping(e.Row.Cells(Me.ColQty.Index).Value, Me.DefPrec, Me.RoundingTipe)
                i.TxtQty.Focus()
                i.TxtQty.SelectAll()
                i.ShowDialog()

                If i.TxtQty.Text = e.Row.Cells(Me.ColQty.Index).Value Then
                    TandaHapusLangung = True
                End If

                'Update data di DGV
                If i.mrOk = True Then
                    With e.Row
                        .Cells(Me.ColQty.Index).Value = CDec(.Cells(Me.ColQty.Index).Value) - CDec(i.TxtQty.Text)
                        .Cells(Me.ColDiscString.Index).Value = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), Me.TxtPrimaryCodeNo.Text, _
                                                            Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, CDec(.Cells(Me.ColQty.Index).Value), Me.TBLDiskon, 0)
                        .Cells(Me.ColDisc_Amount.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString.Index).Value, _
                                                             CDec((CDec(.Cells(Me.ColQty.Index).Value) - CDec(i.TxtQty.Text)) * .Cells(Me.ColUnitPrice.Index).Value), _
                                                             DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        .Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColUnitPrice.Index).Value - .Cells(Me.ColDisc_Amount.Index).Value, Me.DefPrec)
                        .Cells(Me.ColSubTotal.Index).Value = ObjCD.DigitGrouping(CDec(.Cells(Me.ColQty.Index).Value) * .Cells(Me.ColNettPrice.Index).Value, Me.DefPrec)
                        .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                                                              CDec(CDec(.Cells(Me.ColQty.Index).Value) * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                                                              CDec(CDec(.Cells(Me.ColQty.Index).Value) * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                                                              CDec(CDec(.Cells(Me.ColQty.Index).Value) * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        For IL As Integer = 0 To TBLDiskon.Rows.Count - 1
                            RowDiskon = TBLDiskon.Rows(IL)
                            If IL = 0 Then
                                .Cells(Me.ColPrimaryDisc1.Index).Value = RowDiskon.Primary
                                .Cells(Me.ColDiscString1.Index).Value = RowDiskon.Discount_Rate
                                .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                                CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            ElseIf IL = 1 Then
                                .Cells(Me.ColPrimaryDisc2.Index).Value = RowDiskon.Primary
                                .Cells(Me.ColDiscString2.Index).Value = RowDiskon.Discount_Rate
                                .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                                CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            ElseIf IL = 2 Then
                                .Cells(Me.ColPrimaryDisc3.Index).Value = RowDiskon.Primary
                                .Cells(Me.ColDiscString3.Index).Value = RowDiskon.Discount_Rate
                                .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                                CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            End If
                        Next
                    End With
                End If
                If i.mrClose = True Then
                    e.Cancel = True
                    Exit Sub
                End If
            Else
                Hasil = MsgBox("Delete item ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.No Then
                    e.Cancel = True
                    Exit Sub
                Else
                    TandaHapusLangung = True
                End If
            End If

            If String.IsNullOrEmpty(KodeUserDelete) = True Then
                PrimaryUserDelete = "-1"
            Else
                PrimaryUserDelete = ObjUser.GetPrimary(KodeUserDelete)
            End If

            'Lakukan Hapus Baris yang sama pada database
            If Me.mDeletingRow = True Then
                If TandaHapusLangung = True Then
                    '-----
                    If ObjVoid.GetJml(PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value) <= 0 Then
                        ObjVoid.SimpanItemViodSatuSatu(1, _
                        e.Row.Cells(Me.ColUnitPrice.Index).Value, e.Row.Cells(Me.ColDiscString.Index).Value, e.Row.Cells(Me.ColDisc_Amount.Index).Value, _
                        e.Row.Cells(Me.ColNettPrice.Index).Value, e.Row.Cells(Me.ColSubTotal.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc1.Index).Value, e.Row.Cells(Me.ColDiscString1.Index).Value, e.Row.Cells(Me.ColDiscAmount1.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc2.Index).Value, e.Row.Cells(Me.ColDiscString2.Index).Value, e.Row.Cells(Me.ColDiscAmount2.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc3.Index).Value, e.Row.Cells(Me.ColDiscString3.Index).Value, e.Row.Cells(Me.ColDiscAmount3.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryLocation.Index).Value, e.Row.Cells(Me.ColBatchNumber.Index).Value, e.Row.Cells(Me.ColExpiry.Index).Value, _
                        i.TxtReason.Text, e.Row.Cells(Me.ColPrimaryUnit.Index).Value, PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, _
                        e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me)
                    Else
                        'Update data void
                        ObjVoid.UpdateItemViod(CDec(i.TxtQty.Text), _
                        e.Row.Cells(Me.ColUnitPrice.Index).Value, e.Row.Cells(Me.ColDiscString.Index).Value, e.Row.Cells(Me.ColDisc_Amount.Index).Value, _
                        e.Row.Cells(Me.ColNettPrice.Index).Value, e.Row.Cells(Me.ColSubTotal.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc1.Index).Value, e.Row.Cells(Me.ColDiscString1.Index).Value, e.Row.Cells(Me.ColDiscAmount1.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc2.Index).Value, e.Row.Cells(Me.ColDiscString2.Index).Value, e.Row.Cells(Me.ColDiscAmount2.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc3.Index).Value, e.Row.Cells(Me.ColDiscString3.Index).Value, e.Row.Cells(Me.ColDiscAmount3.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryLocation.Index).Value, e.Row.Cells(Me.ColBatchNumber.Index).Value, e.Row.Cells(Me.ColExpiry.Index).Value, _
                        i.TxtReason.Text, e.Row.Cells(Me.ColPrimaryUnit.Index).Value, PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, _
                        e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me)
                    End If
                    '-----
                    ObjItem.HapusItemPerBaris(e.Row.Cells(Me.ColPrimaryAutoNumber.Index).Value)
                    '-----
                    ObjVoid.HapusItemViod(Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                Else
                    If ObjVoid.GetJml(PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value) > 0 Then
                        'Update data void
                        ObjVoid.UpdateItemViod(CDec(i.TxtQty.Text), _
                        e.Row.Cells(Me.ColUnitPrice.Index).Value, e.Row.Cells(Me.ColDiscString.Index).Value, e.Row.Cells(Me.ColDisc_Amount.Index).Value, _
                        e.Row.Cells(Me.ColNettPrice.Index).Value, e.Row.Cells(Me.ColSubTotal.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc1.Index).Value, e.Row.Cells(Me.ColDiscString1.Index).Value, e.Row.Cells(Me.ColDiscAmount1.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc2.Index).Value, e.Row.Cells(Me.ColDiscString2.Index).Value, e.Row.Cells(Me.ColDiscAmount2.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc3.Index).Value, e.Row.Cells(Me.ColDiscString3.Index).Value, e.Row.Cells(Me.ColDiscAmount3.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryLocation.Index).Value, e.Row.Cells(Me.ColBatchNumber.Index).Value, e.Row.Cells(Me.ColExpiry.Index).Value, _
                        i.TxtReason.Text, e.Row.Cells(Me.ColPrimaryUnit.Index).Value, PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, _
                        e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me)
                    Else
                        'Simpan data void
                        ObjVoid.SimpanItemViodSatuSatu(CDec(i.TxtQty.Text), _
                        e.Row.Cells(Me.ColUnitPrice.Index).Value, e.Row.Cells(Me.ColDiscString.Index).Value, e.Row.Cells(Me.ColDisc_Amount.Index).Value, _
                        e.Row.Cells(Me.ColNettPrice.Index).Value, e.Row.Cells(Me.ColSubTotal.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc1.Index).Value, e.Row.Cells(Me.ColDiscString1.Index).Value, e.Row.Cells(Me.ColDiscAmount1.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc2.Index).Value, e.Row.Cells(Me.ColDiscString2.Index).Value, e.Row.Cells(Me.ColDiscAmount2.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryDisc3.Index).Value, e.Row.Cells(Me.ColDiscString3.Index).Value, e.Row.Cells(Me.ColDiscAmount3.Index).Value, _
                        e.Row.Cells(Me.ColPrimaryLocation.Index).Value, e.Row.Cells(Me.ColBatchNumber.Index).Value, e.Row.Cells(Me.ColExpiry.Index).Value, _
                        i.TxtReason.Text, e.Row.Cells(Me.ColPrimaryUnit.Index).Value, PrimaryUserDelete, Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, _
                        e.Row.Cells(Me.ColPrimaryCodeNo.Index).Value, Me)
                    End If
                    'Update data 
                    Me.ObjItem.UpdateDataButtonInputPlusDiskon(e.Row.Cells(Me.ColPrimaryAutoNumber.Index).Value, e.Row.Cells(Me.ColQty.Index).Value, _
                    e.Row.Cells(Me.ColDiscString.Index).Value, e.Row.Cells(Me.ColDisc_Amount.Index).Value, e.Row.Cells(Me.ColNettPrice.Index).Value, _
                    e.Row.Cells(Me.ColSubTotal.Index).Value, e.Row.Cells(Me.ColDiscAmount1.Index).Value, e.Row.Cells(Me.ColDiscAmount2.Index).Value, _
                    e.Row.Cells(Me.ColDiscAmount3.Index).Value, _
                    e.Row.Cells(Me.ColPrimaryDisc1.Index).Value, e.Row.Cells(Me.ColPrimaryDisc2.Index).Value, e.Row.Cells(Me.ColPrimaryDisc3.Index).Value, _
                    e.Row.Cells(Me.ColDiscString1.Index).Value, e.Row.Cells(Me.ColDiscString2.Index).Value, e.Row.Cells(Me.ColDiscString3.Index).Value)
                    For Each Row In Me.dgvItem.Rows
                        Me.SalinKeObjekUntukUpdateDataItem(Row)
                    Next
                    Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                    Me.Obj.UpdateDataHeader(Me.UserCode)
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            'simpan datanya pada item void yang merujuk pada session yang bersangkutan

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnVoid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoid.Click
        Dim ij As New FrmTrans_POS_ConfirmPassword
        Dim iy As New FrmTrans_POS_VoidQty1
        Dim Hasil As MsgBoxResult
        Dim vUserCode As String = ""
        Dim vPrimaryUser As Double = -1

        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Obj.CekApakahSudahPaid(Me.TxtPrimary.Text) = True Then
                MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Obj.GetIsVoid(Me.TxtPrimary.Text) = 1 Then
                MsgBox("This transaction has been void", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> "-1" Then
                Hasil = MsgBox("Do you want to void transaction POS ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    'yang bisa dihapuskan ada transaksi dari session yang masih open
                    If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Open Then
                        'jika perlu konfirmasi, maka kelurkan form konfirmasi terlebih dahulu
                        If ObjGeneralSet.GetAllowUserToCancelWholeTransaction > 0 Then
                            If ObjGeneralSet.GetUseAuthorizationToCancelWholeTransaction > 0 Then
                                Me.mFormPassword = ij
                                Me.mFormPassword.cboUser.Focus()
                                Me.mFormPassword.ShowDialog()
                                If ij.mCancel = True Then
                                    Exit Sub
                                End If
                                vUserCode = ij.cboUser.Text
                                vPrimaryUser = ObjUser.GetPrimary(vUserCode)
                                ij.Dispose()
                            End If
                        Else
                            MsgBox("Transaction POS can't delete", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        iy.ShowDialog()
                        If iy.mrOk = True Then
                            'jika tidak perlu konfirmasi, maka langsung set properties "is void" dalam database menjadi 1
                            Me.Obj.UpdateIsVoid(Me.TxtPrimary.Text, 1)
                            Me.ObjVoid.SimpanItem(Me.TxtPrimary.Text, Me.TxtPrimarySession.Text, Me.dgvItem, Me, iy.TxtReason.Text, vPrimaryUser)
                            Me.btnNew.PerformClick()
                        Else
                            Exit Sub
                        End If
                    Else
                        MsgBox("Transaction POS can't delete because transaction session is closed", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            iy.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            'ketika dgv item ini tidak ada isinya, maka button ini tidak berlaku
            Me.Baru = True
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dgvItem.Rows.Count > 0 Then
                'jika dgv item ada isi, maka tampilkan transaksi baru
                Me.TxtDisplay.Clear()
                Me.TxtPrimary.Text = "-1"
                Me.TxtPrimaryMember.Text = "-1"
                Me.TxtMember.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNo.Clear()
                Me.TxtQty.Text = "1"
                Me.dtpTanggalTransaksi.Value = Now
                Me.TxtIdNumber.Text = Obj.GetNoBaru(Me.dtpTanggalTransaksi.Value, Me.TxtPrimarySession.Text)
                Me.dgvItem.Rows.Clear()
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
            End If
            Me.AktifkanDetail()
            Me.PbxPaid.Visible = False
            Me.PbxViod.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FungsiNextUpdateCekDataNya(ByVal TBLNext As DseTrans_POS.TBLNextDataTable)
        Dim Row As DseTrans_POS.TBLNextRow
        Dim Ketemu As Boolean = False
        Dim Index As Integer = -1
        Try
            For i As Integer = 0 To TBLNext.Rows.Count - 1
                Row = TBLNext.Rows(i)
                If Row.CheckPosisi = 1 Then
                    Index = i
                    Ketemu = True
                    Exit For
                End If
            Next
            If Ketemu = True Then
                If Index = TBLNext.Rows.Count - 1 Then
                    Row = TBLNext.Rows(0)
                    Row.CheckPosisi = 1
                    Row = TBLNext.Rows(Index)
                    Row.CheckPosisi = 0
                Else
                    Row = TBLNext.Rows(Index + 1)
                    Row.CheckPosisi = 1
                    Row = TBLNext.Rows(Index)
                    Row.CheckPosisi = 0
                End If
            Else
                Row = TBLNext.Rows(0)
                Row.CheckPosisi = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function FungsiNextBerjalan() As Boolean
        Dim Row As DseTrans_POS.TBLNextRow
        Try
            'Cek ada data yang belum dibayar
            If Obj.GetFungsiListButtonNext(Me.TxtPrimarySession.Text, Me.TBLNext) = True Then
                Me.FungsiNextUpdateCekDataNya(Me.TBLNext)
                'Bersihkan Field
                Me.TxtPrimary.Text = "-1"
                Me.dtpTanggalTransaksi.Value = Now
                Me.TxtPrimaryMember.Text = "-1"
                Me.TxtMember.Clear()
                Me.KosongkanFieldItem()
                '------
                Obj.ResetObject(Me.dgvItem)
                For Each Row In Me.TBLNext.Rows
                    If Row.CheckPosisi = 1 Then
                        Obj.CariDetail(Row.PrimaryPOS, Me.dgvItem, Me.ObjCD, Me)
                        Exit For
                    Else
                        If TBLNext.Rows.Count = 1 Then
                            Obj.CariDetail(Row.PrimaryPOS, Me.dgvItem, Me.ObjCD, Me)
                            Exit For
                        End If
                    End If
                Next
                Me.SalinKeField()
                Me.TxtPrimaryMember.Text = Obj.PrimaryMember
                Me.TxtMember.Text = ObjMember.GetCode(Me.TxtPrimaryMember.Text)
                Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)

                Me.AktifkanDetail()
                Me.TxtDisplay.BringToFront()
                Me.Baru = False
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'jika tombol ini dipencet, maka tampilkan transaksi selanjutnya, pada session yang sama yang belum memiliki item pembayaran
            If Me.dgvItem.Rows.Count > 0 Then
                If Me.TxtPrimary.Text <> "-1" Then
                    'Btn next dijalankan
                    Me.FungsiNextBerjalan()
                End If
            Else
                If Me.TxtPrimary.Text <> "-1" Then
                    'Check No Transkasi terakhir ada item nya (kalo ga ada hapus no itu)
                    If ObjItem.GetJmlByPrimaryTransaksi(Me.TxtPrimary.Text) <= 0 Then
                        'Ambil Primary Max, lalu dihapus 
                        Obj.Hapus(Me.TxtPrimary.Text)
                    End If
                    If Me.FungsiNextBerjalan() = False Then
                        Me.TxtPrimary.Text = "-1"
                        Me.TxtPrimaryMember.Text = "-1"
                        Me.TxtMember.Clear()
                        Me.dgvItem.Rows.Clear()
                        Me.dtpTanggalTransaksi.Value = Now
                        Me.KosongkanFieldItem()
                        'Get No Baru
                        Me.TxtIdNumber.Text = Obj.GetNoBaru(Me.dtpTanggalTransaksi.Value, Me.TxtPrimarySession.Text)
                        Me.TxtQty.Focus()
                    End If
                Else
                    'Btn next dijalankan
                    Me.FungsiNextBerjalan()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDisplay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDisplay.KeyDown
        Try
            e.SuppressKeyPress = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateHitunganItemUntukDiskon()
        Dim Row As DataGridViewRow
        Dim RowDiskon As DseMast_Diskon.tblmaster_discount2Row
        Try
            For Each Row In Me.dgvItem.Rows
                With Row
                    If Me.TxtPrimaryCodeNo.Text <> -1 Then
                        .Cells(Me.ColDiscString.Index).Value = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                         Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, .Cells(Me.ColQty.Index).Value + Me.TxtQty.Text, Me.TBLDiskon, 0)
                    Else
                        .Cells(Me.ColDiscString.Index).Value = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                         Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, .Cells(Me.ColQty.Index).Value, Me.TBLDiskon, 0)
                    End If
                    .Cells(Me.ColDisc_Amount.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString.Index).Value, _
                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                    .Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColUnitPrice.Index).Value - (CDec(.Cells(Me.ColDisc_Amount.Index).Value) / CDec(.Cells(Me.ColQty.Index).Value)), Me.DefPrec)
                    .Cells(Me.ColSubTotal.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColNettPrice.Index).Value, Me.DefPrec)
                    For IL As Integer = 0 To TBLDiskon.Rows.Count - 1
                        RowDiskon = TBLDiskon.Rows(IL)
                        If IL = 0 Then
                            .Cells(Me.ColPrimaryDisc1.Index).Value = RowDiskon.Primary
                            .Cells(Me.ColDiscString1.Index).Value = RowDiskon.Discount_Rate
                            .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                            CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        ElseIf IL = 1 Then
                            .Cells(Me.ColPrimaryDisc2.Index).Value = RowDiskon.Primary
                            .Cells(Me.ColDiscString2.Index).Value = RowDiskon.Discount_Rate
                            .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                            CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        ElseIf IL = 2 Then
                            .Cells(Me.ColPrimaryDisc3.Index).Value = RowDiskon.Primary
                            .Cells(Me.ColDiscString3.Index).Value = RowDiskon.Discount_Rate
                            .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                            CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                        End If
                    Next
                    'Hitung lagi Total IDR yang di display
                    Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
                    'Lakukan kembali perhitungan salin ke objek
                    Me.SalinKeObjekUntukUpdateDataItem(Row)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
     Handles TxtMember.KeyDown
        Dim objCode As New ClsMast_Member
        Dim i As New FrmMast_Member_Cari
        Dim RowDGV as DataGridViewrow
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Me.TxtMember.Focus()
                Exit Sub
            End If

            If e.KeyValue = 13 Then
                If Me.TxtMember.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtMember.Text) = True Then
                        Me.MunculkanMsgBoxMember = True
                        Me.TxtPrimaryMember.Text = objCode.GetPrimary(Me.TxtMember.Text)
                        Me.TxtMember.Text = objCode.GetCode(Me.TxtPrimaryMember.Text)
                        If ObjMember.GetExpiryDate(Me.TxtPrimaryMember.Text) < Me.dtpTanggalTransaksi.Value.Date Then
                            MsgBox("Please renewal member", MsgBoxStyle.Information, "Information")
                            Me.TxtMember.Clear()
                            Me.TxtPrimaryMember.Text = "-1"
                            Me.TxtMember.Focus()
                            Exit Sub
                        End If
                        If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                            MsgBox("Please enter positif qty", MsgBoxStyle.Information, "Information")
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        Me.UpdateHitunganItemUntukDiskon()
                        Me.Obj.UpdatePrimaryMember(Me.TxtPrimary.Text, Me.TxtPrimaryMember.Text)
                        If Me.mrCloseFind = False Then
                            If Me.Baru = False Then
                                If Me.TxtQty.Enabled = True Then
                                    Me.Obj.UpdateDataHeader(Me.UserCode)
                                    'Update Item
                                    For Each RowDGV In Me.dgvItem.Rows
                                        Me.ObjItem.UpdateDataButtonInputPlusDiskon(RowDGV.Cells(Me.ColPrimaryAutoNumber.Index).Value, RowDGV.Cells(Me.ColQty.Index).Value, _
                                        RowDGV.Cells(Me.ColDiscString.Index).Value, RowDGV.Cells(Me.ColDisc_Amount.Index).Value, RowDGV.Cells(Me.ColNettPrice.Index).Value, _
                                        RowDGV.Cells(Me.ColSubTotal.Index).Value, RowDGV.Cells(Me.ColDiscAmount1.Index).Value, RowDGV.Cells(Me.ColDiscAmount2.Index).Value, _
                                        RowDGV.Cells(Me.ColDiscAmount3.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc1.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc2.Index).Value, _
                                        RowDGV.Cells(Me.ColPrimaryDisc3.Index).Value, RowDGV.Cells(Me.ColDiscString1.Index).Value, RowDGV.Cells(Me.ColDiscString2.Index).Value, _
                                        RowDGV.Cells(Me.ColDiscString3.Index).Value)
                                    Next
                                End If
                            End If
                        End If
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.btnFindMember.PerformClick()
                    End If
                Else
                    Me.btnFindMember.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMember_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtMember.Leave
        Dim objCode As New ClsMast_Member
        Dim i As New FrmMast_Member_Cari
        i.BackColorAll = Me.BackColorAll
        Dim RowDGV As DataGridViewRow
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                Me.TxtMember.Clear()
                Exit Sub
            End If

            If Me.TxtMember.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtMember.Text) = True Then
                    If ObjMember.GetExpiryDate(objCode.GetPrimary(Me.TxtMember.Text)) < Me.dtpTanggalTransaksi.Value.Date Then
                        If Me.MunculkanMsgBoxMember = False Then
                            If Me.TxtQty.Enabled = True Then
                                MsgBox("Please renewal member", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                        Me.TxtMember.Clear()
                        Exit Sub
                    End If
                    Me.TxtPrimaryMember.Text = objCode.GetPrimary(Me.TxtMember.Text)
                    Me.TxtMember.Text = objCode.GetCode(Me.TxtPrimaryMember.Text)
                Else
                    If Me.TxtMember.Text <> "" Then
                        i.txtCari.Text = Me.TxtMember.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        If ObjMember.GetExpiryDate(i.mPrimary) < Me.dtpTanggalTransaksi.Value.Date Then
                            Me.TxtMember.Clear()
                            Exit Sub
                        End If
                        Me.TxtPrimaryMember.Text = i.mPrimary
                        Me.TxtMember.Text = objCode.GetCode(Me.TxtPrimaryMember.Text)
                    Else
                        Me.TxtMember.Clear()
                        Me.TxtPrimaryMember.Text = "-1"
                        Me.TxtMember.Focus()
                    End If
                End If
            Else
                Me.TxtMember.Clear()
                Me.TxtPrimaryMember.Text = "-1"
            End If
            If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
                Exit Sub
            End If
            Me.UpdateHitunganItemUntukDiskon()
            Me.Obj.UpdatePrimaryMember(Me.TxtPrimary.Text, Me.TxtPrimaryMember.Text)
            If Me.mrCloseFind = False Then
                If Me.Baru = False Then
                    If Me.TxtQty.Enabled = True Then
                        Me.Obj.UpdateDataHeader(Me.UserCode)
                        'Update Item
                        For Each RowDGV In Me.dgvItem.Rows
                            Me.ObjItem.UpdateDataButtonInputPlusDiskon(RowDGV.Cells(Me.ColPrimaryAutoNumber.Index).Value, RowDGV.Cells(Me.ColQty.Index).Value, _
                            RowDGV.Cells(Me.ColDiscString.Index).Value, RowDGV.Cells(Me.ColDisc_Amount.Index).Value, RowDGV.Cells(Me.ColNettPrice.Index).Value, _
                            RowDGV.Cells(Me.ColSubTotal.Index).Value, RowDGV.Cells(Me.ColDiscAmount1.Index).Value, RowDGV.Cells(Me.ColDiscAmount2.Index).Value, _
                            RowDGV.Cells(Me.ColDiscAmount3.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc1.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc2.Index).Value, _
                            RowDGV.Cells(Me.ColPrimaryDisc3.Index).Value, RowDGV.Cells(Me.ColDiscString1.Index).Value, RowDGV.Cells(Me.ColDiscString2.Index).Value, _
                            RowDGV.Cells(Me.ColDiscString3.Index).Value)
                        Next
                    End If
                End If
            End If
            Me.MunculkanMsgBoxMember = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindMember.Click
        Dim i As New FrmMast_Member_Cari
        Dim ObjCode As New ClsMast_Member
        i.BackColorAll = Me.BackColorAll
        Dim RowDGV As DataGridViewRow
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Me.TxtMember.Focus()
                Exit Sub
            End If

            If Me.TxtMember.Text <> "" Then
                i.txtCari.Text = Me.TxtMember.Text
            End If

            i.ShowDialog()
            If i.mPrimary > 0 Then
                Me.MunculkanMsgBoxMember = True
                Me.TxtPrimaryMember.Text = i.mPrimary
                Me.TxtMember.Text = ObjCode.GetCode(Me.TxtPrimaryMember.Text)
                If ObjMember.GetExpiryDate(Me.TxtPrimaryMember.Text) < Me.dtpTanggalTransaksi.Value.Date Then
                    MsgBox("Please renewal member", MsgBoxStyle.Information, "Information")
                    Me.TxtMember.Clear()
                    Me.TxtPrimaryMember.Text = "-1"
                    Me.TxtMember.Focus()
                    Exit Sub
                End If
                Me.TxtCodeNo.Focus()
            Else
                Me.TxtMember.Clear()
                Me.TxtPrimaryMember.Text = "-1"
                Me.TxtMember.Focus()
            End If
            If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                MsgBox("Please enter positif qty", MsgBoxStyle.Information, "Information")
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
                Exit Sub
            End If
            Me.UpdateHitunganItemUntukDiskon()
            Me.Obj.UpdatePrimaryMember(Me.TxtPrimary.Text, Me.TxtPrimaryMember.Text)
            If Me.mrCloseFind = False Then
                If Me.Baru = False Then
                    If Me.TxtQty.Enabled = True Then
                        Me.Obj.UpdateDataHeader(Me.UserCode)
                        'Update Item
                        For Each RowDGV In Me.dgvItem.Rows
                            Me.ObjItem.UpdateDataButtonInputPlusDiskon(RowDGV.Cells(Me.ColPrimaryAutoNumber.Index).Value, RowDGV.Cells(Me.ColQty.Index).Value, _
                            RowDGV.Cells(Me.ColDiscString.Index).Value, RowDGV.Cells(Me.ColDisc_Amount.Index).Value, RowDGV.Cells(Me.ColNettPrice.Index).Value, _
                            RowDGV.Cells(Me.ColSubTotal.Index).Value, RowDGV.Cells(Me.ColDiscAmount1.Index).Value, RowDGV.Cells(Me.ColDiscAmount2.Index).Value, _
                            RowDGV.Cells(Me.ColDiscAmount3.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc1.Index).Value, RowDGV.Cells(Me.ColPrimaryDisc2.Index).Value, _
                            RowDGV.Cells(Me.ColPrimaryDisc3.Index).Value, RowDGV.Cells(Me.ColDiscString1.Index).Value, RowDGV.Cells(Me.ColDiscString2.Index).Value, _
                            RowDGV.Cells(Me.ColDiscString3.Index).Value)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'Tambahkan Fungsi Get Diskon
    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCodeNo.KeyDown
        Dim objCode As New clsMasterProduct
        Dim i As New frmCariMasterProduct
        i.UntukPOS = True
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Focus()
                Exit Sub
            End If

            If e.KeyValue = 13 Then
                If Me.TxtCodeNo.Text <> "" Then
                    If objCode.ApakahSudahAdaBarcode(Me.TxtCodeNo.Text) = True Then
                        Me.TxtPrimaryCodeNo.Text = objCode.GetPrimaryByBarcode(Me.TxtCodeNo.Text)
                        Me.TxtCodeNo.Text = objCode.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                        If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                            MsgBox("Please enter positif qty", MsgBoxStyle.Information, "Information")
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        If Me.TxtQty.Text.Length > 15 Then
                            MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        Me.InputDataItem()
                        Me.TxtQty.Focus()
                        Me.TxtQty.SelectAll()
                    Else
                        Me.btnFindCodeNo.PerformClick()
                        If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        If Me.TxtQty.Text.Length > 15 Then
                            MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        Me.InputDataItem()
                    End If
                Else
                    Me.btnFindCodeNo.PerformClick()
                    If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                        Me.TxtQty.Focus()
                        Me.TxtQty.SelectAll()
                        Exit Sub
                    End If
                    If Me.TxtQty.Text.Length > 15 Then
                        MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                        Me.TxtQty.Focus()
                        Me.TxtQty.SelectAll()
                        Exit Sub
                    End If
                    Me.InputDataItem()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCodeNo.Leave
        Dim objCode As New clsMasterProduct
        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        i.UntukPOS = True
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                Me.TxtCodeNo.Clear()
                Exit Sub
            End If

            If Me.TxtCodeNo.Text <> "" Then
                If objCode.ApakahSudahAdaBarcode(Me.TxtCodeNo.Text) = True Then
                    Me.TxtPrimaryCodeNo.Text = objCode.GetPrimaryByBarcode(Me.TxtCodeNo.Text)
                    Me.TxtCodeNo.Text = objCode.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                    If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                        Me.TxtQty.Focus()
                        Me.TxtQty.SelectAll()
                        Exit Sub
                    End If
                    If Me.TxtQty.Text.Length > 15 Then
                        Exit Sub
                    End If
                    Me.InputDataItem()
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = objCode.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                        If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                            Exit Sub
                        End If
                        If Me.TxtQty.Text.Length > 15 Then
                            Exit Sub
                        End If
                        Me.InputDataItem()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                If Me.dgvItem.Rows.Count <= 0 Then
                    Me.TxtQty.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        Dim ObjCode As New clsMasterProduct
        i.BackColorAll = Me.BackColorAll
        i.UntukPOS = True
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Focus()
                Exit Sub
            End If

            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = ObjCode.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
                    MsgBox("Please enter positif qty", MsgBoxStyle.Information, "Information")
                    Me.TxtQty.Focus()
                    Me.TxtQty.SelectAll()
                    Exit Sub
                End If
                If Me.TxtQty.Text.Length > 15 Then
                    MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                    Me.TxtQty.Focus()
                    Me.TxtQty.SelectAll()
                    Exit Sub
                End If
                Me.InputDataItem()
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'If ObjCD.CekBilanganPositif(Me.TxtQty.Text) = False Then
            'MsgBox("Please input bilangan positif", MsgBoxStyle.Information, "Information")
            'Me.TxtQty.Focus()
            'Return False
            'End If

            If ObjSession.GetJmlByPrimary(Me.TxtPrimarySession.Text) <= 0 Then
                'MsgBox("you cannot input transaction POS because transaction session has been deleted", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub KosongkanFieldItem()
        Try
            Me.TxtQty.Text = "1"
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.TxtCodeNo.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimaryCode As Double) As Double
        Dim RowDGV As DataGridViewRow
        Try
            For Each RowDGV In Me.dgvItem.Rows
                If RowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCode Then
                    Return RowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub InputDataItem()
        Dim index As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim RowDiskon As DseMast_Diskon.tblmaster_discount2Row
        Try
            If ApakahDataItemBenar() = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    '-----
                    'cek apakah sudah ada item jenis pekerjaan sama yang masuk
                    index = Me.GetIndexItem(Me.TxtPrimaryCodeNo.Text)
                    'Salin ke objek
                    Me.SalinKeObjekUntukFungsiInput(index)
                    If index < 0 Then
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.ColNo.Index).Value = ""
                            .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                            .Cells(Me.ColDesc.Index).Value = ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                            .Cells(Me.ColLocation.Index).Value = ""
                            .Cells(Me.ColBatchNumber.Index).Value = ""
                            .Cells(Me.ColExpiry.Index).Value = Format(CDate("1/1/1950"), "dd MMM yyyy")
                            If IsNumeric(Me.TxtQty.Text) = True Then
                                .Cells(Me.ColQty.Index).Value = Me.TxtQty.Text
                            Else
                                .Cells(Me.ColQty.Index).Value = 1
                            End If
                            .Cells(Me.ColUnit.Index).Value = ObjProd.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                            .Cells(Me.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjProd.GetUnitPrice(ObjGeneralSet.GetPrimaryCurr(), Me.TxtPrimaryCodeNo.Text), Me.DefPrec, Me.RoundingTipe)
                            .Cells(Me.ColDiscString.Index).Value = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), Me.TxtPrimaryCodeNo.Text, _
                             Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, Me.TxtQty.Text, Me.TBLDiskon, 0)
                            .Cells(Me.ColDisc_Amount.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString.Index).Value, _
                             CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            .Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColUnitPrice.Index).Value - CDec(.Cells(Me.ColDisc_Amount.Index).Value) / CDec(.Cells(Me.ColQty.Index).Value), Me.DefPrec)
                            .Cells(Me.ColSubTotal.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColNettPrice.Index).Value, Me.DefPrec)
                            .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                            .Cells(Me.ColPrimaryUnit.Index).Value = ObjUnit.GetPrimary(.Cells(Me.ColUnit.Index).Value)
                            .Cells(Me.ColPrimaryLocation.Index).Value = "-1"
                            .Cells(Me.ColPrimaryDisc1.Index).Value = -1
                            .Cells(Me.ColDiscString1.Index).Value = "0"
                            .Cells(Me.ColDiscAmount1.Index).Value = 0
                            .Cells(Me.ColPrimaryDisc2.Index).Value = -1
                            .Cells(Me.ColDiscString2.Index).Value = "0"
                            .Cells(Me.ColDiscAmount2.Index).Value = 0
                            .Cells(Me.ColPrimaryDisc3.Index).Value = -1
                            .Cells(Me.ColDiscString3.Index).Value = "0"
                            .Cells(Me.ColDiscAmount3.Index).Value = 0
                            For IL As Integer = 0 To TBLDiskon.Rows.Count - 1
                                RowDiskon = TBLDiskon.Rows(IL)
                                If IL = 0 Then
                                    .Cells(Me.ColPrimaryDisc1.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString1.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                ElseIf IL = 1 Then
                                    .Cells(Me.ColPrimaryDisc2.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString2.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                ElseIf IL = 2 Then
                                    .Cells(Me.ColPrimaryDisc3.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString3.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                End If
                                .Cells(Me.ColPrimaryAutoNumber.Index).Value = "-1"
                            Next
                        End With
                        'set number
                        Me.ObjCD.SetNumber(Me.dgvItem)
                        'Fungsi simpan berlaku
                        If Me.dgvItem.Rows.Count > 1 Then
                            ObjItem.SimpanItemSatuSatu(Me.TxtPrimary.Text, Me.dgvItem, Me.dgvItem.Rows.Count - 1, Me)
                            Me.dgvItem.Item(Me.ColPrimaryAutoNumber.Index, Me.dgvItem.Rows.Count - 1).Value = ObjItem.PrimaryAutoNumber
                            Me.Obj.UpdateDataHeader(Me.UserCode)
                        Else
                            If Me.TxtPrimary.Text <> "-1" Then
                                ObjItem.SimpanItemSatuSatu(Me.TxtPrimary.Text, Me.dgvItem, Me.dgvItem.Rows.Count - 1, Me)
                                Me.dgvItem.Item(Me.ColPrimaryAutoNumber.Index, Me.dgvItem.Rows.Count - 1).Value = ObjItem.PrimaryAutoNumber
                                Me.Obj.UpdateDataHeader(Me.UserCode)
                            Else
                                Obj.SimpanBaru(Me.dtpTanggalTransaksi.Value, Me.TxtPrimarySession.Text, Me.dgvItem, Me.UserCode, Me)
                                Me.dgvItem.Item(Me.ColPrimaryAutoNumber.Index, 0).Value = ObjItem.GetPrimaryAutoNumber(Obj.Primary)
                                Me.SalinKeField()
                            End If
                        End If
                        index = Me.dgvItem.Rows.Count - 1
                    Else
                        'Tambahkan Qty dan Hitung kembali
                        rowDGV = Me.dgvItem.Rows(index)
                        With rowDGV
                            If IsNumeric(Me.TxtQty.Text) = True Then
                                .Cells(Me.ColQty.Index).Value += CDec(Me.TxtQty.Text)
                            Else
                                .Cells(Me.ColQty.Index).Value += 1
                            End If
                            .Cells(Me.ColDiscString.Index).Value = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), Me.TxtPrimaryCodeNo.Text, _
                                                                Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, .Cells(Me.ColQty.Index).Value, Me.TBLDiskon, 0)
                            .Cells(Me.ColDisc_Amount.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString.Index).Value, _
                                                                 CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), _
                                                                 DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            .Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColUnitPrice.Index).Value - (CDec(.Cells(Me.ColDisc_Amount.Index).Value) / CDec(.Cells(Me.ColQty.Index).Value)), Me.DefPrec)
                            .Cells(Me.ColSubTotal.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColNettPrice.Index).Value, Me.DefPrec)
                            .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                                                                  CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                                                                  CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                                                                  CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                            For IL As Integer = 0 To TBLDiskon.Rows.Count - 1
                                RowDiskon = TBLDiskon.Rows(IL)
                                If IL = 0 Then
                                    .Cells(Me.ColPrimaryDisc1.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString1.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount1.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString1.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                ElseIf IL = 1 Then
                                    .Cells(Me.ColPrimaryDisc2.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString2.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount2.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString2.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                ElseIf IL = 2 Then
                                    .Cells(Me.ColPrimaryDisc3.Index).Value = RowDiskon.Primary
                                    .Cells(Me.ColDiscString3.Index).Value = RowDiskon.Discount_Rate
                                    .Cells(Me.ColDiscAmount3.Index).Value = ObjCD.Get_Disc_Amount(.Cells(Me.ColDiscString3.Index).Value, _
                                    CDec(.Cells(Me.ColQty.Index).Value * .Cells(Me.ColUnitPrice.Index).Value), DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                End If
                            Next
                        End With
                        'update data
                        Me.ObjItem.UpdateDataButtonInputPlusDiskon(rowDGV.Cells(Me.ColPrimaryAutoNumber.Index).Value, rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDiscString.Index).Value, _
                        rowDGV.Cells(Me.ColDisc_Amount.Index).Value, rowDGV.Cells(Me.ColNettPrice.Index).Value, rowDGV.Cells(Me.ColSubTotal.Index).Value, rowDGV.Cells(Me.ColDiscAmount1.Index).Value, _
                        rowDGV.Cells(Me.ColDiscAmount2.Index).Value, rowDGV.Cells(Me.ColDiscAmount3.Index).Value, _
                        rowDGV.Cells(Me.ColPrimaryDisc1.Index).Value, rowDGV.Cells(Me.ColPrimaryDisc2.Index).Value, rowDGV.Cells(Me.ColPrimaryDisc3.Index).Value, _
                        rowDGV.Cells(Me.ColDiscString1.Index).Value, rowDGV.Cells(Me.ColDiscString2.Index).Value, rowDGV.Cells(Me.ColDiscString3.Index).Value)
                        'Update data header nya
                        Me.Obj.UpdateDataHeader(Me.UserCode)
                    End If
                    Me.KosongkanFieldItem()
                    Me.TxtQty.Focus()

                    'isi display (Jika ada disc Format Font-nya beda)
                    If Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString = "0" Then
                        Me.TxtDisplay.Font = Me.FontItem_2Baris
                        Me.TxtDisplay.Text = Me.dgvItem.Item(Me.ColCodeNo.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " - " + _
                        Me.dgvItem.Item(Me.ColDesc.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " @" + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + vbCrLf + _
                        Me.dgvItem.Item(Me.ColQty.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.Rows.Count - 1).Value + " x " + _
                        ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + _
                        " = IDR " + Me.dgvItem.Item(Me.ColSubTotal.Index, Me.dgvItem.Rows.Count - 1).Value.ToString
                        'Me.TxtDisplay.Text = "KY-011 - Tempat Dupa Kecil @IDR 50.000" + vbCrLf + "Discount Member 1% IDR 500" + vbCrLf + _
                        '"3 pcs x IDR 49.500 = IDR 148.500"
                        Me.Timer1.Enabled = False
                        Me.Timer1.Enabled = True

                    Else
                        Me.TxtDisplay.Font = Me.FontItem_3Baris
                        Me.TxtDisplay.Text = Me.dgvItem.Item(Me.ColCodeNo.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " - " + _
                        Me.dgvItem.Item(Me.ColDesc.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " @" + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + vbCrLf + _
                        "Discount " + Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString, Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD) + vbCrLf + _
                        Me.dgvItem.Item(Me.ColQty.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.Rows.Count - 1).Value + " x " + _
                        ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString - CDec(ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value, Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)), Me.DefPrec, Me.RoundingTipe).ToString + _
                        " = IDR " + Me.dgvItem.Item(Me.ColSubTotal.Index, Me.dgvItem.Rows.Count - 1).Value.ToString
                        'Me.TxtDisplay.Text = "KY-011 - Tempat Dupa Kecil @IDR 50.000" + vbCrLf + "Discount Member 1% IDR 500" + vbCrLf + _
                        '"3 pcs x IDR 49.500 = IDR 148.500"
                        Me.Timer1.Enabled = False
                        Me.Timer1.Enabled = True
                    End If

                    '-----Scrolling DGV-----
                    If Me.dgvItem.FirstDisplayedScrollingRowIndex < index Then
                        While Me.dgvItem.Rows(index).Displayed = False _
                        Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                            Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                        End While
                    Else
                        While Me.dgvItem.Rows(index).Displayed = False
                            Me.dgvItem.FirstDisplayedScrollingRowIndex = _
                            Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                        End While
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim ij As New FrmTrans_POS_Payment
        Dim ik As New FrmTrans_POS_PayoutDropCash_Input
        Try
            If e.KeyCode = Keys.F1 Then
                'cek dulu apakah ada yang belum di input  
                If Me.dgvItem.Rows.Count > 0 Then
                    If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                        MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                        , MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    'Buka Form Payment
                    If Me.dgvItem.Rows.Count > 0 Then
                        If Obj.CekApakahSudahPaid(Me.TxtPrimary.Text) = False Then
                            If Obj.GetIsVoid(Me.TxtPrimary.Text) = 0 Then
                                'Single Payment
                                ij.ParentForm_1 = Me
                                ij.Obj = Me.Obj
                                ij.cboPaymentTipe.Focus()
                                ij.ShowDialog()
                            Else
                                MsgBox("This transaction has been void", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        Else
                            MsgBox("This transaction has been paid", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If e.KeyCode = Keys.F2 Then
                If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                    MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                    , MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                ik.Parent_Form = Me
                ik.cboPayout_DropCash.Focus()
                ik.ShowDialog()
            End If
            If e.KeyCode = Keys.F4 Then
                Me.btnNew.PerformClick()
            End If
            If e.KeyCode = Keys.F5 Then
                Me.btnNext.PerformClick()
            End If
            If e.KeyCode = Keys.F6 Then
                Me.btnVoid.PerformClick()
            End If
            If e.KeyCode = Keys.F7 Then
                Me.btnReprint.PerformClick()
            End If

            '----
            If e.KeyCode = Keys.F8 Then
                Me.Close()
            End If
            If e.KeyCode = Keys.F3 Then
                Me.btnFind.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyDown
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Me.TxtQty.Text = "1"
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
                Exit Sub
            End If

            If e.KeyValue = 13 Then
                'cek dulu kode product
                If Me.TxtQty.Text = "" Then
                    Me.TxtQty.Text = 1
                    Me.TxtCodeNo.Focus()
                    Me.TxtQty.SelectAll()
                    Exit Sub
                End If
                If ObjProd.ApakahSudahAda(Me.TxtQty.Text) = True Then
                    Me.TxtPrimaryCodeNo.Text = ObjProd.GetPrimary(Me.TxtQty.Text)
                    Me.TxtCodeNo.Text = ObjProd.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtQty.Text = 1
                    If Me.TxtQty.Text.Length > 15 Then
                        MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                        Me.TxtQty.Focus()
                        Me.TxtQty.SelectAll()
                        Exit Sub
                    End If
                    'Turun ke bawah
                    Me.InputDataItem()
                    Me.TxtQty.Focus()
                    Me.TxtQty.SelectAll()
                Else
                    'pake fungsi apakah angka / numeric
                    If ObjCD.CekNumerik(Me.TxtQty.Text) = False Then
                        Me.TxtQty.Text = 1
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtQty.Text = 1
                        Me.TxtCodeNo.Clear()
                        Me.TxtCodeNo.Focus()
                        Me.TxtQty.SelectAll()
                    Else
                        If Me.dgvItem.Rows.Count > 0 Then
                            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                                If Me.TxtQty.Text.Length > 15 Then
                                    MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                                    Me.TxtQty.Focus()
                                    Me.TxtQty.SelectAll()
                                    Exit Sub
                                End If
                                Me.InputDataItem()
                                Me.TxtQty.Focus()
                                Me.TxtQty.SelectAll()
                            Else
                                'cek data yang paling bawah kemudian hitung dan tambahkan Qty
                                'Dim IndexTerakhir As Integer = Me.dgvItem.Rows.Count - 1
                                'Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtQty.Text) + CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value), Me.DefPrec, Me.RoundingTipe)
                                'Me.dgvItem.Item(Me.ColDiscString.Index, IndexTerakhir).Value = Me.ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, IndexTerakhir).Value, _
                                'Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value), Me.TBLDiskon, 0)
                                'Me.dgvItem.Item(Me.ColDisc_Amount.Index, IndexTerakhir).Value = Me.ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString.Index, IndexTerakhir).Value, _
                                'CDec(CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value) * CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, IndexTerakhir).Value)), _
                                'Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                'Me.dgvItem.Item(Me.ColNettPrice.Index, IndexTerakhir).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColUnitPrice.Index, IndexTerakhir).Value - Me.dgvItem.Item(Me.ColDisc_Amount.Index, IndexTerakhir).Value, Me.DefPrec)
                                'Me.dgvItem.Item(Me.ColSubTotal.Index, IndexTerakhir).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value) * Me.dgvItem.Item(Me.ColNettPrice.Index, IndexTerakhir).Value, Me.DefPrec)
                                'Me.dgvItem.Item(Me.ColDiscAmount1.Index, IndexTerakhir).Value = Me.ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString1.Index, IndexTerakhir).Value, _
                                'CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value * Me.dgvItem.Item(Me.ColUnitPrice.Index, IndexTerakhir).Value), Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                'Me.dgvItem.Item(Me.ColDiscAmount2.Index, IndexTerakhir).Value = Me.ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString2.Index, IndexTerakhir).Value, _
                                'CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value * Me.dgvItem.Item(Me.ColUnitPrice.Index, IndexTerakhir).Value), Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                'Me.dgvItem.Item(Me.ColDiscAmount3.Index, IndexTerakhir).Value = Me.ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString3.Index, IndexTerakhir).Value, _
                                'CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value * Me.dgvItem.Item(Me.ColUnitPrice.Index, IndexTerakhir).Value), Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                                'Salin ke Objek
                                'Me.SalinKeObjekUntukFungsiInputKeyDownTxtQty(IndexTerakhir)
                                'update data
                                'Me.ObjItem.UpdateDataButtonInput(Me.dgvItem.Item(Me.ColPrimaryAutoNumber.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColQty.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColDiscString.Index, IndexTerakhir).Value, _
                                'Me.dgvItem.Item(Me.ColDisc_Amount.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColNettPrice.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColSubTotal.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColDiscAmount1.Index, IndexTerakhir).Value, _
                                'Me.dgvItem.Item(Me.ColDiscAmount2.Index, IndexTerakhir).Value, Me.dgvItem.Item(Me.ColDiscAmount3.Index, IndexTerakhir).Value)
                                'Update data header nya
                                'Me.Obj.UpdateDataHeader(Me.UserCode)
                                'display set ulang 
                                'If Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString = "0" Then
                                'Me.TxtDisplay.Font = Me.FontItem_2Baris
                                'Me.TxtDisplay.Text = Me.dgvItem.Item(Me.ColCodeNo.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " - " + _
                                'Me.dgvItem.Item(Me.ColDesc.Index, Me.dgvItem.Rows.Count - 1).Value.ToString(+" @" + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + vbCrLf + _
                                'Me.dgvItem.Item(Me.ColQty.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.Rows.Count - 1).Value + " x " + _
                                'ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + _
                                '" = IDR " + Me.dgvItem.Item(Me.ColSubTotal.Index, Me.dgvItem.Rows.Count - 1).Value.ToString)
                                'Me.Timer1.Enabled = False
                                'Me.Timer1.Enabled = True

                                'Else
                                'Me.TxtDisplay.Font = Me.FontItem_3Baris
                                'Me.TxtDisplay.Text = Me.dgvItem.Item(Me.ColCodeNo.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " - " + _
                                'Me.dgvItem.Item(Me.ColDesc.Index, Me.dgvItem.Rows.Count - 1).Value.ToString(+" @" + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + vbCrLf + _
                                '"Discount " + Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value.ToString, Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD) + vbCrLf + _
                                'Me.dgvItem.Item(Me.ColQty.Index, Me.dgvItem.Rows.Count - 1).Value.ToString + " " + Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.Rows.Count - 1).Value + " x " + _
                                'ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + " " + ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value.ToString - CDec(ObjCD.Get_Disc_Amount(Me.dgvItem.Item(Me.ColDiscString.Index, Me.dgvItem.Rows.Count - 1).Value, Me.dgvItem.Item(Me.ColUnitPrice.Index, Me.dgvItem.Rows.Count - 1).Value, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)), Me.DefPrec, Me.RoundingTipe).ToString + _
                                '" = IDR " + Me.dgvItem.Item(Me.ColSubTotal.Index, Me.dgvItem.Rows.Count - 1).Value.ToString)
                                'Me.Timer1.Enabled = False
                                'Me.Timer1.Enabled = True
                                'End If
                                'Focus
                                'Me.TxtQty.Text = "1"
                                'Me.TxtQty.Focus()
                            End If
                        Else
                            If Me.TxtQty.Text.Length > 15 Then
                                MsgBox("Maximum qty only 15 digits", MsgBoxStyle.Information, "Information")
                                Me.TxtQty.Focus()
                                Me.TxtQty.SelectAll()
                                Exit Sub
                            End If
                            Me.InputDataItem()
                            Me.TxtQty.Focus()
                            Me.TxtQty.SelectAll()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Leave
        Try
            If Me.TxtQty.Text = "" Then
                Me.TxtQty.Text = 1
                Exit Sub
            End If
            If ObjProd.ApakahSudahAda(Me.TxtQty.Text) = True Then
                Me.TxtPrimaryCodeNo.Text = ObjProd.GetPrimary(Me.TxtQty.Text)
                Me.TxtCodeNo.Text = ObjProd.GetBarcode(Me.TxtPrimaryCodeNo.Text)
                Me.TxtQty.Text = 1
                Me.TxtQty.Focus()
                Me.TxtQty.SelectAll()
            Else
                'pake fungsi apakah angka / numeric
                If ObjCD.CekNumerik(Me.TxtQty.Text) = False Then
                    Me.TxtQty.Text = 1
                    Me.TxtPrimaryCodeNo.Text = "-1"
                    Me.TxtQty.Text = 1
                    Me.TxtCodeNo.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyUp
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HitungSubTotalDGV() As Decimal
        Dim RowDGV As DataGridViewRow
        Dim retval As Decimal = 0
        Try
            For Each RowDGV In Me.dgvItem.Rows
                retval += RowDGV.Cells(Me.ColSubTotal.Index).Value
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function HitungSubTotalDGV(ByVal IndexDGV As Integer) As Decimal
        Dim RowDGV As DataGridViewRow
        Dim retval As Decimal = 0
        Try
            For Each RowDGV In Me.dgvItem.Rows
                If RowDGV.Index <> IndexDGV Then
                    retval += RowDGV.Cells(Me.ColSubTotal.Index).Value
                End If
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SalinKeObjekUntukFungsiInput(ByVal IndexDGV As Integer)
        Dim vUnitPrice As Decimal = 0
        Dim vSubTotalItem As Decimal = 0
        Dim vSubTotalDGV As Decimal = 0
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Dim vDiscStringItemInputDGV As String = 0
        Dim vDiscAmountItemInputDGV As Decimal = 0
        Try
            'Single Payment
            With Me.Obj
                If IsNumeric(Me.TxtQty.Text) = False Then
                    Me.TxtQty.Text = 1
                End If
                .NoTrans = Me.TxtIdNumber.Text
                .Tanggal = Me.dtpTanggalTransaksi.Value
                .PrimarySession = Me.TxtPrimarySession.Text
                .PrimaryMember = Me.TxtPrimaryMember.Text
                .PrimaryTransServer = -1
                .PrimaryDisc = -1
                .DiscStr = ObjGeneralSet.GetPaymentDiskon()
                .PrimaryAccountDiscCost = -1
                .UseTax = ObjGeneralSet.GetUseTax()
                .TaxPersen = ObjGeneralSet.GetPersenTax()
                .IsSinglePayment = 0
                vUnitPrice = ObjProd.GetUnitPrice(ObjGeneralSet.GetPrimaryCurr(), Me.TxtPrimaryCodeNo.Text)
                If IndexDGV <> -1 Then
                    vDiscStringItemInputDGV = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), _
                    Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, CDec(Me.TxtQty.Text) + CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value), Me.TBLDiskon, 0)
                    vSubTotalItem = ObjCD.DigitGrouping((Me.TxtQty.Text + CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value)) * vUnitPrice, DefPrec, Me.RoundingTipe)
                    vDiscAmountItemInputDGV = ObjCD.DigitGrouping(CDec(ObjCD.Get_Disc_Amount(vDiscStringItemInputDGV, vSubTotalItem, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD) / (CDec(Me.TxtQty.Text) + CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value))), Me.DefPrec, Me.RoundingTipe)
                    'Hitung lagi dengan perkalian (vSubTotalItem = qty * (unit Price - Diskon))
                    vSubTotalItem = ObjCD.DigitGrouping((Me.TxtQty.Text + CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value)) * (vUnitPrice - vDiscAmountItemInputDGV), DefPrec, Me.RoundingTipe)
                    vSubTotalDGV = ObjCD.DigitGrouping(Me.HitungSubTotalDGV(IndexDGV), DefPrec, Me.RoundingTipe)
                Else
                    vDiscStringItemInputDGV = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), _
                    Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, Me.TxtQty.Text, Me.TBLDiskon, 0)
                    vSubTotalItem = ObjCD.DigitGrouping(Me.TxtQty.Text * vUnitPrice, DefPrec, Me.RoundingTipe)
                    vDiscAmountItemInputDGV = ObjCD.DigitGrouping(ObjCD.Get_Disc_Amount(vDiscStringItemInputDGV, vSubTotalItem, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD), Me.DefPrec, Me.RoundingTipe)
                    vSubTotalDGV = ObjCD.DigitGrouping(Me.HitungSubTotalDGV(), DefPrec, Me.RoundingTipe)
                    vSubTotalItem = vSubTotalItem - vDiscAmountItemInputDGV
                End If
                .Total = ObjCD.DigitGrouping(vSubTotalItem + vSubTotalDGV, DefPrec, Me.RoundingTipe)
                .Discount = ObjCD.Get_Disc_Amount(.DiscStr, .Total, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .Subtotal = ObjCD.DigitGrouping(.Total - .Discount, DefPrec, Me.RoundingTipe)
                If .TaxPersen = 0 Then
                    .TaxAmount = 0
                Else
                    .TaxAmount = ObjCD.DigitGrouping(.TaxPersen / 100 * .Subtotal, DefPrec, Me.RoundingTipe)
                End If
                If ObjPayment.GetBankChargeOption(ObjGeneralSet.GetPrimaryPayment()) = 1 Then
                    If ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) <> 0 Then
                        .BankCharge = ObjCD.DigitGrouping((.TaxAmount + .Subtotal) * ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) / 100, DefPrec, Me.RoundingTipe)
                    Else
                        '.BankCharge = 0
                        .BankCharge = ObjCD.DigitGrouping(ObjPayment.GetBankChargeNominal(ObjGeneralSet.GetPrimaryPayment()), Me.DefPrec, Me.RoundingTipe)
                    End If
                Else
                    .BankCharge = 0
                End If
                vGrandTotalBelumRounding = .Subtotal + .TaxAmount + .BankCharge
                vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .GrandTotal = vGrandTotalRounding
                .Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, DefPrec, Me.RoundingTipe)
                .PrimaryCurr = Me.PrimaryCurr
                .DefPrec = Me.DefPrec
                '-----
                .TotalPayment = 0
                .Change = 0
                .IsVoid = 0
                .Remarks = ""
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'sementara belum di pakai
    Public Sub SalinKeObjekUntukFungsiInputKeyDownTxtQty(ByVal IndexDGV As Integer)
        Dim vUnitPrice As Decimal = 0
        Dim vSubTotalItem As Decimal = 0
        Dim vSubTotalDGV As Decimal = 0
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Dim vDiscStringItemInputDGV As String = 0
        Dim vDiscAmountItemInputDGV As Decimal = 0
        Try
            'Single Payment
            With Me.Obj
                If IsNumeric(Me.TxtQty.Text) = False Then
                    Me.TxtQty.Text = 1
                End If
                .NoTrans = Me.TxtIdNumber.Text
                .Tanggal = Me.dtpTanggalTransaksi.Value
                .PrimarySession = Me.TxtPrimarySession.Text
                .PrimaryMember = Me.TxtPrimaryMember.Text
                .PrimaryTransServer = -1
                .PrimaryDisc = -1
                .DiscStr = Me.ObjGeneralSet.GetPaymentDiskon()
                .PrimaryAccountDiscCost = -1
                .UseTax = Me.ObjGeneralSet.GetUseTax()
                .TaxPersen = Me.ObjGeneralSet.GetPersenTax()
                .IsSinglePayment = 0
                vUnitPrice = Me.ObjProd.GetUnitPrice(Me.ObjGeneralSet.GetPrimaryCurr(), Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, IndexDGV).Value)
                '----
                vDiscStringItemInputDGV = Me.ObjDiskon.Get_Disc_String(Me.ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, IndexDGV).Value, Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, CDec(Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value), Me.TBLDiskon, 0)
                vSubTotalItem = Me.ObjCD.DigitGrouping((Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value) * vUnitPrice, Me.DefPrec, Me.RoundingTipe)
                vDiscAmountItemInputDGV = ObjCD.DigitGrouping(ObjCD.Get_Disc_Amount(vDiscStringItemInputDGV, vSubTotalItem, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD), Me.DefPrec, Me.RoundingTipe)
                'Hitung lagi dengan perkalian (vSubTotalItem = qty * (unit Price - Diskon))
                vSubTotalItem = Me.ObjCD.DigitGrouping((Me.dgvItem.Item(Me.ColQty.Index, IndexDGV).Value) * (vUnitPrice - vDiscAmountItemInputDGV), DefPrec, Me.RoundingTipe)
                vSubTotalDGV = Me.ObjCD.DigitGrouping(Me.HitungSubTotalDGV(IndexDGV), DefPrec, Me.RoundingTipe)
                '----
                .Total = Me.ObjCD.DigitGrouping(vSubTotalItem + vSubTotalDGV, Me.DefPrec, Me.RoundingTipe)
                .Discount = Me.ObjCD.Get_Disc_Amount(.DiscStr, .Total, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .Subtotal = Me.ObjCD.DigitGrouping(.Total - .Discount, Me.DefPrec, Me.RoundingTipe)
                If .TaxPersen = 0 Then
                    .TaxAmount = 0
                Else
                    .TaxAmount = ObjCD.DigitGrouping(.TaxPersen / 100 * .Subtotal, DefPrec, Me.RoundingTipe)
                End If
                If ObjPayment.GetBankChargeOption(ObjGeneralSet.GetPrimaryPayment()) = 1 Then
                    If ObjPayment.GetUserPrecentage(Me.ObjGeneralSet.GetPrimaryPayment()) <> 0 Then
                        .BankCharge = ObjCD.DigitGrouping((.TaxAmount + .Subtotal) * ObjPayment.GetUserPrecentage(Me.ObjGeneralSet.GetPrimaryPayment()) / 100, DefPrec, Me.RoundingTipe)
                    Else
                        '.BankCharge = 0
                        .BankCharge = ObjCD.DigitGrouping(ObjPayment.GetBankChargeNominal(ObjGeneralSet.GetPrimaryPayment()), Me.DefPrec, Me.RoundingTipe)
                    End If
                Else
                    .BankCharge = 0
                End If
                vGrandTotalBelumRounding = .Subtotal + .TaxAmount + .BankCharge
                vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .GrandTotal = vGrandTotalRounding
                .Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.RoundingTipe)
                .PrimaryCurr = Me.PrimaryCurr
                .DefPrec = Me.DefPrec
                '-----
                .TotalPayment = 0
                .Change = 0
                .IsVoid = 0
                .Remarks = ""
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObjekUntukPayment()
        Dim vUnitPrice As Decimal = 0
        Dim vSubTotalItem As Decimal = 0
        Dim vSubTotalDGV As Decimal = 0
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Dim vDiscStringItemInputDGV As String = 0
        Dim vDiscAmountItemInputDGV As Decimal = 0
        Try
            'Single Payment
            With Me.Obj
                If IsNumeric(Me.TxtQty.Text) = False Then
                    Me.TxtQty.Text = 1
                End If
                .NoTrans = Me.TxtIdNumber.Text
                .Tanggal = Me.dtpTanggalTransaksi.Value
                .PrimarySession = Me.TxtPrimarySession.Text
                .PrimaryMember = Me.TxtPrimaryMember.Text
                .PrimaryTransServer = -1
                .PrimaryDisc = -1
                .DiscStr = ObjGeneralSet.GetPaymentDiskon()
                .PrimaryAccountDiscCost = -1
                .UseTax = ObjGeneralSet.GetUseTax()
                .TaxPersen = ObjGeneralSet.GetPersenTax()
                .IsSinglePayment = 0
                vUnitPrice = ObjProd.GetUnitPrice(ObjGeneralSet.GetPrimaryCurr(), Me.TxtPrimaryCodeNo.Text)
                vDiscStringItemInputDGV = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), _
                                        Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, _
                                        Me.TxtQty.Text, Me.TBLDiskon, 0)
                vSubTotalItem = ObjCD.DigitGrouping(Me.TxtQty.Text * vUnitPrice, DefPrec, Me.RoundingTipe)
                vDiscAmountItemInputDGV = ObjCD.DigitGrouping(ObjCD.Get_Disc_Amount(vDiscStringItemInputDGV, _
                                        vSubTotalItem, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD), _
                                        Me.DefPrec, Me.RoundingTipe)
                vSubTotalItem = vSubTotalItem - vDiscAmountItemInputDGV
                vSubTotalDGV = ObjCD.DigitGrouping(Me.HitungSubTotalDGV(), DefPrec, Me.RoundingTipe)
                .Total = ObjCD.DigitGrouping(vSubTotalItem + vSubTotalDGV, DefPrec, Me.RoundingTipe)
                .Discount = ObjCD.Get_Disc_Amount(.DiscStr, .Total, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .Subtotal = ObjCD.DigitGrouping(.Total - .Discount, DefPrec, Me.RoundingTipe)
                If .TaxPersen = 0 Then
                    .TaxAmount = 0
                Else
                    .TaxAmount = ObjCD.DigitGrouping(.TaxPersen / 100 * .Subtotal, DefPrec, Me.RoundingTipe)
                End If
                If ObjPayment.GetBankChargeOption(ObjGeneralSet.GetPrimaryPayment()) = 1 Then
                    If ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) <> 0 Then
                        .BankCharge = ObjCD.DigitGrouping((.TaxAmount + .Subtotal) * ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) / 100, DefPrec, Me.RoundingTipe)
                    Else
                        '.BankCharge = 0
                        .BankCharge = ObjCD.DigitGrouping(ObjPayment.GetBankChargeNominal(ObjGeneralSet.GetPrimaryPayment()), Me.DefPrec, Me.RoundingTipe)
                    End If
                Else
                    .BankCharge = 0
                End If
                vGrandTotalBelumRounding = .Subtotal + .TaxAmount + .BankCharge
                vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .GrandTotal = vGrandTotalRounding
                .Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, DefPrec, Me.RoundingTipe)
                .PrimaryCurr = Me.PrimaryCurr
                .DefPrec = Me.DefPrec
                '-----
                .TotalPayment = 0
                .Change = 0
                .IsVoid = 0
                .Remarks = ""
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObjekUntukUpdateDataItem(ByVal RowDGV As DataGridViewRow)
        Dim vUnitPrice As Decimal = 0
        Dim vSubTotalItem As Decimal = 0
        Dim vSubTotalDGV As Decimal = 0
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Dim vDiscStringItemInputDGV As String = 0
        Dim vDiscAmountItemInputDGV As Decimal = 0
        Try
            'Single Payment
            With Me.Obj
                '.NoTrans = Me.TxtIdNumber.Text
                '.Tanggal = Me.dtpTanggalTransaksi.Value
                '.PrimarySession = Me.TxtPrimarySession.Text
                '.PrimaryTransServer = -1
                '.PrimaryDisc = -1
                '.DiscStr = ObjGeneralSet.GetPaymentDiskon()
                '.PrimaryAccountDiscCost = -1
                '.UseTax = ObjGeneralSet.GetUseTax()
                '.TaxPersen = ObjGeneralSet.GetPersenTax()
                '.IsSinglePayment = 0
                If IsNumeric(Me.TxtQty.Text) = False Then
                    Me.TxtQty.Text = 1
                End If
                vUnitPrice = ObjProd.GetUnitPrice(ObjGeneralSet.GetPrimaryCurr(), Me.TxtPrimaryCodeNo.Text)
                .PrimaryMember = Me.TxtPrimaryMember.Text
                vDiscStringItemInputDGV = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(Me.TxtPrimarySession.Text), _
                RowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.TxtPrimaryMember.Text, Me.dtpTanggalTransaksi.Value, RowDGV.Cells(Me.ColQty.Index).Value, Me.TBLDiskon, 0)
            
                vSubTotalItem = ObjCD.DigitGrouping(RowDGV.Cells(Me.ColQty.Index).Value * vUnitPrice, DefPrec, Me.RoundingTipe)

                vDiscAmountItemInputDGV = ObjCD.DigitGrouping(ObjCD.Get_Disc_Amount(vDiscStringItemInputDGV, vSubTotalItem, Me.DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD), Me.DefPrec, Me.RoundingTipe)
                vSubTotalItem = ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQty.Index, RowDGV.Index).Value) * (vUnitPrice - (vDiscAmountItemInputDGV / CDec(Me.dgvItem.Item(Me.ColQty.Index, RowDGV.Index).Value))), DefPrec, Me.RoundingTipe)
                vSubTotalItem = vSubTotalItem - vDiscAmountItemInputDGV
                vSubTotalDGV = ObjCD.DigitGrouping(Me.HitungSubTotalDGV(), DefPrec, Me.RoundingTipe)
                .Total = ObjCD.DigitGrouping(vSubTotalItem + vSubTotalDGV, DefPrec, Me.RoundingTipe)
                .Discount = ObjCD.Get_Disc_Amount(.DiscStr, .Total, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .Subtotal = ObjCD.DigitGrouping(.Total - .Discount, DefPrec, Me.RoundingTipe)
                If .TaxPersen = 0 Then
                    .TaxAmount = 0
                Else
                    .TaxAmount = ObjCD.DigitGrouping(.TaxPersen / 100 * .Subtotal, DefPrec, Me.RoundingTipe)
                End If
                If ObjPayment.GetBankChargeOption(ObjGeneralSet.GetPrimaryPayment()) = 1 Then
                    If ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) <> 0 Then
                        .BankCharge = ObjCD.DigitGrouping((.TaxAmount + .Subtotal) * ObjPayment.GetUserPrecentage(ObjGeneralSet.GetPrimaryPayment()) / 100, DefPrec, Me.RoundingTipe)
                    Else
                        '.BankCharge = 0
                        .BankCharge = ObjCD.DigitGrouping(ObjPayment.GetBankChargeNominal(ObjGeneralSet.GetPrimaryPayment()), Me.DefPrec, Me.RoundingTipe)
                    End If
                Else
                    .BankCharge = 0
                End If
                vGrandTotalBelumRounding = .Subtotal + .TaxAmount + .BankCharge
                vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.RoundingTipe, Me.RoundingValue, Me.ObjCD)
                .GrandTotal = vGrandTotalRounding
                .Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, DefPrec, Me.RoundingTipe)
                .PrimaryCurr = Me.PrimaryCurr
                .DefPrec = Me.DefPrec
                '-----
                .TotalPayment = 0
                .Change = 0
                .IsVoid = 0
                .Remarks = ""
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtIdNumber.Text = .NoTrans

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Salin Ke Field -----> ini yang sesuai dengan standar
    Public Sub SalinKeFieldUntukFind()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtIdNumber.Text = .NoTrans
                Me.TxtPrimarySession.Text = .PrimarySession
                Me.TxtQty.Text = "1"
                Me.TxtPrimaryMember.Text = .PrimaryMember
                Me.TxtMember.Text = ObjMember.GetCode(Me.TxtPrimaryMember.Text)
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNo.Text = ""

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.TxtDisplay.Font = Me.FontSubTotal
            Me.TxtDisplay.Text = "Total Sales: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()).ToString + ". " + ObjCD.DigitGrouping(Me.HitungSubTotalDGV().ToString, Me.DefPrec, Me.RoundingTipe)
            Me.Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMember_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMember.GotFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtMember.BackColor = Me.BackColorAll
                    Else
                        TxtMember.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMember_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMember.LostFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtMember.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtMember.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.GotFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtQty.Enabled = True Then
                    If Me.TxtQty.ReadOnly = False Then
                        TxtQty.BackColor = Me.BackColorAll
                        TxtQty.SelectAll()
                    Else
                        TxtQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.LostFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtQty.Enabled = True Then
                    If Me.TxtQty.ReadOnly = False Then
                        TxtQty.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtQty.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtCodeNo.BackColor = Me.BackColorAll
                    Else
                        TxtCodeNo.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtCodeNo.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtCodeNo.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtUnitPrice.BackColor = Me.BackColorAll
                    Else
                        TxtUnitPrice.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtUnitPrice.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtUnitPrice.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.GotFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtDiscString.BackColor = Me.BackColorAll
                    Else
                        TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.LostFocus
        Try
            If Me.TxtQty.Enabled = True Then
                If Me.TxtMember.Enabled = True Then
                    If Me.TxtMember.ReadOnly = False Then
                        TxtDiscString.BackColor = Me.TxtPrimary.BackColor
                    Else
                        TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnReprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReprint.Click
        Dim i As New FrmTrans_POS_Cetak
        Dim index As Integer = 0
        Try
            If ObjSession.GetStatus(Me.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> "-1" Then
                If ObjPOS.GetIsVoid(Me.TxtPrimary.Text) <= 0 Then
                    If ObjPOS.GetJmlPayment(Me.TxtPrimary.Text) > 0 Then
                        Me.ObjPOS.CetakPOSUkuranStngahLetter(i.CRV1, Me.UserCode, Me.TxtPrimary.Text)
                        'i.ShowDialog()
                    Else
                        MsgBox("Transaction POS is unpaid, you cannot print receipt.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQty.TextChanged

    End Sub
End Class