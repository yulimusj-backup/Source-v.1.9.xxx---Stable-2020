Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmMasterAkun_Detail
    Public UserCode As String
    Public BackColorAll As Color = Color.LightSeaGreen
    Public ObjAkun As New ClsMasterAkun
    Public Baru As Boolean
    Public Def_Prec As Integer = 0
    Public Def_PrecOriginal As Integer = 0

    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim FrmMasterAkun As FrmMasterAkun
    Dim Ok As Boolean
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim TBLCur As New Core_Module.dseMasterCurrency.tblmaster_currencyDataTable

    Private Sub TxtBeginningBalance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalance.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtBeginningBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalanceOriginalCur.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtBeginningBalanceOriginalCur, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
  Handles TxtBeginningBalanceOriginalCur.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtBeginningBalanceOriginalCur, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBeginningBalanceOriginalCur.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtBeginningBalanceOriginalCur, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalance_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBeginningBalance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtBeginningBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtBeginningBalance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtBeginningBalance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkHeaderAccount_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHeaderAccount.CheckedChanged
        Try
            Me.SetTxtBeginningBalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            'ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, True)
            'ModCekData.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryAccount.BackColor, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.TxtAccountNo.Focus()
            'Me.cboAccountType.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            'ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, True, False)
            'ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            Me.TxtPrimaryAccount.Text = "-1"
            Me.TxtBeginningBalance.Text = Me.ObjCD.DigitGrouping("0", Me.Def_Prec)
            Me.chkActiveAccount.Checked = False
            Me.chkHeaderAccount.Checked = False
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

        Try
            'cek data yang harus diisi
            If Me.ObjAkun.LevelPerkiraan <> 1 Then
                KoleksiNamaData.Add("Account Number")
                KoleksiTextBox.Add(Me.TxtAccountNo)
            End If
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)


            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            ' header ato bukan
            If Me.chkHeaderAccount.Checked = True Then
                If Me.TxtBeginningBalance.Text <> 0 Then
                    MsgBox("This is data header, beginning balance = 0", MsgBoxStyle.Information, "information")
                    Return False
                End If
            End If

            'cek data yang salah, double dalam kode
            Dim AkunCode As String = ""
            AkunCode = Me.TxtKodeAkun1.Text + Me.TxtAccountNo.Text
            If Me.baru = True Then
                If Objakun.ApakahSudahAda(AkunCode) = True Then
                    MsgBox("Duplication in account code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtAccountNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Objakun.GetAccountCode(Me.TxtPrimaryAccount.Text)
                If TipeCodeLama <> AkunCode Then
                    If Me.Objakun.ApakahSudahAda(AkunCode) = True Then
                        MsgBox("Duplication in account code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtAccountNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Objakun.GetTimestamp(Me.TxtPrimaryAccount.Text)
                If Me.Objakun.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Objakun.CariDetail(Me.Objakun.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek statu aktif akun ketika perubahan data
            'jika akun aktif maka:
            '1. jika non header --> tidak ada yang perlu dicek
            '2. jika header --> tidak ada yang perlu dicek
            If Me.Baru = False Then
                'jika akun tidak aktif:
                If Me.chkActiveAccount.Checked = False Then
                    If Me.ObjAkun.GetHeader(Me.TxtPrimaryAccount.Text) = 0 Then
                        '1. jika akun non-header --> cek berkaitan dengan transaksi setting transfer
                        '   a. jika sudah digunakan pada setting transfer invoice
                        Dim ObjSett_Inv As New ClsTransaksiSettingTransf_Inv
                        If ObjSett_Inv.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in invoice transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_Inv = Nothing
                        '   b. jika sudah digunakan pada setting transfer invoice DP
                        Dim ObjSett_InvDP As New ClsTransaksiSettingTransf_InvDP
                        If ObjSett_InvDP.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in downpayment invoice transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_InvDP = Nothing
                        '   c. Jika sudah digunakan pada setting transder sales return
                        Dim ObjSett_SR As New ClsTransaksiSettingTransf_SR
                        If ObjSett_SR.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in sales return transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_SR = Nothing
                        '   d. Jika sudah digunakan pada setting transfer AOP
                        Dim ObjSett_AOP As New ClsTransaksiSettingTransf_AOP
                        If ObjSett_AOP.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in AOP transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_AOP = Nothing
                        '   e. jika sudah digunakan pada setting transfer AOP-DP
                        Dim ObjSett_AOP_DP As New ClsTransaksiSettingTransf_AOPDP
                        If ObjSett_AOP_DP.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in downpayment AOP transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_AOP_DP = Nothing
                        '   f. jika sudah digunakan pada setting transfer AOPR
                        Dim ObjSett_AOPR As New ClsTransaksiSettingTrans_AOPR
                        If ObjSett_AOPR.GetJml_ByPrimaryAkun(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("This account is used in downpayment AOP transfer, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        ObjSett_AOPR = Nothing
                    Else
                        '2. jika akun header --> 
                        '   a. cek apakah ada akun aktif yang merujuk ke akun header ybs, jika ada maka
                        '   akun header tidak bisa non-aktif
                        If Me.ObjAkun.GetJmlAnak_Aktif(Me.TxtPrimaryAccount.Text) > 0 Then
                            MsgBox("There active accounts refering to this account, you can not set this account inactive", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            'kalo nol, maka keduanya harus nol, untuk currency non idr
            If Me.TxtPrimaryCurrency.Text <> Core_Module.PrimaryBaseCurrency Then
                If CDec(Me.TxtBeginningBalance.Text) = 0 Then
                    If CDec(Me.TxtBeginningBalanceOriginalCur.Text) <> 0 Then
                        MsgBox("Inconsistent beginning balance.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'kalo positif, maka keduanya harus positif
                If CDec(Me.TxtBeginningBalance.Text) > 0 Then
                    If CDec(Me.TxtBeginningBalanceOriginalCur.Text) <= 0 Then
                        MsgBox("Inconsistent beginning balance.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'kalo negatif, maka keduanya harus negatif
                If CDec(Me.TxtBeginningBalance.Text) < 0 Then
                    If CDec(Me.TxtBeginningBalanceOriginalCur.Text) >= 0 Then
                        MsgBox("Inconsistent beginning balance.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                Me.TxtBeginningBalanceOriginalCur.Text = Me.TxtBeginningBalance.Text
            End If


            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Dim spasi As String = "     "

        Try
            With Me.ObjAkun
                'Masukan data master akun detail
                .AccountCode = Format(Me.TxtKodeAkun1.Text + Me.TxtAccountNo.Text, "")
                .Description = Me.TxtDesc.Text
                If .LevelPerkiraan = 1 Then
                    .Tampilan = .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 2 Then
                    .Tampilan = spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 3 Then
                    .Tampilan = spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 4 Then
                    .Tampilan = spasi + spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 5 Then
                    .Tampilan = spasi + spasi + spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 6 Then
                    .Tampilan = spasi + spasi + spasi + spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 7 Then
                    .Tampilan = spasi + spasi + spasi + spasi + spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 8 Then
                    .Tampilan = spasi + spasi + spasi + spasi + spasi + spasi + spasi + .AccountCode + " - " + .Description
                End If
                If Me.chkActiveAccount.Checked = True Then
                    .Aktif = 1
                Else
                    .Aktif = 0
                End If
                If Me.chkHeaderAccount.Checked = True Then
                    .Header = 1
                Else
                    .Header = 0
                End If
                .BeginningBalance = Me.TxtBeginningBalance.Text
                .AccountNature = Me.TxtRemarks.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                If Me.TxtPrimaryCurrency.Text = 1 Then
                    Me.TxtBeginningBalanceOriginalCur.Text = Me.TxtBeginningBalance.Text
                End If
                .BeginningBalanceOriginalCurrency = Me.TxtBeginningBalanceOriginalCur.Text
                .TipeAkun2 = Me.ObjAkun.GetTipeAkun2(Me.cboAccountType.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Try
            Me.Def_PrecOriginal = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtBeginningBalance()
        Try
            Me.LblIDR.Text = Me.ObjCur.GetCurrencyCode(1)
            If Me.TxtPrimaryCurrency.Text = "" Then
                Me.TxtPrimaryCurrency.Text = "1"
            End If
            Me.LblOriginalCur.Text = "In " + Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)

            If Me.chkHeaderAccount.Checked = True Then
                Me.TxtBeginningBalance.Enabled = False
                Me.TxtBeginningBalance.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
                Me.TxtBeginningBalanceOriginalCur.Enabled = False
                Me.TxtBeginningBalanceOriginalCur.Text = Me.ObjCD.DigitGrouping(0, Me.Def_Prec)
            Else
                Me.TxtBeginningBalance.Enabled = True
                Me.TxtBeginningBalanceOriginalCur.Enabled = True
            End If

            If Me.TxtPrimaryCurrency.Text = "1" Then
                Me.TxtBeginningBalance.Width = Me.cboCurrency.Left + Me.cboCurrency.Width - Me.TxtBeginningBalance.Left
                Me.LblOriginalCur.Visible = False
                Me.TxtBeginningBalanceOriginalCur.Visible = False
                Me.TxtBeginningBalanceOriginalCur.TabStop = False
            Else
                Me.TxtBeginningBalance.Width = Me.LblOriginalCur.Left - 10 - Me.TxtBeginningBalance.Left
                Me.LblOriginalCur.Visible = True
                Me.TxtBeginningBalanceOriginalCur.Visible = True
                Me.TxtBeginningBalanceOriginalCur.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Dim row As Core_Module.dseMasterCurrency.tblmaster_currencyRow
        Dim DV As DataRowView
        Dim i As Integer = 0

        Try
            Me.ObjAkun.IsiCBOTipeAkun2(Me.ObjAkun.TipePerkiraan, Me.cboAccountType, Me.ObjAkun.LevelPerkiraan)
            Me.cboAccountType.SelectedIndex = 0
            Me.ObjCur.IsiCBO(Me.cboCurrency, Me.TBLCur)
            Me.cboCurrency.SelectedIndex = -1

            With Me.ObjAkun
                Me.TxtPrimaryAccount.Text = .Primary
                If Me.Baru = False Then
                    Me.TxtAccountNo.Text = .AccountCode.Remove(0, 1)
                    Me.TxtKodeAkun1.Text = .AccountCode.Substring(0, 1)
                End If
                Me.TxtDesc.Text = .Description
                If .Aktif = 0 Then
                    Me.chkActiveAccount.Checked = False
                Else
                    Me.chkActiveAccount.Checked = True
                End If
                If .Header = 0 Then
                    Me.chkHeaderAccount.Checked = False
                Else
                    Me.chkHeaderAccount.Checked = True
                End If
                Me.TxtRemarks.Text = .AccountNature
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.TxtPrimaryCurrency.Text = Me.ObjAkun.PrimaryCurrency
                For Each DV In Me.TBLCur.DefaultView
                    row = DV.Row
                    If row.Primary = Me.TxtPrimaryCurrency.Text Then
                        Exit For
                    End If
                    i += 1
                Next
                Me.cboCurrency.SelectedIndex = i
                Me.TxtBeginningBalance.Text = ObjCD.DigitGrouping(.BeginningBalance, Me.Def_Prec)
                Me.TxtBeginningBalanceOriginalCur.Text = Me.ObjCD.DigitGrouping(.BeginningBalanceOriginalCurrency, Me.Def_PrecOriginal)
                Me.cboAccountType.Text = Me.ObjAkun.GetTipeAkun2(Me.ObjAkun.TipeAkun2)

                If Me.Baru = False Then
                    'jika sudah ada jurnal yang merujuk pada akun tersebut maka:
                    '1. Header account tidak bisa diubah
                    '2. Currency tidak bisa diubah
                    Dim ObjJurnal As New ClsTransaksiJournal
                    If ObjJurnal.GetJml_ByPrimaryAkun(Me.ObjAkun.Primary) > 0 Then
                        Me.chkHeaderAccount.Enabled = False
                        Me.cboCurrency.Enabled = False
                        Me.cboAccountType.Enabled = False
                    End If
                    ObjJurnal = Nothing

                    'jika sudah ada anak yang merujuk pada akun ini maka, chk header juga tidak bisa diubah
                    If Me.ObjAkun.GetJmlAnak(Me.ObjAkun.Primary) > 0 Then
                        Me.chkHeaderAccount.Enabled = False
                    End If

                    'jika sudah dirujuk oleh master currency maka cbo currency tidak bisa diedit
                    '1. untuk primary account payable pd master currency
                    If Me.ObjCur.GetJmlByPrimaryAcc_AccountPayable(Me.ObjAkun.Primary) > 0 Then
                        Me.cboCurrency.Enabled = False
                        Me.cboAccountType.Enabled = False
                        Me.chkActiveAccount.Enabled = False
                        Me.chkActiveAccount.Checked = True
                        Me.chkHeaderAccount.Enabled = False
                        Me.chkHeaderAccount.Checked = False
                    End If
                    '2. untuk primary account receivable pd master currency
                    If Me.ObjCur.GetJmlByPrimaryAcc_AccountReceivable(Me.ObjAkun.Primary) > 0 Then
                        Me.cboCurrency.Enabled = False
                        Me.cboAccountType.Enabled = False
                        Me.chkActiveAccount.Enabled = False
                        Me.chkActiveAccount.Checked = True
                        Me.chkHeaderAccount.Enabled = False
                        Me.chkHeaderAccount.Checked = False
                    End If
                    If Me.ObjCur.GetJmlByPrimaryAcc_RealizedExchangeGainLossAccount(Me.ObjAkun.Primary) > 0 Then
                        Me.chkActiveAccount.Enabled = False
                        Me.chkActiveAccount.Checked = True
                        Me.chkHeaderAccount.Enabled = False
                        Me.chkHeaderAccount.Checked = False
                    End If
                    If Me.ObjCur.GetJmlByPrimaryAcc_UnrealizedExchangeGainLossAccount(Me.ObjAkun.Primary) > 0 Then
                        Me.chkActiveAccount.Enabled = False
                        Me.chkActiveAccount.Checked = True
                        Me.chkHeaderAccount.Enabled = False
                        Me.chkHeaderAccount.Checked = False
                    End If

                    'jika sudah ada jurnal yang merujuk, pada cbo currency tidak bisa diedit

                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Objakun.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            ''Me.NonAktifkanDetail()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Close()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Objakun.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Ok = True
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Not Me.TxtAccountNo.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnCancel.Focus()
            Me.Ok = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkunDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.Ok = False Then
                If Me.TxtAccountNo.Enabled = True Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If Baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjAkun.SimpanBaru(Me.UserCode) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.DialogResult = Windows.Forms.DialogResult.OK
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjAkun.SimpanPerubahan(Me.UserCode) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.DialogResult = Windows.Forms.DialogResult.OK
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
            End If
            
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWeightBalanceCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountNo.GotFocus
        Try
            If Me.TxtAccountNo.Enabled = True Then
                If Me.TxtAccountNo.ReadOnly = False Then
                    TxtAccountNo.BackColor = Me.BackColorAll
                Else
                    TxtAccountNo.BackColor = Me.TxtKodeAkun1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWeightBalanceCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountNo.LostFocus
        Try
            If Me.TxtAccountNo.Enabled = True Then
                If Me.TxtAccountNo.ReadOnly = False Then
                    TxtAccountNo.BackColor = Me.TxtPrimaryAccount.BackColor
                Else
                    TxtAccountNo.BackColor = Me.TxtKodeAkun1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIdentificationNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            If Me.TxtDesc.Enabled = True Then
                If Me.TxtDesc.ReadOnly = False Then
                    TxtDesc.BackColor = Me.BackColorAll
                Else
                    TxtDesc.BackColor = Me.TxtKodeAkun1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIdentificationNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            If Me.TxtDesc.Enabled = True Then
                If Me.TxtDesc.ReadOnly = False Then
                    TxtDesc.BackColor = Me.TxtPrimaryAccount.BackColor
                Else
                    TxtDesc.BackColor = Me.TxtKodeAkun1.BackColor
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
                    TxtRemarks.BackColor = Me.TxtKodeAkun1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryAccount.BackColor
                Else
                    TxtRemarks.BackColor = Me.TxtKodeAkun1.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalance.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginningBalance.Enabled = True Then
                    If Me.TxtBeginningBalance.ReadOnly = False Then
                        TxtBeginningBalance.BackColor = Me.BackColorAll
                    Else
                        TxtBeginningBalance.BackColor = Me.TxtKodeAkun1.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalance.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBeginningBalance.Enabled = True Then
                    If Me.TxtBeginningBalance.ReadOnly = False Then
                        TxtBeginningBalance.BackColor = Me.TxtPrimaryAccount.BackColor
                    Else
                        TxtBeginningBalance.BackColor = Me.TxtKodeAkun1.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccountNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAccountNo.TextChanged
        Try
            Me.SetText()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetText()
        Try
            Me.Text = "Account Detail: " + Me.TxtKodeAkun1.Text + Me.TxtAccountNo.Text + " - " + Me.TxtDesc.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeAkun1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeAkun1.GotFocus
        Try
            Me.TxtAccountNo.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeAkun1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtKodeAkun1.TextChanged
        Try
            Me.SetText()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDesc.TextChanged
        Try
            Me.SetText()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkun_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.AktifkanDetail()
            Me.TxtAccountNo.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalance.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtBeginningBalance, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalanceOriginalCur.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtBeginningBalanceOriginalCur, sender, e, Me.Def_PrecOriginal, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrency_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrency.Leave
        Try
            Me.TxtBeginningBalanceOriginalCur.Text = Me.ObjCD.DigitGrouping(Me.TxtBeginningBalanceOriginalCur.Text, Me.Def_PrecOriginal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrency_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrency.SelectedValueChanged
        Dim row As Core_Module.dseMasterCurrency.tblmaster_currencyRow
        Dim DV As DataRowView

        Try
            If Me.cboCurrency.SelectedIndex = -1 Then
                Me.TxtPrimaryCurrency.Text = -1
            Else
                DV = Me.TBLCur.DefaultView(Me.cboCurrency.SelectedIndex)
                row = DV.Row
                Me.TxtPrimaryCurrency.Text = row.Primary
            End If

            Me.SetTxtBeginningBalance()
            Me.SetDefaultPrecision()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalanceOriginalCur.GotFocus
        Try
            If Me.TxtBeginningBalanceOriginalCur.Enabled = True Then
                If Me.TxtBeginningBalanceOriginalCur.ReadOnly = True Then
                    Me.TxtBeginningBalanceOriginalCur.BackColor = Me.TxtKodeAkun1.BackColor
                Else
                    Me.TxtBeginningBalanceOriginalCur.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBeginningBalanceOriginalCur_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBeginningBalanceOriginalCur.LostFocus
        Try
            If Me.TxtBeginningBalanceOriginalCur.Enabled = True Then
                If Me.TxtBeginningBalanceOriginalCur.ReadOnly = True Then
                    Me.TxtBeginningBalanceOriginalCur.BackColor = Me.TxtKodeAkun1.BackColor
                Else
                    Me.TxtBeginningBalanceOriginalCur.BackColor = Me.TxtRemarks.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkun_Detail_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        Try
            If Me.Visible = True Then
                Me.AktifkanDetail()
                Me.TxtAccountNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccountType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccountType.SelectedValueChanged
        Try
            If Me.ObjAkun.ApakahBolehGunakanForeignCurrency(Me.ObjAkun.GetTipeAkun2(Me.cboAccountType.Text)) = False Then
                Me.cboCurrency.Text = Me.ObjCur.GetCurrencyCode(1)
                Me.cboCurrency.Enabled = False
                Me.cboCurrency.TabStop = False
            Else
                Me.cboCurrency.Enabled = True
                Me.cboCurrency.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class