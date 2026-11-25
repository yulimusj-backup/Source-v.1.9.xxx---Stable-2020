Option Explicit On
Option Strict Off

Public Class FrmMasterSupplier
    Public UserCode As String = ""
    Dim baru As Boolean
    Dim objSupplier As New clsMasterSupplier
    Dim ObjCekData As New ClsCekData
    Dim ObjCurrency As New clsMasterCurrency
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim buffer As String = ""

    'Comment: Perubahan NPWP 16 Digit
    Dim ObjSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryCurrency.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvAddress.ReadOnly = False
            Me.dgvBank.ReadOnly = False
            Me.dgvCP.ReadOnly = False
            Me.dgvAddress.AllowUserToAddRows = True
            Me.dgvBank.AllowUserToAddRows = True
            Me.dgvCP.AllowUserToAddRows = True
            Me.rboYes.Enabled = True
            Me.rboNo.Enabled = True

            'tentukan nilai default untuk baris terakhir
            With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvAddress)
            'tentukan nilai default untuk baris terakhir
            With Me.dgvBank.Rows(Me.dgvBank.Rows.Count - 1)
                .Cells(1).Value = "-"
                .Cells(3).Value = "IDR"
                .Cells(Me.ColPrimaryCurrency.Index).Value = Me.ObjCurrency.GetPrimary("IDR")
            End With
            Me.ObjCekData.SetNumber(Me.dgvBank)
            'tentukan nilai default untuk baris terakhir
            With Me.dgvCP.Rows(Me.dgvCP.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvCP)
            'Comment: Perubahan NPWP 16 Digit
            Me.LlbEditNPWPLama.Enabled = True
            Me.CboTipeNPWP.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvAddress.ReadOnly = True
            Me.dgvBank.ReadOnly = True
            Me.dgvCP.ReadOnly = True
            Me.rboYes.Enabled = False
            Me.rboNo.Enabled = False
            Me.dgvAddress.AllowUserToAddRows = False
            Me.dgvBank.AllowUserToAddRows = False
            Me.dgvCP.AllowUserToAddRows = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            'Comment: Perubahan NPWP 16 Digit
            Me.LlbEditNPWPLama.Enabled = False
            Me.CboTipeNPWP.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, True)

            Me.txtprimarysupplier.Text = "-1"
            Me.TxtPrimaryCurrency.Text = "1"
            Me.TxtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
            Me.rboNo.Checked = True
            '2025 Core Tax Tahap 2.1
            If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                Me.TxtNPWP.Text = "0000 0000 0000 0000"
            Else
                Me.TxtNPWP.Text = "00.000.000.0-000.000"
            End If
            Me.txtLeadTime.Text = "0"
            Me.TxtPaymentTerm.Text = "0"

            Me.dgvAddress.Rows.Clear()
            Me.dgvBank.Rows.Clear()
            Me.dgvCP.Rows.Clear()

            'Comment: NPWP no 16 Digit
            Me.CboTipeNPWP.SelectedIndex = 0
            Me.TxtNPWPLama.Text = Me.TxtNPWP.Text

            'Update PPN 12%
            Me.TxtTKUID.Text = "000000"
            Me.TxtEmail.Text = ""

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
            KoleksiNamaData.Add("supplier code")
            KoleksiTextBox.Add(Me.txtsuppliercode)
            KoleksiNamaData.Add("supplier name")
            KoleksiTextBox.Add(Me.txtsuppliername)
            KoleksiNamaData.Add("currency")
            KoleksiTextBox.Add(Me.TxtCurrencyCode)
            'Perubahan PPN 12%
            KoleksiNamaData.Add("TKU ID")
            KoleksiTextBox.Add(Me.TxtTKUID)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data salah bila:
            '1.tidak ada data address
            If Me.dgvAddress.RowCount = 1 Then
                MsgBox("Please input at least one address data.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Return False
            End If

            '2.cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objSupplier.ApakahSudahAda(Me.txtsuppliercode.Text) = True Then
                    MsgBox("Duplication in supplier code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtsuppliercode.Focus()
                    Return False
                End If
            Else
                Dim SupplierGroupLama As String = Me.objSupplier.GetSupplierCode(Me.txtprimarysupplier.Text)
                If SupplierGroupLama <> Me.txtsuppliercode.Text Then
                    If Me.objSupplier.ApakahSudahAda(Me.txtsuppliercode.Text) = True Then
                        MsgBox("Duplication in supplier code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtsuppliercode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objSupplier.GetTimeStamp(Me.txtprimarysupplier.Text)
                If Me.objSupplier.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objSupplier.CariDetail(Me.objSupplier.Primary, Me.dgvAddress, Me.dgvCP, Me.dgvBank)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. creadit limit/payment terms negatif
            If Me.ObjCekData.CekNonNegatif(Me.txtLeadTime.Text) = False Then
                MsgBox("Lead time must be a non negatif number.", MsgBoxStyle.Information, _
                "Information")
                Me.txtLeadTime.Focus()
                Return False
            End If

            Dim row As DataGridViewRow
            Dim cell As DataGridViewCell
            '4. untuk item address:
            '   a. untuk baris yang bukan baris terakhir, tidak boleh sampai data line 1-5 kosong
            For Each row In Me.dgvAddress.Rows
                With row
                    If row.Index <> Me.dgvAddress.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        If .Cells(2).Value = "" And .Cells(3).Value = "" And _
                        .Cells(4).Value = "" And .Cells(5).Value = "" And .Cells(6).Value = "" Then
                            Me.TabControl1.SelectedIndex = 0
                            MsgBox("Wrong address data, please fill in address data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvAddress.ClearSelection()
                            .Selected = True
                            Return False
                        End If
                    End If
                End With
            Next

            '5. untuk item bank tidak boleh ada nama, currency, dan account no yang kosong
            For Each row In Me.dgvBank.Rows
                With row
                    If row.Index <> Me.dgvBank.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        If .Cells(1).Value = "" Then
                            MsgBox("Wrong bank's name, please fill in bank's name data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                        If .Cells(2).Value = "" Then
                            MsgBox("Wrong account no, please fill in account no data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                        If .Cells(3).Value = "" Then
                            MsgBox("Wrong currency code, please fill in currency code data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvBank.ClearSelection()
                            Me.TabControl1.SelectedIndex = 2
                            .Selected = True
                            Return False
                        End If
                    End If
                End With
            Next

            '6. untuk item contact person, tidak boleh ada nama contact person yang kosong
            For Each row In Me.dgvCP.Rows
                With row
                    If row.Index <> Me.dgvCP.Rows.Count - 1 Then
                        For Each cell In row.Cells
                            If IsNothing(cell.Value) = True Then
                                cell.Value = ""
                            End If
                        Next
                        If .Cells(1).Value = "" Then
                            MsgBox("Wrong contact person's name, please fill in contact person's name data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvCP.ClearSelection()
                            Me.TabControl1.SelectedIndex = 1
                            .Selected = True
                            Return False
                        End If
                    End If
                End With
            Next

            'Comment: Perubahan NPWP 16 Digit
            If Me.TxtNPWPLama.Text = "" Or Me.TxtNPWPLama.Text = "-" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWP.Text
            End If
            If Today.Date >= Me.ObjSettingNPWP.Tglberlaku Then
                '2025 Core Tax Tahap 2.1
                If Me.ObjSettingNPWP.ApakahDataSudahBenar_MasterSupplier_new(Me.TxtNPWP.Text, Me.TxtNPWPLama.Text, _
                   Me.CboTipeNPWP, True) = False Then
                    Return False
                End If
            End If

            'Update PPN 12%
            If ObjCD.CekNumerik(Me.TxtTKUID.Text) = False Then
                MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                Me.TxtTKUID.Focus()
                Return False
            End If
            Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
            If Me.TxtEmail.Text <> "" Then
                Me.TxtEmail.Text = Me.TxtEmail.Text.Trim
                If Expression.IsMatch(Me.TxtEmail.Text) = False Then
                    MsgBox("The email address is not valid. Please check again.", MsgBoxStyle.Information, "Information")
                    Me.TxtEmail.Focus()
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.objSupplier
                .CompanyWebsite = Me.TxtCompanyWebsite.Text
                .country = Me.TxtCountry.Text
                .ForeignSupplier = Me.rboYes.Checked
                .LeadTime = Me.txtLeadTime.Text
                .NPWP = Me.TxtNPWP.Text
                .PaymentTerm = Me.TxtPaymentTerm.Text
                .PriceType = Me.CboPriceType.SelectedIndex
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .SupplierCode = Me.txtsuppliercode.Text
                .SupplierName = Me.txtsuppliername.Text
                'Comment: NPWP 16 Digit
                .NPWPLama = Me.TxtNPWPLama.Text
                .NPWPType = Me.CboTipeNPWP.SelectedIndex
                'Update PPN 12%
                .TKU_ID = Me.TxtTKUID.Text
                .Email = Me.TxtEmail.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objSupplier
                Me.TxtCompanyWebsite.Text = .CompanyWebsite
                Me.TxtCountry.Text = .country
                If .ForeignSupplier = True Then
                    Me.rboYes.Checked = True
                Else
                    Me.rboNo.Checked = True
                End If
                Me.txtLeadTime.Text = .LeadTime
                Me.TxtNPWP.Text = .NPWP
                Me.TxtPaymentTerm.Text = .PaymentTerm
                Me.CboPriceType.SelectedIndex = .PriceType
                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency
                Me.TxtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                Me.txtsuppliercode.Text = .SupplierCode
                Me.txtsuppliername.Text = .SupplierName
                Me.txtLeadTime.Text = Me.ObjCekData.DigitGrouping(Me.txtLeadTime.Text)
                Me.txtprimarysupplier.Text = Me.objSupplier.Primary
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                'Comment: NPWP No 16 Digit
                Me.TxtNPWPLama.Text = .NPWPLama
                Me.CboTipeNPWP.SelectedIndex = .NPWPType

                If IsNumeric(Me.TxtPaymentTerm.Text) = False Then
                    Me.TxtPaymentTerm.Text = "0"
                Else
                    Me.ObjCD.DigitGrouping(Me.TxtPaymentTerm.Text)
                End If
                'Update PPN 12%
                Me.TxtTKUID.Text = .TKU_ID
                Me.TxtEmail.Text = .Email
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSupplier_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.ObjCekData = Nothing
            Me.objSupplier = Nothing
            Me.ObjCurrency = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvAddress.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvBank.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.TxtPaymentTerm.Text = "0"

            '@bankbranch_swiftcode
            Me.dgvBank.Columns(Me.ColSwiftCode.Index).DisplayIndex = 3
            Me.dgvBank.Columns(Me.ColBankBranch.Index).DisplayIndex = 4
            Me.TxtPaymentTerm.MaxLength = 5
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objSupplier.SimpanBaru(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objSupplier.SimpanPerubahan(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank) Then
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
            If Not Me.txtsuppliercode.Text = "" And Me.txtsuppliername.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objSupplier.CariDetail(Me.objSupplier.Primary, Me.dgvAddress, _
                        Me.dgvCP, Me.dgvBank)
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.txtsuppliercode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariMasterSupplier
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimarySupplier = -1 Then
                If Me.objSupplier.CariDetail(hj.PrimarySupplier, Me.dgvAddress, _
                Me.dgvCP, Me.dgvBank) = True Then
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
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.txtsuppliercode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtsuppliercode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.txtsuppliercode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtsuppliercode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtsuppliercode.Text <> "" Then
                If Me.objSupplier.ApakahBolehHapus(Me.txtprimarysupplier.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objSupplier.Hapus(Me.txtprimarysupplier.Text) = True Then
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

    Private Sub FrmMasterSupplierGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtsuppliercode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSupplier.SimpanBaru(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank) = True Then
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
                            If Me.objSupplier.SimpanPerubahan(Me.UserCode, Me.dgvCP, Me.dgvAddress, Me.dgvBank) = True Then
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

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        Dim Obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                    Me.TxtPrimaryCurrency.Text = Obj.GetPrimary(Me.TxtCurrencyCode.Text)
                    Me.rboYes.Focus()
                Else
                    If Me.TxtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode.Text = Obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                        Me.rboYes.Focus()
                    Else
                        Me.TxtCurrencyCode.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtCurrencyCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCurrencyCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                    Me.rboYes.Focus()
                    Me.TxtPrimaryCurrency.Text = obj.GetPrimary(Me.TxtCurrencyCode.Text)
                Else
                    If Me.TxtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode.Text = obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                        Me.rboYes.Focus()
                    Else
                        Me.TxtCurrencyCode.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtCurrencyCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If obj.ApakahSudahAda(Me.TxtCurrencyCode.Text) = True Then
                Me.rboYes.Focus()
                Me.TxtPrimaryCurrency.Text = obj.GetPrimary(Me.TxtCurrencyCode.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.TxtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.TxtCurrencyCode.Text = obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    Me.rboYes.Focus()
                Else
                    Me.TxtCurrencyCode.Clear()
                    Me.TxtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub dgvAddress_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAddress.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAddress.CellEndEdit
        Try
            If IsNothing(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '2024-03-25 Ilham, dicomment dan ditambah karena address tidak bisa kosong
                'Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                If e.ColumnIndex = 1 Then
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            Else
                If e.ColumnIndex = 1 Then
                    'untuk address code tidak boleh kosong
                    If Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Address code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'harus huruf kapital semua
                    Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                    'kode alamat tidak ada yang boleh double
                    Dim row As DataGridViewRow
                    For Each row In Me.dgvAddress.Rows
                        If row.Cells(1).Value = Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value And row.Index <> e.RowIndex Then
                            MsgBox("Duplication in address code, please change your address code.", MsgBoxStyle.Information, "Information")
                            Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    Next
                Else
                    If Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 40 Then
                        Me.dgvAddress.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAddress.Item(e.ColumnIndex, _
                        e.RowIndex).Value.ToString.Substring(0, 40)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCP.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCP.CellEndEdit
        Try
            If IsNothing(Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            Else
                If e.ColumnIndex = 1 Then
                    'untuk ItemCP nama tidak boleh kosong
                    If Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Contact person name cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvCP.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemCP_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserAddedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemBank_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvBank.CellBeginEdit
        Try
            If e.ColumnIndex = 0 Then
                e.Cancel = True
            Else
                Me.buffer = Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan nilai true jika ada 2 baris dalam dgv bank yang double. Dianggap sama jika
    ''' Nama bank, bank account no sama
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahAdaDoubleBank() As Boolean
        Try
            If Me.dgvBank.RowCount > 2 Then
                For i As Integer = 0 To Me.dgvBank.RowCount - 3
                    For j As Integer = i + 1 To Me.dgvBank.RowCount - 2
                        If Me.dgvBank.Item(1, i).Value = Me.dgvBank.Item(1, j).Value And Me.dgvBank.Item(2, i).Value = Me.dgvBank.Item(2, j).Value Then
                            Return True
                        End If
                    Next
                Next
            End If

            'kalo sampe sini maka tidak ada yang sama
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemBank_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBank.CellEndEdit
        Try
            If IsNothing(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            Else
                If e.ColumnIndex = 1 Then
                    'untuk ItemBank nama tidak boleh kosong
                    If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Bank's name cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'cek tidak terima ada data double
                    If Me.ApakahAdaDoubleBank = True Then
                        MsgBox("You cannot have 2 bank's data with the same bank's name and account no", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                ElseIf e.ColumnIndex = 2 Then
                    'untuk account number juga tidak boleh kosong
                    If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        MsgBox("Account number cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    'cek tidak terima ada data double
                    If Me.ApakahAdaDoubleBank = True Then
                        MsgBox("You cannot have 2 bank's data with the same bank's name and account no", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                ElseIf e.ColumnIndex = 3 Then   'untuk currency harus yang ada didalam data master currency
                    Dim Obj As New clsMasterCurrency
                    If Obj.ApakahSudahAda(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Dim i As New FrmCariMasterCurrency
                        If Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                            i.txtCari.Text = Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency <> -1 Then
                            Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.ObjCurrency.GetCurrencyCode(i.PrimaryCurrency)).ToUpper
                            Me.dgvBank.Item(8, e.RowIndex).Value = i.PrimaryCurrency
                        Else
                            Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        i = Nothing
                    Else
                        Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                        Me.dgvBank.Item(8, e.RowIndex).Value = Me.ObjCurrency.GetPrimary(Me.dgvBank.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                    Obj = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvBank.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvBank)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvBank.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvBank.Rows.Count - 1 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBank_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvBank.UserAddedRow
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvBank.Rows(Me.dgvBank.Rows.Count - 1)
                .Cells(1).Value = "-"
                .Cells(3).Value = "IDR"
                .Cells(Me.ColPrimaryCurrency.Index).Value = Me.ObjCurrency.GetPrimary("IDR")
            End With
            Me.ObjCekData.SetNumber(Me.dgvBank)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvCP.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvCP.Rows.Count - 1 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCP_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCP.UserAddedRow
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvCP.Rows(Me.dgvCP.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvCP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAddress.UserDeletedRow
        Try
            Me.ObjCekData.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvAddress.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvAddress.Rows.Count - 1 Then
                e.Cancel = True
            Else
                'jika sudah pernah digunakan maka tidak bisa dihapus.
                If Me.baru = False Then
                    If Me.objSupplier.ApakahBolehHapusAddress(Me.dgvAddress.SelectedRows(0).Cells(Me.ColIdentifierAdd.Index).Value) = False Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAddress_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAddress.UserAddedRow
        Try
            'tentukan nilai default untuk baris terakhir
            With Me.dgvAddress.Rows(Me.dgvAddress.Rows.Count - 1)
                .Cells(1).Value = "-"
            End With
            Me.ObjCekData.SetNumber(Me.dgvAddress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtsuppliercode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsuppliercode.GotFocus
        txtsuppliercode.BackColor = BackColorAll
    End Sub

    Private Sub txtsuppliercode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsuppliercode.LostFocus
        txtsuppliercode.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub txtsuppliername_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsuppliername.GotFocus
        txtsuppliername.BackColor = BackColorAll
    End Sub

    Private Sub txtsuppliername_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsuppliername.LostFocus
        txtsuppliername.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub TxtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.GotFocus
        TxtPaymentTerm.BackColor = BackColorAll
    End Sub

    Private Sub TxtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        TxtPaymentTerm.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub txtLeadTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadTime.GotFocus
        txtLeadTime.BackColor = BackColorAll
    End Sub

    Private Sub txtLeadTime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeadTime.KeyDown
        Me.ObjCekData.Number_KeyDown(Me.txtLeadTime, sender, e)
    End Sub

    Private Sub txtLeadTime_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeadTime.KeyUp
        Me.ObjCekData.Number_KeyUp(Me.txtLeadTime, sender, e)
    End Sub

    Private Sub txtLeadTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadTime.LostFocus
        txtLeadTime.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub TxtCountry_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.GotFocus
        TxtCountry.BackColor = BackColorAll
    End Sub

    Private Sub TxtCountry_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.LostFocus
        TxtCountry.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub txtDefaultCurrency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.GotFocus
        TxtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtDefaultCurrency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode.LostFocus
        TxtCurrencyCode.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub TxtNPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP.GotFocus
        TxtNPWP.BackColor = BackColorAll
    End Sub

    Private Sub TxtNPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP.LostFocus
        TxtNPWP.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub txtCompanywebSide_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyWebsite.GotFocus
        TxtCompanyWebsite.BackColor = BackColorAll
    End Sub

    Private Sub txtCompanywebSide_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyWebsite.LostFocus
        TxtCompanyWebsite.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub TxtTKUID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.GotFocus
        TxtTKUID.BackColor = BackColorAll
    End Sub

    Private Sub TxtTKUID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.LostFocus
        TxtTKUID.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub TxtEmail_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.GotFocus
        TxtEmail.BackColor = BackColorAll
    End Sub

    Private Sub TxtEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.LostFocus
        TxtEmail.BackColor = Me.txtprimarysupplier.BackColor
    End Sub

    Private Sub dgvAddress_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAddress.CellContentClick

    End Sub

    Private Sub TxtPaymentTerm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.TextChanged

    End Sub

    Private Sub gboDetail_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gboDetail.Enter

    End Sub

    Private Sub LlbEditNPWPLama_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbEditNPWPLama.Click
        'Comment: Perubahan NPWP 16 Digit
        Dim hj As New FrmMaster_EditNPWPLama
        Try
            If Me.TxtNPWPLama.Text = "" Then
                Me.TxtNPWPLama.Text = Me.TxtNPWP.Text
            End If

            hj.TxtNPWPNo.Text = Me.TxtNPWPLama.Text
            If hj.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TxtNPWPLama.Text = hj.TxtNPWPNo.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
End Class