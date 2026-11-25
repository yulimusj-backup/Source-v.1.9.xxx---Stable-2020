Option Explicit On
Option Strict Off

Public Class FrmMasterUser_TempPass
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim Buffer As String = ""
    Dim Baru As Boolean
    Dim ObjTempPass As New ClsUser_TempPass
    Dim ObjCD As New ClsCekData
    Dim ObjUser As New ClsUser
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim TargetIndex As Integer = -1
    Dim CheckTempPass As Boolean = False

    Private Sub FrmTransaksiJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal
        Dim NullValue As String = ""

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.Baru = True
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.Rows.Clear()
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'set null value untuk dgv
            Me.ColEndTime.DefaultCellStyle.NullValue = "17.00"
            Me.ColStartTime.DefaultCellStyle.NullValue = "08.00"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True And Me.dgvItem.AllowUserToAddRows = True Then
                'set nilai default baris baru
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColConfirmPassword.Index).Value = ""
                rowDGV.Cells(Me.ColEndDate.Index).Value = ""
                rowDGV.Cells(Me.ColEndTime.Index).Value = "17.00"
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                rowDGV.Cells(Me.ColPass.Index).Value = ""
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                rowDGV.Cells(Me.ColStartDate.Index).Value = ""
                rowDGV.Cells(Me.ColStartTime.Index).Value = "08.00"
                rowDGV.Cells(Me.ColConfirmPassword_Show.Index).Value = ""
                rowDGV.Cells(Me.ColPass_Show.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.RowCount > 0 Then
                Me.ObjCD.SetNumber(Me.dgvItem)
            End If

            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.AllowUserToAddRows = True Then
                    If Me.dgvItem.RowCount = 1 Then
                        Me.TxtCode.Enabled = True
                        Me.btnFindUserCode.Enabled = True
                    Else
                        Me.TxtCode.Enabled = False
                        Me.btnFindUserCode.Enabled = False
                    End If
                Else
                    If Me.dgvItem.RowCount = 0 Then
                        Me.TxtCode.Enabled = True
                        Me.btnFindUserCode.Enabled = True
                    Else
                        Me.TxtCode.Enabled = False
                        Me.btnFindUserCode.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmMasterUser_TempPass_Cari

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.Primary = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjTempPass.CariDetail(ij.Primary, Me.dgvItem, Me) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    ''fungsi standar 
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToAddRows = True

            If Me.dgvItem.RowCount <= 1 Then
                Me.TxtCode.Enabled = True
                Me.btnFindUserCode.Enabled = True
            Else
                Me.TxtCode.Enabled = False
                Me.btnFindUserCode.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.CheckTempPass = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjTempPass
                .PrimaryUser = Me.TxtPrimaryUser.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Try
            With ObjTempPass
                Me.TxtPrimary.Text = .Primary
                Me.TxtPrimaryUser.Text = .PrimaryUser
                Me.TxtCode.Text = Me.ObjUser.GetUserCode(.PrimaryUser)
                Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim row As DataGridViewRow
        Dim LedgerAccess_YgMengubah, LedgerAccess_YgDiUbah As ClsUser.EnumJournalAccesType

        Try
            'user yang memiliki hak akses lebih rendah pada ledger access, tidak boleh melakuka reset terhadap pass dari user memiliki ledger access lebih tinggi
            If Me.UserCode <> Me.TxtCode.Text Then
                LedgerAccess_YgMengubah = Me.ObjUser.GetLedgerAccess(Me.UserCode)
                LedgerAccess_YgDiUbah = Me.ObjUser.GetLedgerAccess(Me.TxtCode.Text)

                If LedgerAccess_YgDiUbah = ClsUser.EnumJournalAccesType.FullAccess Then
                    If LedgerAccess_YgMengubah <> ClsUser.EnumJournalAccesType.FullAccess Then
                        MsgBox("User " + Me.TxtCode.Text + _
                        " has 'full view' ledger access, only users with the same ledger access authorization that can provide tempory password for him/her.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                ElseIf LedgerAccess_YgDiUbah = ClsUser.EnumJournalAccesType.Selected_Users_Only Then
                    If LedgerAccess_YgMengubah <> ClsUser.EnumJournalAccesType.FullAccess Then
                        MsgBox("User " + Me.TxtCode.Text + _
                        " has 'selected user view' ledger access, only users with the 'full view' ledger access authorization that can provide tempory password for him/her.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("User Code")
            KoleksiTextBox.Add(Me.TxtCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'Cek DGVItem
            If Me.dgvItem.Rows.Count < 1 Then
                MsgBox("Please input at least 1 item.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Focus()
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.ObjTempPass.ApakahSudahAda(Me.TxtPrimaryUser.Text) = True Then
                    MsgBox("User code duplication, please change your user code", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.ObjTempPass.GetJmlTranskasiLainByPrimaryUser(Me.TxtPrimaryUser.Text, Me.TxtPrimary.Text) > 0 Then
                    MsgBox("User code duplication, please change your user code", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjTempPass.GetTimeStamp(Me.TxtPrimary.Text)
                If Me.ObjTempPass.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjTempPass.CariDetail(Me.ObjTempPass.Primary, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                    Return False
                End If
            End If

            'inconsistent password untuk baris terakhir
            row = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            If row.Cells(Me.ColConfirmPassword.Index).Value <> row.Cells(Me.ColPass.Index).Value Then
                MsgBox("Inconsistent password.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.ClearSelection()
                Me.dgvItem.Rows(Me.dgvItem.RowCount - 1).Selected = True
                Return False
            End If
            'password tidak boleh kosong
            If row.Cells(Me.ColPass.Index).Value = "" Then
                MsgBox("Please specify a valid password.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.ClearSelection()
                Me.dgvItem.Rows(Me.dgvItem.RowCount - 1).Selected = True
                Return False
            End If

            If Me.TxtPrimaryUser.Text <> "" And Me.TxtPrimaryUser.Text <> "-1" Then
                If Me.ObjUser.ApakahBolehMemilikiTempPass(Me.TxtPrimaryUser.Text) = False Then
                    MsgBox("This user can't have temporary password.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
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

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Enter
        Try
            Me.Buffer = Me.TxtRemarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''back color dan tampilan 
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

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim PrimaryDummy As Double = -1
        Dim PrimaryAkunRE As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dgvItem.AllowUserToAddRows = True
            Me.ObjTempPass.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.CheckTempPass = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim FrmO As New FrmMasterUser_Otorisasi
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.ObjTempPass.ApakahBolehEdit(Me.ObjTempPass.Primary, True, UserCode, True) = True Then
                    Me.Baru = False
                    Me.ObjTempPass.CariDetail(Me.ObjTempPass.Primary, Me.dgvItem, Me)
                    Me.SalinkeField()
                    Me.AktifkanDetail()
                    Me.TxtCode.Focus()
                End If
            End If

            Me.CheckTempPass = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.ObjTempPass.ApakahBolehHapus(Me.TxtPrimary.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjTempPass.Hapus(Me.TxtPrimary.Text, UserCode, True, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If Me.ApakahDataBenar() = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjTempPass.SimpanBaru(Me.UserCode, Me.dgvItem, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = True
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjTempPass.ApakahBolehEdit(Me.TxtPrimary.Text, True, UserCode, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjTempPass.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                            Me.dgvItem.AllowUserToAddRows = False
                        ElseIf hasil = MsgBoxResult.No Then
                            Me.dgvItem.AllowUserToAddRows = True
                        End If
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRemarks.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtCode.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjTempPass.CariDetail(Me.TxtPrimary.Text, Me.dgvItem, Me)
                        Me.SalinkeField()
                    End If
                Else
                    Me.dgvItem.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.dgvItem.AllowUserToAddRows = False
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjTempPass.SimpanBaru(Me.UserCode, Me.dgvItem, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                Me.dgvItem.AllowUserToAddRows = True
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjTempPass.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                Me.dgvItem.AllowUserToAddRows = True
                                e.Cancel = True
                            End If
                        End If
                    Else
                        Me.dgvItem.AllowUserToAddRows = True
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim row As DataGridViewRow

        Try
            If e.ColumnIndex <> Me.ColStartDate.Index Then
                If e.RowIndex <> Me.dgvItem.RowCount - 2 Then
                    'utnuk kolom selain start date, yang bisa diedit hanya baris terakhir data saja
                    e.Cancel = True
                Else
                    If e.ColumnIndex = Me.ColEndDate.Index Or _
                    e.ColumnIndex = Me.ColEndTime.Index Or _
                    e.ColumnIndex = Me.ColRemarks.Index Or _
                    e.ColumnIndex = Me.ColStartTime.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    ElseIf e.ColumnIndex = Me.ColConfirmPassword_Show.Index Then
                        If Me.dgvItem.Item(Me.ColPass.Index, e.RowIndex).Value = "" Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColPass_Show.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                'col start date
                If e.ColumnIndex = Me.ColStartDate.Index Then
                    'menambahkan baris baru atau mengedit data baris terakhir
                    If e.RowIndex = Me.dgvItem.RowCount - 2 Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    ElseIf e.RowIndex = Me.dgvItem.RowCount - 1 Then
                        If Me.TxtPrimaryUser.Text = "-1" Or Me.TxtPrimaryUser.Text = "" Then
                            MsgBox("Please specify user first.", MsgBoxStyle.Information, "Information")
                            Me.TxtCode.Focus()
                            e.Cancel = True
                        ElseIf e.RowIndex - 1 >= 0 Then
                            'inconsistent password untuk baris sebelumnya
                            row = Me.dgvItem.Rows(e.RowIndex - 1)
                            If row.Cells(Me.ColConfirmPassword.Index).Value <> row.Cells(Me.ColPass.Index).Value Then
                                MsgBox("Inconsistent password.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row.Selected = True
                                e.Cancel = True
                                Exit Sub
                            End If
                            'password tidak boleh kosong utnuk baris sebelumnya
                            If row.Cells(Me.ColPass.Index).Value = "" Then
                                MsgBox("Please specify a valid password.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.ClearSelection()
                                row.Selected = True
                                e.Cancel = True
                                Exit Sub
                            End If
                            'sampe sini catatkan nilai buffernya
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Dummy As String = ""
        Dim rowdgv As DataGridViewRow

        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                If e.ColumnIndex = Me.ColConfirmPassword_Show.Index Then
                    'trim password
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    'tidak boleh string kosong
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'maksimal 16 karakter
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 16 Then
                        MsgBox("Please use maximum 16 character string for password.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'langsung cocokan dengan kolom password
                    Dummy = Me.ObjUser.EnkripsiPassword(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    If Dummy <> Me.dgvItem.Item(Me.ColPass.Index, e.RowIndex).Value Then
                        MsgBox("Inconsistent password.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'copykan ke dalam kolom sesunggunhnya
                    Me.dgvItem.Item(Me.ColConfirmPassword.Index, e.RowIndex).Value = Dummy
                    'kembalikan ke ****
                    Me.dgvItem.Item(Me.ColConfirmPassword_Show.Index, e.RowIndex).Value = "*****"
                ElseIf e.ColumnIndex = Me.ColPass_Show.Index Then
                    'trim password
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    'tidak boleh string kosong
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'maksimal 16 karakter
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 16 Then
                        MsgBox("Please use maximum 16 character string for password.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'enkripsi
                    Dummy = Me.ObjUser.EnkripsiPassword(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    'copykan ke dalam kolom sesunggunhnya
                    Me.dgvItem.Item(Me.ColPass.Index, e.RowIndex).Value = Dummy
                    'kembalikan ke ****
                    Me.dgvItem.Item(Me.ColPass_Show.Index, e.RowIndex).Value = "*****"
                    'kosongkan kolom confirm password dan confirm password show
                    Me.dgvItem.Item(Me.ColConfirmPassword.Index, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColConfirmPassword_Show.Index, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColRemarks.Index Then
                    'maksimal 250 karakter
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                ElseIf e.ColumnIndex = Me.ColStartDate.Index Then
                    'pastikan tanggal
                    If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        If Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value = "" Then
                            'delete row
                            Me.TargetIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    'format tanggal
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                    'minimal end date baris sebelumnya +1
                    If e.RowIndex - 1 >= 0 Then
                        rowdgv = Me.dgvItem.Rows(e.RowIndex - 1)
                        If CDate(rowdgv.Cells(Me.ColEndDate.Index).Value).Date >= CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date Then
                            MsgBox("Minimum date is: " + Format(CDate(rowdgv.Cells(Me.ColEndDate.Index).Value) + TimeSpan.FromDays(1), "dd MMM yyyy"), _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value = "" Then
                                'delete row
                                Me.TargetIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    End If
                    'maksimal sama dengan end date
                    If IsDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value) = True Then
                        If CDate(Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value).Date > _
                        CDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value).Date Then
                            MsgBox("Maximum date is: " + Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value.ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'maksimum 7 hari
                        If (CDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value).Date - _
                        CDate(Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value).Date).TotalDays > 6 Then
                            MsgBox("Maximum time span for temporary password is 7 days.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    'tidak boleh beririsan dengan baris lainny
                    For i As Integer = e.RowIndex - 1 To 0 Step -1
                        rowdgv = Me.dgvItem.Rows(i)
                        If CDate(Me.dgvItem.Item(Me.ColStartDate.Index, i).Value) <= CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) _
                        And CDate(Me.dgvItem.Item(Me.ColEndDate.Index, i).Value) >= CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Date duplication please change date.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value = "" Then
                                'delete row
                                Me.TargetIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    Next
                    'khusus untuk kondisi dimana start date = end date baris sebelumnya + 1
                    'dan pada baris sebelumnya end time < start time, maka start time pada baris selanjut harus > dari baris sebelumnya
                    If e.RowIndex - 1 >= 0 Then
                        rowdgv = Me.dgvItem.Rows(e.RowIndex - 1)
                        If (CDate(rowdgv.Cells(Me.ColEndDate.Index).Value) + TimeSpan.FromDays(1)).Date = _
                        CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date Then
                            If Me.ObjCD.Convert_ToMinutes(rowdgv.Cells(Me.ColEndTime.Index).Value) < _
                            Me.ObjCD.Convert_ToMinutes(rowdgv.Cells(Me.ColStartTime.Index).Value) Then
                                If Me.ObjCD.Convert_ToMinutes(rowdgv.Cells(Me.ColEndTime.Index).Value) >= _
                                Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(Me.ColStartTime.Index, e.RowIndex).Value) Then
                                    MsgBox("Minimum start time is: " + Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(rowdgv.Cells(Me.ColEndTime.Index).Value) + 1), _
                                    MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColStartTime.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(rowdgv.Cells(Me.ColEndTime.Index).Value) + 1)
                                    Me.dgvItem.Item(Me.ColEndTime.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(Me.ColStartTime.Index, e.RowIndex).Value) + _
                                    (8 * 60))
                                    If Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEndTime.Index, e.RowIndex).Value) > Me.ObjCD.Convert_ToMinutes("23.59") Then
                                        Me.dgvItem.Item(Me.ColEndTime.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(Me.ColEndTime.Index, e.RowIndex).Value) - Me.ObjCD.Convert_ToMinutes("24.00"))
                                    End If
                                End If
                            End If
                        End If
                    End If

                    'copy nilainya ke end date
                    If IsDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Me.TxtCode.Enabled = False
                        Me.btnFindUserCode.Enabled = False
                        Me.btnFindUserCode.TabStop = False
                    End If
                ElseIf e.ColumnIndex = Me.ColEndDate.Index Then
                    'pastikan tanggal
                    If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'format tanggal
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                    'minimal sama dengan start date
                    If CDate(Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value).Date > _
                    CDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value).Date Then
                        MsgBox("Minimum date is: " + Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'tidak boleh beririsan dengan tanggal lainny
                    For i As Integer = e.RowIndex - 1 To 0 Step -1
                        rowdgv = Me.dgvItem.Rows(i)
                        If CDate(Me.dgvItem.Item(Me.ColStartDate.Index, i).Value) <= CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) _
                        And CDate(Me.dgvItem.Item(Me.ColEndDate.Index, i).Value) >= CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Date duplication please change date.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    Next
                    'maksimum 7 hari
                    If (CDate(Me.dgvItem.Item(Me.ColEndDate.Index, e.RowIndex).Value).Date - _
                    CDate(Me.dgvItem.Item(Me.ColStartDate.Index, e.RowIndex).Value).Date).TotalDays > 6 Then
                        MsgBox("Maximum time span for temporary password is 7 days.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColEndTime.Index Or e.ColumnIndex = Me.ColStartTime.Index Then
                    'pastikan time string
                    If Me.ObjCD.Is_Time_String(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify valid time data.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'nilai max 23.59
                    If Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > Me.ObjCD.Convert_ToMinutes("23.59") Then
                        MsgBox("Maximum value is 23.59.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'format ulang
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Me.ObjCD.Convert_ToMinutes(Me.dgvItem.Item(e.ColumnIndex, _
                    e.RowIndex).Value))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                TxtCode.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                TxtCode.BackColor = Me.TxtPrimary.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCode.KeyDown
        Dim i As New FrmCariMasterUser
        Try
            i.BackColorAll = i.BackColorAll
            If e.KeyValue = 13 Then
                If ObjUser.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    Me.TxtPrimaryUser.Text = ObjUser.GetPrimary(Me.TxtCode.Text)
                    Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                    Me.TxtRemarks.Focus()
                Else
                    If Me.TxtCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtPrimaryUser.Text = Me.ObjUser.GetPrimary(i.Kodeuser)
                        Me.TxtCode.Text = i.Kodeuser
                        Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                        Me.TxtRemarks.Focus()
                    Else
                        Me.TxtCode.Clear()
                        Me.TxtPrimaryUser.Text = "-1"
                        Me.TxtDesc.Clear()
                        Me.TxtCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.Leave
        Dim i As New FrmCariMasterUser
        Dim obj As New ClsMasterProductGroup

        Try
            If Me.TxtRemarks.Enabled = True Then
                i.backColorAll = Me.BackColorAll
                If Me.TxtCode.Text <> "" Then
                    If ObjUser.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        Me.TxtPrimaryUser.Text = ObjUser.GetPrimary(Me.TxtCode.Text)
                        Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                    Else
                        If Me.TxtCode.Text <> "" Then
                            i.TxtCari.Text = Me.TxtCode.Text
                        End If

                        i.ShowDialog()
                        If i.Kodeuser <> "" Then
                            Me.TxtPrimaryUser.Text = Me.ObjUser.GetPrimary(i.Kodeuser)
                            Me.TxtCode.Text = i.Kodeuser
                            Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                        Else
                            Me.TxtCode.Clear()
                            Me.TxtPrimaryUser.Text = "-1"
                            Me.TxtCode.Focus()
                            Me.TxtDesc.Clear()
                        End If
                    End If

                    If Me.TxtPrimaryUser.Text <> "" And Me.TxtPrimaryUser.Text <> "-1" Then
                        If Me.CheckTempPass = True Then
                            If Me.ObjUser.ApakahBolehMemilikiTempPass(Me.TxtPrimaryUser.Text) = False Then
                                MsgBox("This user can't have temporary password.", MsgBoxStyle.Information, "Information")
                                Me.TxtCode.Clear()
                                Me.TxtPrimaryUser.Text = "-1"
                                Me.TxtCode.Focus()
                                Me.TxtDesc.Clear()
                            End If
                        End If
                    End If
                Else
                    Me.TxtPrimaryUser.Text = "-1"
                    Me.TxtCode.Clear()
                    Me.TxtDesc.Clear()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindUserCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUserCode.Click
        Dim i As New FrmCariMasterUser
        Dim obj As New ClsMasterProductGroup

        Try
            i.backColorAll = Me.BackColorAll
            If Me.TxtCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtCode.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtPrimaryUser.Text = Me.ObjUser.GetPrimary(i.Kodeuser)
                Me.TxtCode.Text = i.Kodeuser
                Me.TxtDesc.Text = Me.ObjUser.GetNama(Me.TxtCode.Text)
                Me.TxtRemarks.Focus()
            Else
                Me.TxtCode.Clear()
                Me.TxtDesc.Clear()
                Me.TxtCode.Focus()
            End If

            If Me.TxtPrimaryUser.Text <> "" And Me.TxtPrimaryUser.Text <> "-1" Then
                If Me.ObjUser.ApakahBolehMemilikiTempPass(Me.TxtPrimaryUser.Text) = False Then
                    MsgBox("This user can't have temporary password.", MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Clear()
                    Me.TxtPrimaryUser.Text = "-1"
                    Me.TxtCode.Focus()
                    Me.TxtDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            If e.Row.Index <> Me.dgvItem.RowCount - 2 Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TargetIndex <> -1 Then
                Me.dgvItem.Rows.RemoveAt(Me.TargetIndex)
            End If
            Me.TargetIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCode.TextChanged

    End Sub
End Class