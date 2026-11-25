Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPL_Setting
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objDepartement As New Core_Module.ClsMasterDepartement
    Dim objPosition As New Core_Module.ClsMasterJobPosition
    Dim objEmp As New Core_Module.ClsMasterKaryawan
    Dim objUser As New Core_Module.ClsUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New ClsTrans_SPL_Setting
    Dim ObjLicense As New ClsLicense
    Dim ObjAntrian As New ClsAntrian

    Dim UserLain, NamaKomputerLain As String
    Dim mClose As Boolean = False
    Dim _ok As Boolean = False
    Dim e As System.Windows.Forms.DataGridViewCellEventArgs
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0

    Enum EnumAcuanItem As Integer
        ForDepartment
        ForPosition
        ForEmployee
    End Enum

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            If Me.chkApplyToAllEmployee.Checked = False Then
                Me.dgvScope.AllowUserToAddRows = True
            End If
            Me.DgvUser.AllowUserToAddRows = True
            Me.chkApplyToAllEmployee.Enabled = True
            Me.dgvScope.ReadOnly = False
            Me.DgvUser.ReadOnly = False

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvScope.TabStop = True
            Me.DgvUser.TabStop = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.dgvScope.AllowUserToAddRows = False
            Me.DgvUser.AllowUserToAddRows = False
            Me.chkApplyToAllEmployee.Enabled = False
            Me.dgvScope.ReadOnly = True
            Me.DgvUser.ReadOnly = True

            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, False)
            Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, False)

            Me.txtPrimary.Text = "-1"
            Me.txtCode.Text = ""
            Me.TxtDesc.Text = ""
            Me.txtRemarks.Text = ""
            Me.chkApplyToAllEmployee.Checked = False

            Me.dgvScope.Rows.Clear()
            Me.DgvUser.Rows.Clear()

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
            'tidak boleh kosong
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.txtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.txtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.txtCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.Obj.GetCode(Me.txtPrimary.Text)
                If CodeLama <> Me.txtCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.txtCode.Text) = True Then
                        MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                        Me.txtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.Obj.GetTimeStamp(Me.txtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvScope, Me.DgvUser)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'harus ada minimal 1 item jika chk all employee di unchecklist
            If Me.chkApplyToAllEmployee.Checked = False Then
                If Me.dgvScope.RowCount - 1 < 1 Then
                    Me.TabControl1.SelectedIndex = 0
                    MsgBox("Please input item scope", MsgBoxStyle.Information, "Information")
                    Me.dgvScope.ClearSelection()
                    Me.dgvScope.Rows(0).Selected = True
                    Return False
                End If
            End If

            If Me.DgvUser.RowCount - 1 < 1 Then
                Me.TabControl1.SelectedIndex = 1
                MsgBox("Please input item user", MsgBoxStyle.Information, "Information")
                Me.DgvUser.Rows(0).Selected = True
                Return False
            Else
                Dim row As DataGridViewRow
                For i As Integer = 0 To Me.DgvUser.RowCount - 1
                    row = Me.DgvUser.Rows(i)
                    If String.IsNullOrEmpty(row.Cells(Me.ColUserCode.Index).Value) = False Then
                        If String.IsNullOrEmpty(row.Cells(Me.ColUserEmail.Index).Value) = True Then
                            MsgBox("Please input email address.", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvUser.ClearSelection()
                            row.Cells(Me.ColUserEmail.Index).Selected = True
                            Return False
                        End If
                        If String.IsNullOrEmpty(row.Cells(Me.ColUserApproved.Index).Value) = True Then
                            MsgBox("Please input approved by.", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvUser.ClearSelection()
                            row.Cells(Me.ColUserApproved.Index).Selected = True
                            Return False
                        End If
                        If String.IsNullOrEmpty(row.Cells(Me.ColUserApprovedPosition.Index).Value) = True Then
                            MsgBox("Please input approved by position.", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvUser.ClearSelection()
                            row.Cells(Me.ColUserApprovedPosition.Index).Selected = True
                            Return False
                        End If
                    End If
                Next
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .Code = Me.txtCode.Text
                .Description = Me.TxtDesc.Text
                .Remarks = Me.txtRemarks.Text
                If Me.chkApplyToAllEmployee.Checked = False Then
                    .AllEmployee = ClsTrans_SPL_Setting.EnumAllEmployee.No
                Else
                    .AllEmployee = ClsTrans_SPL_Setting.EnumAllEmployee.Yes
                End If
                .Remarks = Me.txtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.txtPrimary.Text = .Primary
                Me.txtCode.Text = .Code
                Me.TxtDesc.Text = .Description
                If .AllEmployee = ClsTrans_SPL_Setting.EnumAllEmployee.No Then
                    Me.chkApplyToAllEmployee.Checked = False
                Else
                    Me.chkApplyToAllEmployee.Checked = True
                End If
                Me.txtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Setting_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'fungsi batasan hanya satu user yang dapat membuka form pada satu waktu
            If Me.mClose = True Then
                MsgBox("Can't access salary formula transaction because user code : " + Me.UserLain.ToString + _
                " at computer name-ip address : " + Me.NamaKomputerLain.ToString + " is using this transaction.", _
                MsgBoxStyle.Information, "Information")
            End If
            Me.ObjAntrian.EndAntrianTunggal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Setting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.mClose = False Then
                If _ok = False Then
                    If Me.txtRemarks.Enabled = True Then
                        Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar() = True Then
                                Me.SalinKeObjek()
                                If baru = True Then
                                    'simpan baru
                                    If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvScope, Me.DgvUser, Me) = True Then
                                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                        Me.SalinKeField()
                                        Me.NonAktifkanDetail()
                                        Me.btnFind.Focus()
                                        Me._ok = True
                                        Me.Close()
                                    Else
                                        e.Cancel = True
                                    End If
                                Else
                                    'simpan perubahan
                                    If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvScope, Me.DgvUser, Me) = True Then
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
                        ElseIf hasil = Windows.Forms.DialogResult.No Then
                            Me.BackgroundWorker_CloseForm.RunWorkerAsync()
                        ElseIf hasil = MsgBoxResult.Cancel Then
                            e.Cancel = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'fungsi batasan hanya satu user yang dapat membuka form pada satu waktu
            System.Windows.Forms.Form.CheckForIllegalCrossThreadCalls = False
            UserLain = ""
            NamaKomputerLain = ""
            If Me.ObjAntrian.StartAntrianTunggal(ClsAntrian.EnumTipeAntrian.AntrianTransaksiSPLGroup, Me.UserCode, UserLain, NamaKomputerLain) = False Then
                Me.mClose = True
                Me.BackgroundWorker_CloseForm.RunWorkerAsync()
            End If

            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvScope.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.ColEmployeeName.Visible = True
            Me.dgvScope.MultiSelect = False
            Me.DgvUser.MultiSelect = False

            Me.txtCode.Width = Me.SplitContainer1.Panel1.Width - Me.txtCode.Left
            Me.TxtDesc.Width = Me.SplitContainer1.Panel1.Width - Me.TxtDesc.Left
            Me.txtRemarks.Width = Me.SplitContainer1.Panel2.Width - Me.txtRemarks.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkApplyToAllEmployee_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkApplyToAllEmployee.CheckedChanged
        Try
            If Me.chkApplyToAllEmployee.Checked = True Then
                If Me.dgvScope.RowCount - 1 > 0 Then
                    Me.TabControl1.SelectedIndex = 0
                    If MsgBox("Checklist all employee will clear table scope", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.dgvScope.Rows.Clear()
                        Me.dgvScope.AllowUserToAddRows = False
                    Else
                        Me.chkApplyToAllEmployee.Checked = False
                    End If
                Else
                    Me.dgvScope.AllowUserToAddRows = False
                End If
            Else
                Me.dgvScope.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.dgvScope, Me.DgvUser, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.dgvScope, Me.DgvUser, Me) Then
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
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtCode.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvScope, Me.DgvUser)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If

            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_SPL_Setting_Cari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Frm = Me
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.Obj.CariDetail(hj.PrimaryTransaksi, Me, Me.dgvScope, Me.DgvUser) = True Then
                    Me.baru = False
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
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

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True

            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.txtCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        Try
            If Me.txtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtCode.Text <> "" Then
                If Me.Obj.ApakahBolehedit(Me.UserCode, Me.txtPrimary.Text) = True Then
                    Me.baru = False
                    Me.Obj.CariDetail(Me.txtPrimary.Text, Me, Me.dgvScope, Me.DgvUser)
                    Me.AktifkanDetail()
                    Me.txtCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.txtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.UserCode, Me.txtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.txtPrimary.Text, Me.UserCode, Me) = True Then
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
        Me.Close()
    End Sub

    Private Sub dgvScope_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvScope.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEmployeeCode.Index Then
                If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value) = False Or _
                String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value) = False Then
                    'beri pesan kenapa tidak bisa input employee
                    e.Cancel = True
                Else
                    Me.Buffer = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            ElseIf e.ColumnIndex = Me.ColDepartmentCode.Index Or e.ColumnIndex = Me.ColJobPositionCode.Index Then
                If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value) = True Then
                    Me.Buffer = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.colRemarks.Index Then
                    If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value) = False Or _
                                    String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value) = False Or _
                                    String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value) = False Then
                        Me.Buffer = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimaryAcuan As Double, _
    ByVal vRowIdxFocus As Integer, ByVal vEnumAcuanItem As EnumAcuanItem) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Dim IndexColumn As Integer = -1
        Dim PrimaryDept As Double = Me.dgvScope.Item(Me.ColPrimaryDept.Index, vRowIdxFocus).Value
        Dim PrimaryPosition As Double = Me.dgvScope.Item(Me.ColPrimaryPosition.Index, vRowIdxFocus).Value

        Try
            If vEnumAcuanItem = EnumAcuanItem.ForDepartment Then
                PrimaryDept = vPrimaryAcuan
            ElseIf vEnumAcuanItem = EnumAcuanItem.ForPosition Then
                PrimaryPosition = vPrimaryAcuan
            End If

            For ij As Integer = 0 To Me.dgvScope.Rows.Count - 2
                RowDGV = Me.dgvScope.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If vEnumAcuanItem = EnumAcuanItem.ForDepartment Or vEnumAcuanItem = EnumAcuanItem.ForPosition Then
                        If RowDGV.Cells(Me.ColPrimaryDept.Index).Value = PrimaryDept And RowDGV.Cells(Me.ColPrimaryPosition.Index).Value = PrimaryPosition Then
                            Retval = ij
                            Exit For
                        End If
                    Else
                        If RowDGV.Cells(Me.ColPrimaryEmp.Index).Value = vPrimaryAcuan Then
                            Retval = ij
                            Exit For
                        End If
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvScope_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvScope.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvScope.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryAcuan As Double = -1

        Try
            If IsNothing(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDepartmentCode.Index Then
                    'FUNGSI COL DEPARTMENT CODE
                    Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objDepartement.ApakahSudahAda(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcuan = Me.objDepartement.GetPrimary(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItem(PrimaryAcuan, e.RowIndex, EnumAcuanItem.ForDepartment)
                        If index < 0 Then
                            Me.dgvScope.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = PrimaryAcuan
                            Me.dgvScope.Item(Me.ColDepartmentDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(PrimaryAcuan)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication department code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterDepartementCari
                        ij.txtCari.Text = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryDepartement <> -1 Then
                            index = Me.GetIndexItem(ij.PrimaryDepartement, e.RowIndex, EnumAcuanItem.ForDepartment)
                            If index < 0 Then
                                Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value = Me.objDepartement.GetDeptCode(ij.PrimaryDepartement)
                                Me.dgvScope.Item(Me.ColDepartmentDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(ij.PrimaryDepartement)
                                Me.dgvScope.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication department code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value = Me.objDepartement.GetDeptCode(ij.PrimaryDepartement)
                                    Me.dgvScope.Item(Me.ColDepartmentDesc.Index, e.RowIndex).Value = Me.objDepartement.GetDeptDesc(ij.PrimaryDepartement)
                                    Me.dgvScope.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                End If
                            End If
                        ElseIf ij.PrimaryDepartement = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If
                ElseIf e.ColumnIndex = Me.ColJobPositionCode.Index Then
                    'FUNGSI COL JOB POSITION CODE
                    Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objPosition.ApakahSudahAda(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcuan = Me.objPosition.GetPrimary(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItem(PrimaryAcuan, e.RowIndex, EnumAcuanItem.ForPosition)
                        If index < 0 Then
                            Me.dgvScope.Item(Me.ColPrimaryPosition.Index, e.RowIndex).Value = PrimaryAcuan
                            Me.dgvScope.Item(Me.ColJobPositionDesc.Index, e.RowIndex).Value = Me.objPosition.GetDesc(PrimaryAcuan)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication job position code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterJobPosition_Cari
                        ij.txtCari.Text = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.mPrimary <> -1 Then
                            index = Me.GetIndexItem(ij.mPrimary, e.RowIndex, EnumAcuanItem.ForPosition)
                            If index < 0 Then
                                Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value = Me.objPosition.GetCode(ij.mPrimary)
                                Me.dgvScope.Item(Me.ColJobPositionDesc.Index, e.RowIndex).Value = Me.objPosition.GetDesc(ij.mPrimary)
                                Me.dgvScope.Item(Me.ColPrimaryPosition.Index, e.RowIndex).Value = ij.mPrimary
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication job position code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value = Me.objPosition.GetCode(ij.mPrimary)
                                    Me.dgvScope.Item(Me.ColJobPositionDesc.Index, e.RowIndex).Value = Me.objPosition.GetDesc(ij.mPrimary)
                                    Me.dgvScope.Item(Me.ColPrimaryPosition.Index, e.RowIndex).Value = ij.mPrimary
                                End If
                            End If
                        ElseIf ij.mPrimary = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If
                ElseIf e.ColumnIndex = Me.ColEmployeeCode.Index Then
                    'FUNGSI COL EMPLOYEE CODE
                    Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objEmp.ApakahSudahAda(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryAcuan = Me.objEmp.GetPrimaryByCode(Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value)
                        index = Me.GetIndexItem(PrimaryAcuan, e.RowIndex, EnumAcuanItem.ForEmployee)
                        If index < 0 Then
                            Me.dgvScope.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = PrimaryAcuan
                            Me.dgvScope.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.objEmp.GetFullNameByPrimary(PrimaryAcuan)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmMasterKaryawanCari
                        ij.txtCari.Text = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryEmploye <> -1 Then
                            index = Me.GetIndexItem(ij.PrimaryEmploye, e.RowIndex, EnumAcuanItem.ForEmployee)
                            If index < 0 Then
                                Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.objEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                Me.dgvScope.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.objEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                Me.dgvScope.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = ij.PrimaryEmploye
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.objEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvScope.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.objEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.dgvScope.Item(Me.ColPrimaryEmp.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                End If
                            End If
                        ElseIf ij.PrimaryEmploye = -1 Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If
                ElseIf e.ColumnIndex = Me.colRemarks.Index Then
                    Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                End If
            Else
DeleteRows:
                If String.IsNullOrEmpty(Me.Buffer) = True Then
                    If Me.dgvScope.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColDepartmentCode.Index Then
                            If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value) = False Or _
                            String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value) = False Then
                                GoTo ToBuffer
                            End If
                        ElseIf e.ColumnIndex = Me.ColJobPositionCode.Index Then
                            If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value) = False Or _
                            String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value) = False Then
                                GoTo ToBuffer
                            End If
                        ElseIf e.ColumnIndex = Me.ColEmployeeCode.Index Then
                            If String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value) = False Or _
                            String.IsNullOrEmpty(Me.dgvScope.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value) = False Then
                                GoTo ToBuffer
                            End If
                        End If

                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                Else
ToBuffer:
                    If e.ColumnIndex <> Me.colRemarks.Index Then
                        Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
else
                        Me.dgvScope.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvScope_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvScope.RowsAdded
        Dim row As DataGridViewRow = Me.dgvScope.Rows(Me.dgvScope.RowCount - 1)

        Try
            row.Cells(Me.colNo.Index).Value = Me.dgvScope.RowCount
            row.Cells(Me.ColDepartmentCode.Index).Value = ""
            row.Cells(Me.ColDepartmentDesc.Index).Value = ""
            row.Cells(Me.ColJobPositionCode.Index).Value = ""
            row.Cells(Me.ColJobPositionDesc.Index).Value = ""
            row.Cells(Me.ColEmployeeCode.Index).Value = ""
            row.Cells(Me.ColEmployeeName.Index).Value = ""
            row.Cells(Me.colRemarks.Index).Value = ""
            row.Cells(Me.ColPrimaryDept.Index).Value = -1
            row.Cells(Me.ColPrimaryPosition.Index).Value = -1
            row.Cells(Me.ColPrimaryEmp.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvScope_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvScope.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvScope, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvScope_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvScope.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvScope, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvScope_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvScope.UserDeletingRow
        'batasan jika hapus baris item
    End Sub

    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColUserCode.Index Then
                Me.Buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColUserEmail.Index Or e.ColumnIndex = Me.ColUserAllowOverBudgetSPL.Index Or _
            e.ColumnIndex = Me.ColUserApproved.Index Or e.ColumnIndex = Me.ColUserApprovedPosition.Index Or _
            e.ColumnIndex = Me.ColUserRemarks.Index Then
                If String.IsNullOrEmpty(Me.DgvUser.Item(Me.ColUserCode.Index, e.RowIndex).Value) = False Then
                    Me.Buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellContentClick

    End Sub

    Private Function GetIndexItemUser(ByVal vUserCode As String, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1

        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColUserCode.Index).Value = vUserCode Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim UserCode As String = ""

        Try
            If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        index = Me.GetIndexItemUser(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                        If index < 0 Then
                            Me.DgvUser.Item(Me.ColUserPrimaryUser.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.DgvUser.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.objUser.GetNama(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                GoTo DeleteRows
                            End If
                        End If
                    Else
                        Dim ij As New FrmCariMasterUser
                        ij.txtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            index = Me.GetIndexItemUser(ij.Kodeuser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                Me.DgvUser.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.objUser.GetNama(ij.Kodeuser)
                                Me.DgvUser.Item(Me.ColUserPrimaryUser.Index, e.RowIndex).Value = Me.objUser.GetPrimary(ij.Kodeuser)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    GoTo DeleteRows
                                Else
                                    Me.DgvUser.Item(Me.ColUserCode.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.objUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColUserPrimaryUser.Index, e.RowIndex).Value = Me.objUser.GetPrimary(ij.Kodeuser)
                                End If
                            End If
                        ElseIf ij.Kodeuser = "" Then
                            GoTo DeleteRows
                        End If

                        ij = Nothing
                    End If

                    If Me.Obj.ApakahUserSudahDigunakan(Me.txtPrimary.Text, Me.DgvUser.Item(Me.ColUserPrimaryUser.Index, e.RowIndex).Value) = True Then
                        MsgBox(Me.DgvUser.Item(Me.ColUserCode.Index, e.RowIndex).Value & " has been used in other transaction", MsgBoxStyle.Information, "Information")
                        GoTo DeleteRows
                    End If
                ElseIf e.ColumnIndex = Me.ColUserEmail.Index Then
                    'PASTIKAN YANG DIINPUT FORMAT EMAIL
                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If String.IsNullOrEmpty(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                        If Expression.IsMatch(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Expression = Nothing
                    End If
                ElseIf e.ColumnIndex = Me.ColUserRemarks.Index Then
                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
                End If
            Else
DeleteRows:
                If String.IsNullOrEmpty(Me.Buffer) = True Then
                    If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                Else
                    If e.ColumnIndex <> Me.ColUserRemarks.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Else
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim row As DataGridViewRow = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)

        Try
            row.Cells(Me.colNo.Index).Value = Me.DgvUser.RowCount
            row.Cells(Me.ColUserCode.Index).Value = ""
            row.Cells(Me.ColUserName.Index).Value = ""
            row.Cells(Me.ColUserEmail.Index).Value = ""
            row.Cells(Me.ColUserAllowOverBudgetSPL.Index).Value = Boolean.FalseString
            row.Cells(Me.ColUserApproved.Index).Value = ""
            row.Cells(Me.ColUserApprovedPosition.Index).Value = ""
            row.Cells(Me.ColUserRemarks.Index).Value = ""
            row.Cells(Me.ColUserPrimaryUser.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvUser.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvUser, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvUser.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvUser, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvUser.UserDeletingRow
        'batasan jika hapus baris item
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
            Me.dgvScope.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvScope)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.GotFocus, TxtDesc.GotFocus, txtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Me.txtRemarks.Text = Me.txtRemarks.Text.ToString.Replace("'", "").Replace("*", "").Replace("%", "")
    End Sub

    Private Sub txtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.LostFocus, TxtDesc.LostFocus, txtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadonly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CloseForm.DoWork
        Try
            Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_CloseForm_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_CloseForm.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class