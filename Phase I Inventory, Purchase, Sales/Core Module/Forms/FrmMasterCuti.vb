Option Explicit On
Option Strict Off

Public Class FrmMasterCuti
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCuti As New ClsMasterCuti
    Dim ObjItem As New ClsMasterCuti_Item
    Dim ObjKaryawan As New ClsMasterKaryawan
    Dim ObjParameter As New ClsMasterParameterGaji
    Dim buffer As Double = 0
    Dim buffer_dgv As String

    Dim mDgv As DataGridView
    Dim EraseRowIndex As Integer = 0

    ''@rosadinnn
    ''@20150908 1451
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Public vTblItemDesc As New DseMasterCuti.tblmaster_cuti_item_descDataTable
    Dim vRowItemDesc As DseMasterCuti.tblmaster_cuti_item_descRow

    Public Enum EnumTipeLeave As Integer
        Addition
        Deduction
    End Enum

    Public Enum EnumTipeLeave2 As Integer
        MonthlyLeaveBySystem
        LongLeaveBySystem
        Add_Ded_ByUser
    End Enum

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Me.TxtCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCode.KeyPress
        Me.BlokSimbol(Me.TxtCode, e)
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Me.TxtCode.BackColor = Color.White
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Me.TxtDesc.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDesc.KeyPress
        Me.BlokSimbol(Me.TxtDesc, e)
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Me.TxtDesc.BackColor = Color.White
    End Sub

    Private Sub TxtLeaveQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLeaveQty.Enter
        Try
            Me.buffer = Me.TxtLeaveQty.Text
            Me.TxtLeaveQty.Text = Me.ObjCD.BuangNolDibelakangDecimal(Me.TxtLeaveQty.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLeaveQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLeaveQty.GotFocus
        Me.TxtLeaveQty.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLeaveQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLeaveQty.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtLeaveQty, sender, e)
    End Sub

    Private Sub TxtLeaveQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLeaveQty.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtLeaveQty, sender, e, False)
    End Sub

    Private Sub TxtLeaveQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLeaveQty.Leave
        Try
            Me.TxtLeaveQty.Text = _
                       Me.ObjCD.DigitGrouping(Me.ObjCD.MiddleRound(0.5, Me.TxtLeaveQty.Text), 1)
            If CDbl(Me.TxtLeaveQty.Text) < Me.buffer Then
                If Me.ObjCuti.ApakahLeaveEditSah(Me.TxtPrimary.Text, Me, Me.dgvEmployee) = False Then
                    Me.TxtLeaveQty.Text = Me.ObjCD.DigitGrouping(Me.buffer, 1)
                    Exit Sub
                End If
            End If
            For Each rowDGV As DataGridViewRow In Me.dgvEmployee.Rows
                If rowDGV.Index <> Me.dgvEmployee.Rows.Count - 1 Then
                    Me.HitungBalance(rowDGV.Index)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLeaveQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLeaveQty.LostFocus
        Me.TxtLeaveQty.BackColor = Color.White
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Me.TxtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRemarks.KeyPress
        Me.BlokSimbol(Me.TxtRemarks, e)
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Me.TxtRemarks.BackColor = Color.White
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimary As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvEmployee.Rows
                If rowDGV.Cells(Me.ColPrimary.Index).Value = vPrimary Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

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

    Private Sub FrmMasterCuti_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjCuti.SimpanBaru(Me.UserCode, Me.dgvEmployee, Me, Me.vTblItemDesc) = True Then
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
                            If Me.ObjCuti.SimpanPerubahan(Me.UserCode, Me.dgvEmployee, Me, Me.vTblItemDesc) = True Then
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

    Private Sub FrmMasterCuti_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.NonAktifkanDetail()
            Me.dtpPeriod.Value = Me.ObjCD.GetFisrtDateOfTheYear(Today)
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.AllowUserToResizeColumns = True
            Me.dgvEmployee.AllowUserToResizeRows = True
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@20150908 1342
            Me.ChkIsDefault.Enabled = True

            Me.dgvEmployee.AllowUserToAddRows = True
            Me.dgvEmployee.AllowUserToDeleteRows = True
            Me.LblAdvance.Enabled = True
            Me.CancelButton = Me.btnCancel
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@20150908 1342
            Me.ChkIsDefault.Enabled = False

            Me.dgvEmployee.AllowUserToAddRows = False
            Me.dgvEmployee.AllowUserToDeleteRows = False
            Me.LblAdvance.Enabled = False
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjCuti
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Period = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)
                .Qty = Me.TxtLeaveQty.Text
                .Remarks = Me.TxtRemarks.Text

                ''@rosadinnn
                ''@20150908 1340
                If Me.ChkIsDefault.Checked = True Then
                    .IsDefault = 1
                Else
                    .IsDefault = 0
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjCuti
                Me.TxtPrimary.Text = .Primary
                Me.TxtCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.dtpPeriod.Value = .Period
                Me.TxtLeaveQty.Text = Me.ObjCD.DigitGrouping(.Qty, 1)
                Me.TxtRemarks.Text = .Remarks

                ''@rosadinnn
                ''@20150908 1340
                If .IsDefault = 0 Then
                    Me.ChkIsDefault.Checked = False
                Else
                    Me.ChkIsDefault.Checked = True
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, False)

            ''@rosadinnn
            ''@20150908 1343
            Me.ChkIsDefault.Checked = False
            Me.vTblItemDesc.Rows.Clear()

            Me.dgvEmployee.Rows.Clear()
            Me.TxtPrimary.Text = -1
            Me.TxtLeaveQty.Text = 0
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
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)
            KoleksiNamaData.Add("Leave Qty")
            KoleksiTextBox.Add(Me.TxtLeaveQty)
            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            ''@rosadinnn
            ''@20150908 1357
            If Me.ChkIsDefault.Checked = True Then
                If Me.ObjCuti.IsAlreadyDefaultLeave(Me.TxtPrimary.Text, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpPeriod.Value)) = True Then
                    MsgBox("There is already default leave in this period, you cannot save multiple default leave.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'harus memiliki minimal 1 item
            If Me.dgvEmployee.RowCount <= 1 Then
                MsgBox("Please input Employee data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'Qty harus positif
            If Me.ObjCD.CekBilanganPositif(Me.TxtLeaveQty.Text) = False Then
                MsgBox("Qty must positive number.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh double
            If Me.baru = True Then
                If Me.ObjCuti.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjCuti.GetCodeByPrimary(Me.TxtPrimary.Text)
                If CodeLama <> Me.TxtCode.Text Then
                    If Me.ObjCuti.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjCuti.GetTimeStampByPrimary(Me.TxtPrimary.Text)
                If Me.ObjCuti.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjCuti.CariDetail(Me.ObjCuti.Primary, Me, Me.dgvEmployee, Me.ObjCD, Me.vTblItemDesc)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtCode.Focus()
            Me.dtpPeriod.Value = Now
            Me.tslCreatedBy.Text = "Created By: " + UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                hasil = MsgBox("Save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjCuti.SimpanBaru(Me.UserCode, Me.dgvEmployee, Me, Me.vTblItemDesc) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.ObjItem.IsiDGVItem(Me.dgvEmployee, Me.ObjCuti.Primary, Me, Me.ObjCD, Me.vTblItemDesc)
                        End If
                    Else
                        If Me.ObjCuti.SimpanPerubahan(Me.UserCode, Me.dgvEmployee, Me, Me.vTblItemDesc) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.ObjItem.IsiDGVItem(Me.dgvEmployee, Me.ObjCuti.Primary, Me, Me.ObjCD, Me.vTblItemDesc)
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text <> "" And Me.TxtCode.Enabled = True Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjCuti.CariDetail(Me.TxtPrimary.Text, Me, Me.dgvEmployee, Me.ObjCD, Me.vTblItemDesc)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
                Me.NonAktifkanDetail()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterCuti_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryCuti = -1 Then
                If Me.ObjCuti.CariDetail(hj.PrimaryCuti, Me, Me.dgvEmployee, Me.ObjCD, Me.vTblItemDesc) = True Then
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

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtCode.Text <> "" Then
                Me.baru = False
                Me.ObjCuti.CariDetail(Me.TxtPrimary.Text, Me, Me.dgvEmployee, Me.ObjCD, Me.vTblItemDesc)
                If Me.ObjCuti.ApakahBolehEdit(Me.TxtPrimary.Text) = True Then
                    Me.AktifkanDetail()
                    Me.TxtCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyByDepartementToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyByDepartementToolStripMenuItem.Click
        Dim Frm As New FrmMasterDepartementCari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryDepartement <> -1 Then
                Me.ObjCuti.CariDataKaryawan(ClsMasterCuti.enumCariData.dept, Frm.PrimaryDepartement, Me, Me.vTblItemDesc)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAJobPositionToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToAJobPositionToolStripMenuItem.Click
        Dim Frm As New FrmMasterJobPosition_Cari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.mPrimary <> -1 Then
                Me.ObjCuti.CariDataKaryawan(ClsMasterCuti.enumCariData.Job, Frm.mPrimary, Me, Me.vTblItemDesc)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAllEmployeeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToAllEmployeeToolStripMenuItem.Click
        Try
            Me.ObjCuti.CariDataKaryawan(ClsMasterCuti.enumCariData.AllEmp, -1, Me, Me.vTblItemDesc)
            For Each rowdgv As DataGridViewRow In Me.dgvEmployee.Rows
                Me.HitungBalance(rowdgv.Index)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyFromOtherDataToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyFromOtherDataToolStripMenuItem.Click
        Dim Frm As New FrmMasterCuti_Cari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryCuti <> -1 Then
                Me.ObjCuti.CariDataKaryawan(ClsMasterCuti.enumCariData.Other, Frm.PrimaryCuti, Me, Me.vTblItemDesc)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If Me.TxtCode.Enabled = True Then
                If String.IsNullOrEmpty(Me.dgvEmployee(Me.ColCode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColLeaveAddition.Index Or e.ColumnIndex = Me.ColLeaveDeduction.Index Or _
                    e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.buffer_dgv = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value

                        If e.ColumnIndex = Me.ColLeaveAddition.Index Or e.ColumnIndex = Me.ColLeaveDeduction.Index Then
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                        End If
                    Else
                        e.Cancel = True
                    End If
                Else
                    If e.ColumnIndex = Me.ColCode.Index Then
                        Me.buffer_dgv = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellContentClick
        Dim vFrm As New FrmMasterCuti_Balance
        Dim vFrmDesc As New FrmMasterCuti_Item_Desc

        Try
            If e.ColumnIndex = -1 Then
                Exit Sub
            End If
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            'If Me.TxtCode.Enabled = True Then
            If e.ColumnIndex = Me.ColBalance.Index And Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value <> -1 Then
                vFrm.mPeriod = Me.dtpPeriod.Value
                vFrm.mPrimaryLeave = Me.TxtPrimary.Text
                vFrm.mPrimaryEmployee = Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value
                vFrm.ObjCd = Me.ObjCD
                vFrm.FrmMasterLeaveRef = Me
                vFrm.FrmMasterLeaveRef_vIndexRef = e.RowIndex

                vFrm.ShowDialog()
            End If
            'End If

            ''@rosadinnn
            ''@20150908 1501
            Dim DI As Integer = 0
            If e.ColumnIndex = Me.ColLeaveAddition.Index Or e.ColumnIndex = Me.ColLeaveDeduction.Index Then
                vFrmDesc.mIdentifier = Me.dgvEmployee.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                vFrmDesc.FrmRef = Me
                If e.ColumnIndex = Me.ColLeaveAddition.Index Then
                    vFrmDesc.mTipe = EnumTipeLeave.Addition
                ElseIf e.ColumnIndex = Me.ColLeaveDeduction.Index Then
                    vFrmDesc.mTipe = EnumTipeLeave.Deduction
                End If
                vFrmDesc.mYear = Me.dtpPeriod.Value.Year
                vFrmDesc.mTotalAdditionBuffer = CDbl(Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value)
                vFrmDesc.mTotalDeductionBuffer = CDbl(Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value)
                vFrmDesc.mPrimaryKaryawan = CDbl(Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                If Me.TxtRemarks.Enabled = False Then
                    vFrmDesc.mReadOnly = True
                    vFrmDesc.ShowDialog()
                    Exit Sub
                End If
                If vFrmDesc.ShowDialog = Windows.Forms.DialogResult.OK Then
                    DI = 0
                    While DI <= Me.vTblItemDesc.Rows.Count - 1
                        Me.vRowItemDesc = Me.vTblItemDesc.Rows(DI)
                        If Me.vRowItemDesc.identifier = vFrmDesc.mIdentifier Then
                            Me.vTblItemDesc.Rows.Remove(Me.vRowItemDesc)
                        Else
                            DI += 1
                        End If
                    End While
                    vFrmDesc.DgvItem.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vFrmDesc.DgvItem.Rows
                        vRowItemDesc = Me.vTblItemDesc.Newtblmaster_cuti_item_descRow
                        With vRowItemDesc
                            .primary = -1
                            .primary_leave = Me.TxtPrimary.Text
                            .identifier = vFrmDesc.mIdentifier
                            If rowDGV.Cells(vFrmDesc.ColTipe.Index).Value = "Addition" Then
                                .tipe = FrmMasterCuti.EnumTipeLeave.Addition
                            ElseIf rowDGV.Cells(vFrmDesc.ColTipe.Index).Value = "Deduction" Then
                                .tipe = FrmMasterCuti.EnumTipeLeave.Deduction
                            End If
                            ._date = CDate(rowDGV.Cells(vFrmDesc.ColDate.Index).Value)
                            .desc = rowDGV.Cells(vFrmDesc.ColDesc.Index).Value
                            .qty = CDbl(rowDGV.Cells(vFrmDesc.ColQty.Index).Value)
                            If rowDGV.Cells(vFrmDesc.ColTipe2.Index).Value = "MonthlyLeaveBySystem" Then
                                .tipe2 = FrmMasterCuti.EnumTipeLeave2.MonthlyLeaveBySystem
                            ElseIf rowDGV.Cells(vFrmDesc.ColTipe2.Index).Value = "LongLeaveBySystem" Then
                                .tipe2 = FrmMasterCuti.EnumTipeLeave2.LongLeaveBySystem
                            ElseIf rowDGV.Cells(vFrmDesc.ColTipe2.Index).Value = "Add_Ded_ByUser" Then
                                .tipe2 = FrmMasterCuti.EnumTipeLeave2.Add_Ded_ByUser
                            End If
                        End With
                        Me.vTblItemDesc.Rows.Add(vRowItemDesc)
                    Next
                    Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(vFrmDesc.mTotalAddition, 1)
                    Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(vFrmDesc.mTotalDeduction, 1)
                    Me.dgvEmployee.Item(Me.ColBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDbl(Me.TxtLeaveQty.Text) + _
                            CDbl(Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value) - _
                            CDbl(Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value) - _
                            Me.ObjItem.GetNilaiTerpakai(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value), 1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vFrm.Dispose()
        End Try
    End Sub

    Private Sub HitungBalance(ByVal vRowIndex As Integer)
        Try
            Me.dgvEmployee.Item(Me.ColBalance.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(CDbl(Me.TxtLeaveQty.Text) + _
                                  CDbl(Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, vRowIndex).Value) - _
                                  CDbl(Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, vRowIndex).Value) - _
                             Me.ObjItem.GetNilaiTerpakai(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, vRowIndex).Value), 1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double
        Dim RowIdx As Integer = e.RowIndex
        Try
            If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                      Or Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                With row
                    '===code -> seharusnya tidak ke sini
                    If e.ColumnIndex = Me.ColCode.Index Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    '===Addition/Deduction
                    If e.ColumnIndex = Me.ColLeaveAddition.Index Or e.ColumnIndex = Me.ColLeaveDeduction.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                        Exit Sub
                    End If

                    '===remarks
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Exit Sub
                    End If
                End With
            Else
                With row
                    '===code
                    If e.ColumnIndex = Me.ColCode.Index Then
                        If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If Me.ObjKaryawan.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                Primary = Me.ObjKaryawan.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemInput(Primary)
                                If index < 0 Then
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKaryawan.GetCodeByPrimary(Primary)
                                    Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = Me.ObjKaryawan.GetFullNameByPrimary(Primary)
                                    Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value = Primary
                                    Me.HitungBalance(e.RowIndex)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                                        If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                            Me.mDgv = Me.dgvEmployee
                                            Me.EraseRowIndex = row.Index
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                            Me.BackgroundWorker1.Dispose()
                                        End If
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New FrmMasterKaryawanCari
                                ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                index = Me.GetIndexItemInput(ij.PrimaryEmploye)
                                If ij.PrimaryEmploye = -1 Then
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.mDgv = Me.dgvEmployee
                                    Me.EraseRowIndex = row.Index
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Me.BackgroundWorker1.Dispose()
                                Else
                                    If index < 0 Then
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKaryawan.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = Me.ObjKaryawan.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.HitungBalance(e.RowIndex)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                                            If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                                Me.mDgv = Me.dgvEmployee
                                                Me.EraseRowIndex = row.Index
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                                Me.BackgroundWorker1.Dispose()
                                            End If
                                            Exit Sub
                                        Else
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjKaryawan.GetCodeByPrimary(ij.PrimaryEmploye)
                                            Me.HitungBalance(e.RowIndex)
                                        End If
                                    End If
                                End If
                                ij = Nothing
                            End If
                        Else
                            If Me.buffer_dgv <> "" Then
                                Me.mDgv = Me.dgvEmployee
                                Me.EraseRowIndex = row.Index
                                Me.BackgroundWorker1.RunWorkerAsync()
                                Me.BackgroundWorker1.Dispose()
                            End If
                        End If
                    End If

                    '===Addition
                    If e.ColumnIndex = Me.ColLeaveAddition.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Leave addition must input positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                            Exit Sub
                        Else
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.ObjCD.MiddleRound(0.5, Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value), 1)

                            If CDbl(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) < CDbl(Me.buffer_dgv) Then
                                If Me.ObjCuti.ApakahEditSah(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value, _
                                Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value, _
                                CDbl(Me.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.add) = False Then
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                                    Exit Sub
                                End If
                            End If

                            Me.dgvEmployee.Item(Me.ColBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDbl(Me.TxtLeaveQty.Text) + _
                             CDbl(Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value) - _
                             CDbl(Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value) - _
                             Me.ObjItem.GetNilaiTerpakai(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value), 1)
                        End If
                    End If

                    '===Deduction
                    If e.ColumnIndex = Me.ColLeaveDeduction.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Leave addition must input positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                            Exit Sub
                        Else
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = _
                             Me.ObjCD.DigitGrouping(Me.ObjCD.MiddleRound(0.5, Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value), 1)

                            If CDbl(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) > CDbl(Me.buffer_dgv) Then
                                If Me.ObjCuti.ApakahEditSah(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value, _
                                Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value, Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, _
                                CDbl(Me.TxtLeaveQty.Text), ClsMasterCuti.EnumAddDed.ded) = False Then
                                    Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_dgv
                                    Exit Sub
                                End If
                            End If

                            Me.dgvEmployee.Item(Me.ColBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDbl(Me.TxtLeaveQty.Text) + _
                            CDbl(Me.dgvEmployee.Item(Me.ColLeaveAddition.Index, e.RowIndex).Value) - _
                            CDbl(Me.dgvEmployee.Item(Me.ColLeaveDeduction.Index, e.RowIndex).Value) - _
                             Me.ObjItem.GetNilaiTerpakai(Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.RowIndex).Value), 1)
                        End If
                    End If

                    '===remarks
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    End If
                End With
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Dim rowD As DataGridViewRow

        Try
            Me.ObjCD.SetNumber(Me.dgvEmployee, False)
            rowD = Me.dgvEmployee.Rows(e.RowIndex)
            rowD.Cells(Me.ColLeaveAddition.Index).Value = "0.0"
            rowD.Cells(Me.ColLeaveDeduction.Index).Value = "0.0"
            rowD.Cells(Me.ColBalance.Index).Value = "0.0"
            rowD.Cells(Me.ColPrimary.Index).Value = -1

            ''@rosadinnn
            ''@20150908 1424
            If IsNothing(rowD.Cells(Me.ColPrimary.Index).Value) = False Then
                rowD.Cells(Me.ColIdentifier.Index).Value = Me.GetIdentifierBaru
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvEmployee.RowsRemoved
        Me.ObjCD.SetNumber(Me.dgvEmployee, False)
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.ObjCuti.ApakahBolehHapus(Me.TxtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjCuti.Hapus(Me.TxtPrimary.Text, Me.UserCode) = True Then
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.mDgv.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvEmployee.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvEmployee.UserDeletingRow
        Dim Dta1 As New DseMasterCutiTableAdapters.DataTable2TableAdapter
        Dim tbl1 As New DseMasterCuti.DataTable2DataTable
        Dim row1 As DseMasterCuti.DataTable2Row
        Dim Dta2 As New DseMasterCutiTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim tbl2 As New DseMasterCuti.tbltrans_cuti_pengajuanDataTable
        Dim row2 As DseMasterCuti.tbltrans_cuti_pengajuanRow
        Dim list As String = ""
        Try
            If Me.baru = False Then
                Dta1.Fill(tbl1, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.Row.Index).Value, Me.TxtPrimary.Text)
                If tbl1.Rows.Count > 0 Then
                    Dim i As Integer = 1
                    For Each row1 In tbl1.Rows
                        list = list + i.ToString + ". " + row1.no_transaksi.ToString + " Period : " + Format(row1.start_date.Date, "dd MMM yyyy") + " - " + Format(row1.end_date.Date, "dd MMM yyyy") + vbNewLine
                        i += 1
                    Next
                    MsgBox("Can't delete this data because this data is refering to attendance transaction" + vbNewLine + _
                    list, MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If

                list = ""
                Dta2.FillBy_PrimaryCutiNPrimaryEmp(tbl2, Me.TxtPrimary.Text, Me.dgvEmployee.Item(Me.ColPrimary.Index, e.Row.Index).Value)
                If tbl2.Rows.Count > 0 Then
                    Dim i As Integer = 1
                    For Each row2 In tbl2.Rows
                        list = list + i.ToString + ". " + row2.no_transaksi.ToString + vbNewLine
                        i += 1
                    Next
                    MsgBox("Can't delete this data because this data is refering to leave application" + vbNewLine + _
                    list, MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            ''@rosadinnn
            ''@20150909 1615
            Me.vTblItemDesc.DefaultView.RowFilter = "identifier = '" + Me.dgvEmployee.Item(Me.ColIdentifier.Index, e.Row.Index).Value.ToString + "'"
            While Me.vTblItemDesc.DefaultView.Count > 0
                Me.vRowItemDesc = Me.vTblItemDesc.DefaultView(0).Row
                Me.vTblItemDesc.Rows.Remove(Me.vRowItemDesc)
            End While
            Me.vTblItemDesc.DefaultView.RowFilter = ""

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta1.Dispose()
            tbl1.Dispose()
            Dta2.Dispose()
            tbl2.Dispose()
        End Try
    End Sub

    Private Sub LblAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAdvance.LinkClicked
        Me.ContextMenuStrip1.Show(Me.LblAdvance, 0, 20)
    End Sub

    ''@rosadinnn
    ''20150908 1422
    Public Function GetIdentifierBaru() As String
        Dim IdentifierD As Integer = 0
        Dim NilaiMax As Integer = 0
        Dim Retval As String = ""
        Try
            For Each rowDGV As DataGridViewRow In Me.dgvEmployee.Rows
                If IsNothing(rowDGV.Cells(Me.ColIdentifier.Index).Value) = False Then
                    IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.IndexOf("-", 1)
                    IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.Substring(IdentifierD + 1)
                    If NilaiMax < IdentifierD Then
                        NilaiMax = IdentifierD
                    End If
                End If
            Next
            NilaiMax += 1
            Retval = Me.TxtPrimary.Text + "-" + NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ''@rosadinnn
    ''@20150910 0812
    Private Sub AddMonthlyLeaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddMonthlyLeaveToolStripMenuItem.Click
        Try
            AddLeaveBySystem(EnumTipeLeave2.MonthlyLeaveBySystem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddLongLeaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddLongLeaveToolStripMenuItem.Click
        Try
            AddLeaveBySystem(EnumTipeLeave2.LongLeaveBySystem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetKaryawanAktif(ByRef ListKaryawanAktif As List(Of Double))
        Dim DtaEmp As New DseMasterCutiTableAdapters.tblmaster_karyawanTableAdapter
        Dim TblEmp As New DseMasterCuti.tblmaster_karyawanDataTable
        Dim rowEmp As DseMasterCuti.tblmaster_karyawanRow

        Try
            ListKaryawanAktif.Clear()

            DtaEmp.Connection.ConnectionString = Core_Module.GetConstring
            DtaEmp.Fill(TblEmp, 0, 0)
            For Each rowEmp In TblEmp
                If ListKaryawanAktif.Contains(rowEmp.Primary) = False Then
                    ListKaryawanAktif.Add(rowEmp.Primary)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function AddLeaveBySystem(ByVal vTipe2 As FrmMasterCuti.EnumTipeLeave2) As Boolean
        Dim FrmPeriod As New FrmMasterCuti_Period
        Dim RowDgv As DataGridViewRow
        Dim LoS As Decimal = 72
        Dim ListKaryAktif As New List(Of Double)

        Try
            If Me.dgvEmployee.Rows.Count = 1 Then
                MsgBox("Please input at least 1 item.", MsgBoxStyle.Information, "Information")
                Exit Function
            End If
            FrmPeriod.mPeriod = Me.dtpPeriod.Value
            If FrmPeriod.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.GetKaryawanAktif(ListKaryAktif) 'new

                For Each RowDgv In Me.dgvEmployee.Rows
                    If IsNothing(RowDgv.Cells(Me.ColPrimary.Index).Value) = False Then
                        If RowDgv.Cells(Me.ColPrimary.Index).Value <> "-1" Then
                            If ListKaryAktif.Contains(RowDgv.Cells(Me.ColPrimary.Index).Value) = True Then 'hanya berlaku untuk karyawan yang aktif saja
                                LoS = Me.ObjKaryawan.GetLengthOfService_Months_Total(RowDgv.Cells(Me.ColPrimary.Index).Value, FrmPeriod.DtpPeriod.Value.Date)
                                While LoS > 96
                                    LoS = LoS - 96
                                End While

                                If vTipe2 = EnumTipeLeave2.MonthlyLeaveBySystem Then
                                    ''==hanya berlaku untuk karyawan yang LoS nya < 7 tahun masa kerja
                                    If LoS < 72 Then
                                        Me.vTblItemDesc.DefaultView.RowFilter = "[identifier] = '" & RowDgv.Cells(Me.ColIdentifier.Index).Value.ToString & "' AND [date] = #" & Format(FrmPeriod.DtpPeriod.Value, "MM/dd/yyyy") & "# AND [tipe2] = " & FrmMasterCuti.EnumTipeLeave2.MonthlyLeaveBySystem
                                        ''==hanya tambah yang belum dapet cuti bulanan atas periode ybs
                                        If Me.vTblItemDesc.DefaultView.Count = 0 Then
                                            vRowItemDesc = Me.vTblItemDesc.Newtblmaster_cuti_item_descRow
                                            With vRowItemDesc
                                                .primary = -1
                                                .primary_leave = Me.TxtPrimary.Text
                                                .identifier = RowDgv.Cells(Me.ColIdentifier.Index).Value.ToString
                                                .tipe = FrmMasterCuti.EnumTipeLeave.Addition
                                                ._date = FrmPeriod.DtpPeriod.Value
                                                .desc = "Monthly leave"
                                                .qty = 1
                                                .tipe2 = FrmMasterCuti.EnumTipeLeave2.MonthlyLeaveBySystem
                                            End With
                                            Me.vTblItemDesc.Rows.Add(vRowItemDesc)
                                            RowDgv.Cells(Me.ColLeaveAddition.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDgv.Cells(Me.ColLeaveAddition.Index).Value) + 1, 1)
                                            RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDgv.Cells(Me.ColBalance.Index).Value) + 1, 1)
                                        End If
                                    End If
                                ElseIf vTipe2 = EnumTipeLeave2.LongLeaveBySystem Then
                                    ''==hanya berlaku untuk karyawan yg LoS masuk tahun ke 7 dan masuk ke tahun 8
                                    If (LoS >= 72 And LoS < 73) Or (LoS >= 84 And LoS < 85) Then
                                        Me.vTblItemDesc.DefaultView.RowFilter = "[identifier] = '" & RowDgv.Cells(Me.ColIdentifier.Index).Value.ToString & "' AND [date] = #" & Format(FrmPeriod.DtpPeriod.Value, "MM/dd/yyyy") & "# AND [tipe2] = " & FrmMasterCuti.EnumTipeLeave2.LongLeaveBySystem
                                        ''==hanya tambah yang belum dapet long leave atas periode ybs
                                        If Me.vTblItemDesc.DefaultView.Count = 0 Then
                                            vRowItemDesc = Me.vTblItemDesc.Newtblmaster_cuti_item_descRow
                                            With vRowItemDesc
                                                .primary = -1
                                                .primary_leave = Me.TxtPrimary.Text
                                                .identifier = RowDgv.Cells(Me.ColIdentifier.Index).Value.ToString
                                                .tipe = FrmMasterCuti.EnumTipeLeave.Addition
                                                ._date = FrmPeriod.DtpPeriod.Value
                                                .desc = "Long leave"
                                                .qty = 21
                                                .tipe2 = FrmMasterCuti.EnumTipeLeave2.LongLeaveBySystem
                                            End With
                                            Me.vTblItemDesc.Rows.Add(vRowItemDesc)
                                            RowDgv.Cells(Me.ColLeaveAddition.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDgv.Cells(Me.ColLeaveAddition.Index).Value) + 21, 1)
                                            RowDgv.Cells(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDgv.Cells(Me.ColBalance.Index).Value) + 21, 1)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPeriod.Dispose()
        End Try
    End Function
End Class