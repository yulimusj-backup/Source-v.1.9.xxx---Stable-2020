Option Explicit On
Option Strict Off

Public Class FrmMasterWorkUnit
    Public UserCode As String = "USER"
    Public PrimaryEmployee As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objWU As New ClsMasterWorkUnit
    Dim objWUItem As New ClsMasterWorkUnitItem
    Dim objShift As New ClsMasterWorkShift
    Dim objEmployee As New ClsMasterKaryawan
    Dim objMachine As New ClsMasterMachineData
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objSetup As New Core_Module.ClsMasterSetupTime
    Dim buffer As String = ""
    Dim EraseRowIndex As Integer = 0

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryUnit.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.DgvEmployee.ReadOnly = False
            Me.DgvEmployee.AllowUserToAddRows = True
            Me.DgvEmployee.AllowUserToDeleteRows = True
            Me.DgvMachine.ReadOnly = False
            Me.DgvMachine.AllowUserToAddRows = True
            Me.DgvMachine.AllowUserToDeleteRows = True
            Me.dgvSetuptime.ReadOnly = False
            Me.dgvSetuptime.AllowUserToAddRows = True
            Me.dgvSetuptime.AllowUserToDeleteRows = True
            Me.ChkActive.Enabled = True
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.DgvEmployee.AllowUserToAddRows = False
            Me.DgvEmployee.AllowUserToDeleteRows = False
            Me.DgvEmployee.ReadOnly = True
            Me.DgvMachine.AllowUserToAddRows = False
            Me.DgvMachine.AllowUserToDeleteRows = False
            Me.DgvMachine.ReadOnly = True
            Me.dgvSetuptime.AllowUserToAddRows = False
            Me.dgvSetuptime.AllowUserToDeleteRows = False
            Me.dgvSetuptime.ReadOnly = True
            Me.ChkActive.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateNumbering()
        Try
            'cek apakah sudah ada
            '1. transaksi-nya belum ada di tabel numbering
            '2. data pada table ybs masih kosong

            'jika belum ada, maka insert manual
            'jika sudah ada, exit sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.objWU.ResetObject(Me.DgvEmployee)
            Me.objWU.ResetObject(Me.DgvMachine)
            Me.objWU.ResetObject(Me.dgvSetuptime)
            Me.TxtPrimaryUnit.Text = "-1"
            Me.TxtUnitCode.Text = ""
            Me.TxtUnitDesc.Text = ""
            Me.TxtRemarks.Text = ""
            Me.txtCodeLocationIn.Text = ""
            Me.txtCodeLocationOut.Text = ""
            Me.txtDescLocationIn.Text = ""
            Me.txtDescLocationOut.Text = ""
            Me.txtPrimaryLocationIn.Text = "-1"
            Me.txtPrimaryLocationOut.Text = "-1"
            Me.tslModifiedBy.Text = "ModifiedBy: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.tslCreatedBy.Text = "Created  By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.DgvEmployee.Rows.Clear()
            Me.DgvMachine.Rows.Clear()
            Me.dgvSetuptime.Rows.Clear()
            Me.ChkActive.Checked = True
            Me.buffer = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Work Unit Code")
            KoleksiTextBox.Add(Me.TxtUnitCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtUnitDesc)
            KoleksiNamaData.Add("Group Code")
            KoleksiTextBox.Add(Me.txtGroupCode)
            KoleksiNamaData.Add("Description Group")
            KoleksiTextBox.Add(Me.txtGroupDesc)
            KoleksiNamaData.Add("Code Location In")
            KoleksiTextBox.Add(Me.txtCodeLocationIn)
            KoleksiNamaData.Add("Code Location Out")
            KoleksiTextBox.Add(Me.txtCodeLocationOut)
            KoleksiNamaData.Add("Production Line")
            KoleksiTextBox.Add(Me.TxtProductionLineCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objWU.ApakahSudahAda(Me.TxtUnitCode.Text) = True Then
                    MsgBox("Duplication in work unit code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtUnitCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objWU.GetCode(Me.TxtPrimaryUnit.Text)
                If ProductGroupLama <> Me.TxtUnitCode.Text Then
                    If Me.objWU.ApakahSudahAda(Me.TxtUnitCode.Text) = True Then
                        MsgBox("Duplication in work unit code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtUnitCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objWU.GetTimestamp(Me.TxtPrimaryUnit.Text)
                If Me.objWU.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objWU.CariDetail(Me.objWU.Primary, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'Item tidak boleh kosong 
            If Me.DgvMachine.Rows.Count = 1 AndAlso Me.DgvEmployee.Rows.Count = 1 AndAlso Me.dgvSetuptime.Rows.Count = 1 Then
                If MsgBox("There is no machine nor employee nor Setup time refering to this work unit, continue save anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            ''Item tidak boleh kosong 
            'If Me.DgvEmployee.Rows.Count = 1 Then
            '    MsgBox("Please input item data employee", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If

            Dim rowDgvShift As DataGridViewRow
            For Each rowDgvShift In Me.DgvEmployee.Rows
                If rowDgvShift.Cells(Me.ColEmployee_EmployeeCode.Index).Value <> "" Then
                    If String.IsNullOrEmpty(Trim(rowDgvShift.Cells(Me.ColEmployee_ShiftCode.Index).Value)) = True Then
                        Me.DgvEmployee.ClearSelection()
                        rowDgvShift.Selected = True
                        MsgBox("Work Shift code item cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            Dim rowDgvEmp As DataGridViewRow
            For Each rowDgvEmp In Me.DgvEmployee.Rows
                If rowDgvEmp.Cells(Me.ColEmployee_ShiftCode.Index).Value <> "" Then
                    If String.IsNullOrEmpty(Trim(rowDgvEmp.Cells(Me.ColEmployee_EmployeeCode.Index).Value)) = True Then
                        Me.DgvEmployee.ClearSelection()
                        rowDgvEmp.Selected = True
                        MsgBox("Employee code item cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next



            Dim rowDGVSetup As New DataGridViewRow
            For Each rowDGVSetup In Me.dgvSetuptime.Rows
                If rowDGVSetup.Index <> Me.dgvSetuptime.Rows.Count - 1 Then
                    If rowDGVSetup.Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = -1 Then
                        MsgBox("Process code 2 cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If rowDGVSetup.Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1 Then
                        MsgBox("Product Code 1 cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If rowDGVSetup.Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1 Then
                        MsgBox("Product Code 2 cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If rowDGVSetup.Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1 Then
                        MsgBox("Setup time code cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            ''primary location in dan out tidak boleh sama
            'If Me.txtPrimaryLocationIn.Text = Me.txtPrimaryLocationOut.Text Then
            '    MsgBox("Location in with location out cannot same.", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

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
        Try
            With Me.objWU
                .WorkUnitCode = Me.TxtUnitCode.Text
                .PrimaryWorkUnitGroup = Me.TxtPrimaryWorkUnitGroup.Text
                .Desc = Me.TxtUnitDesc.Text
                .Remarks = Me.TxtRemarks.Text
                .PrimaryLocationIN = Me.txtPrimaryLocationIn.Text
                .PrimaryLocationOUT = Me.txtPrimaryLocationOut.Text
                .CodeWUGroup = Me.txtGroupCode.Text
                .DescWUGroup = Me.txtGroupDesc.Text
                .CodeLocIN = Me.txtCodeLocationIn.Text
                .CodeLocOUT = Me.txtCodeLocationOut.Text
                .DescLocIN = Me.txtDescLocationIn.Text
                .DescLocOUT = Me.txtDescLocationOut.Text

                If Me.ChkActive.Checked = False Then
                    .Active = ClsMasterWorkUnit.enumActive.NonActive
                Else
                    .Active = ClsMasterWorkUnit.enumActive.Active
                End If
                .PrimaryProductionLine = Me.TxtPrimaryProductionLine.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim objGroup As New ClsMasterWorkUnitGroup
        Try
            With Me.objWU
                Me.TxtUnitCode.Text = .WorkUnitCode
                Me.TxtUnitDesc.Text = .Desc
                Me.TxtRemarks.Text = .Remarks
                Me.TxtPrimaryUnit.Text = .Primary
                Me.TxtPrimaryWorkUnitGroup.Text = .PrimaryWorkUnitGroup
                Me.txtGroupCode.Text = .CodeWUGroup
                Me.txtGroupDesc.Text = .DescWUGroup
                If .Active = ClsMasterReject.enumActive.NonActive Then
                    Me.ChkActive.Checked = ClsMasterWorkUnit.enumActive.NonActive
                Else
                    Me.ChkActive.Checked = ClsMasterWorkUnit.enumActive.Active
                End If
                Me.txtPrimaryLocationIn.Text = .PrimaryLocationIN
                Me.txtPrimaryLocationOut.Text = .PrimaryLocationOUT
                Me.txtCodeLocationIn.Text = .CodeLocIN
                Me.txtCodeLocationOut.Text = .CodeLocOUT
                Me.txtDescLocationIn.Text = .DescLocIN
                Me.txtDescLocationOut.Text = .DescLocOUT

                Dim objLine As New Core_Module.ClsMasterMachineLine
                Me.TxtPrimaryProductionLine.Text = .PrimaryProductionLine
                Me.TxtProductionLineCode.Text = objLine.GetCode(.PrimaryProductionLine)
                Me.TxtProductionLineDesc.Text = objLine.GetDescription(.PrimaryProductionLine)

                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.objWU = New ClsMasterWorkUnit
            Me.dgvSetuptime.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvMachine.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.baru = True
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objWU.SimpanBaru(Me.UserCode, Me.TxtPrimaryUnit.Text, Me.ObjCD, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me) = True Then
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
                        If Me.objWU.SimpanPerubahan(Me.UserCode, Me.TxtPrimaryUnit.Text, Me.ObjCD, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me) = True Then
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
            If Not Me.TxtUnitCode.Text = "" And Me.TxtUnitDesc.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objWU.CariDetail(Me.TxtPrimaryUnit.Text, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me.ObjCD, Me)
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
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()

            Me.TxtUnitCode.Focus()
            Me.tslModifiedBy.Text = "Created By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslCreatedBy.Text = "Modified By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterWorkUnit_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Status = "Show All"
            hj.ShowDialog()
            If Not hj.PrimaryWorkUnit = -1 Then
                If Me.objWU.CariDetail(hj.PrimaryWorkUnit, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me.ObjCD, Me) = True Then
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
        If Me.TxtUnitCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.TxtUnitCode.Text <> "" Then
            Me.baru = False
            Me.AktifkanDetail()
            Me.TxtUnitCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtUnitCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtUnitCode.Text <> "" Then
                If Me.objWU.ApakahBolehHapus(Me.TxtPrimaryUnit.Text, Me.TxtUnitCode.Text) = False Then

                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objWU.Hapus(Me.TxtPrimaryUnit.Text, Me.UserCode, True, True, Me.DgvEmployee, Me) = True Then
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

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtUnitCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objWU.SimpanBaru(Me.UserCode, Me.TxtPrimaryUnit.Text, _
                            Me.ObjCD, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me) = True Then
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
                            If Me.objWU.SimpanPerubahan(Me.UserCode, Me.TxtPrimaryUnit.Text, _
                            Me.ObjCD, Me.DgvEmployee, Me.DgvMachine, Me.dgvSetuptime, Me) = True Then
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

    Private Sub TxtUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitCode.GotFocus, TxtUnitDesc.GotFocus, _
    TxtRemarks.GotFocus, txtGroupCode.GotFocus, txtCodeLocationIn.GotFocus, txtCodeLocationOut.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
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

    Private Sub TxtUnitCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitCode.LostFocus, TxtUnitDesc.LostFocus, _
        TxtRemarks.LostFocus, txtGroupCode.LostFocus, txtCodeLocationIn.LostFocus, txtCodeLocationOut.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryUnit.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvEmployee.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvEmployee.Item(Me.ColEmployee_ShiftCode.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColEmployee_ShiftCode.Index Or e.ColumnIndex = Me.ColEmployee_EmployeeCode.Index) Then
                    'Me.DgvEmployee.Columns(Me.ColEmployee_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_EmployeeCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_ShiftCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
                If String.IsNullOrEmpty(Me.DgvEmployee.Item(Me.ColEmployee_EmployeeCode.Index, e.RowIndex).Value) = False Then
                    If (e.ColumnIndex = Me.ColEmployee_Remarks.Index) Then
                        'Me.DgvEmployee.Columns(Me.ColEmployee_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_EmployeeCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_ShiftCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        Me.buffer = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                        e.Cancel = False
                    Else
                        'e.Cancel = True
                    End If
                Else
                    If e.ColumnIndex = Me.ColEmployee_EmployeeCode.Index Then
                        'Me.DgvEmployee.Columns(Me.ColEmployee_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_EmployeeCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        'Me.DgvEmployee.Columns(Me.ColEmployee_ShiftCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                        Me.buffer = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        'e.Cancel = True
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColEmployee_ShiftCode.Index Then
                    'Me.DgvEmployee.Columns(Me.ColEmployee_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_EmployeeCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvEmployee.Columns(Me.ColEmployee_ShiftCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
            If e.RowIndex > 0 Then
                If String.IsNullOrEmpty(Me.DgvEmployee.Item(Me.ColEmployee_EmployeeCode.Index, e.RowIndex - 1).Value) Or _
                String.IsNullOrEmpty(Me.DgvEmployee.Item(Me.ColEmployee_ShiftCode.Index, e.RowIndex - 1).Value) = True Then
                    If e.ColumnIndex = Me.ColEmployee_ShiftCode.Index Then
                        e.Cancel = True
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemEmployee(ByVal vPrimaryEmployee As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvEmployee.Rows.Count - 2
                RowDGV = Me.DgvEmployee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColEmployee_Primary.Index).Value = vPrimaryEmployee Then
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

    Private Sub DgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryShift As Double = -1
        Dim PrimaryEmployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColEmployee_ShiftCode.Index Then
                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objShift.ApakahSudahAda(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.Production) = True Then
                            PrimaryShift = Me.objShift.GetPrimary(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            Me.DgvEmployee.Item(Me.ColEmployee_PrimaryShift.Index, e.RowIndex).Value = Me.objShift.GetPrimary(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            Me.DgvEmployee.Item(Me.ColEmployee_ShiftCode.Index, e.RowIndex).Value = Me.objShift.GetCodeByPrimary(Me.DgvEmployee.Item(Me.ColEmployee_PrimaryShift.Index, e.RowIndex).Value)
                        Else
                            Dim ij As New FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.Production)
                            ij.txtCari.Text = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWorkshift <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Me.DgvEmployee.Item(Me.ColEmployee_PrimaryShift.Index, e.RowIndex).Value = ij.PrimaryWorkshift
                                Me.DgvEmployee.Item(Me.ColEmployee_ShiftCode.Index, e.RowIndex).Value = Me.objShift.GetCodeByPrimary(ij.PrimaryWorkshift)
                            ElseIf ij.PrimaryWorkshift = -1 Then
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    If Me.DgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If
                    If e.ColumnIndex = Me.ColEmployee_EmployeeCode.Index Then
                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objEmployee.ApakahSudahAda(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            PrimaryEmployee = Me.objEmployee.GetPrimaryByCode(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemEmployee(PrimaryEmployee, e.RowIndex)
                            If index < 0 Then
                                Me.DgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = Me.objEmployee.GetPrimaryByCode(Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.DgvEmployee.Item(Me.ColEmployee_EmployeeCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(Me.DgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value)
                                Me.DgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(Me.DgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Else
                                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvEmployee.Item(Me.ColEmployee_EmployeeCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.DgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.DgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                        Else
                                            Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.DgvEmployee.Item(Me.ColEmployee_EmployeeCode.Index, e.RowIndex).Value = Me.objEmployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.DgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.objEmployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.DgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                'If Me.DgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                '    Me.EraseRowIndex = e.RowIndex
                                '    Me.BackgroundWorker1.RunWorkerAsync()
                                'End If
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Else
                                    Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvEmployee)
                    End If
                Else
                    If e.ColumnIndex = Me.ColEmployee_Remarks.Index Then
                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColEmployee_ShiftCode.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            If Me.DgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        Else
                            Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvEmployee.Columns(Me.ColEmployee_No.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvEmployee.Columns(Me.ColEmployee_EmployeeCode.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvEmployee.Columns(Me.ColEmployee_Remarks.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvEmployee.Columns(Me.ColEmployee_ShiftCode.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub DgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvEmployee.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.TxtUnitDesc.Enabled = True Then
                rowDGV = Me.DgvEmployee.Rows(Me.DgvEmployee.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColEmployee_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColEmployee_ShiftCode.Index).Value = ""
                    .Cells(Me.ColEmployee_EmployeeCode.Index).Value = ""
                    .Cells(Me.ColEmployee_Name.Index).Value = ""
                    .Cells(Me.ColEmployee_Remarks.Index).Value = ""
                    .Cells(Me.ColEmployee_Primary.Index).Value = "-1"
                    .Cells(Me.ColEmployee_PrimaryShift.Index).Value = "-1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmployee_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvEmployee.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvEmployee_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvEmployee.UserDeletedRow
        Try
            'panggil fungsi set number
            Me.ObjCD.SetNumber(Me.DgvEmployee)
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
            Me.DgvEmployee.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMachine_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMachine.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvMachine.Item(Me.ColMachine_MachineCode.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColMachine_MachineCode.Index Or e.ColumnIndex = Me.ColMachine_Remarks.Index) Then
                    'Me.DgvMachine.Columns(Me.ColMachine_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_MachineCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColMachine_MachineCode.Index Then
                    'Me.DgvMachine.Columns(Me.ColMachine_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_MachineCode.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvMachine.Columns(Me.ColMachine_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemMachine(ByVal vPrimaryMachine As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvMachine.Rows.Count - 2
                RowDGV = Me.DgvMachine.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColMachine_PrimaryMachine.Index).Value = vPrimaryMachine Then
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

    Private Sub DgvMachine_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMachine.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvMachine.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColMachine_MachineCode.Index Then
                        Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.objMachine.ApakahSudahAdaOnlyActive(Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Primary = Me.objMachine.GetPrimary(Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemMachine(Primary, e.RowIndex)
                            If index < 0 Then
                                Me.DgvMachine.Item(Me.ColMachine_PrimaryMachine.Index, e.RowIndex).Value = Me.objMachine.GetPrimary(Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.DgvMachine.Item(Me.ColMachine_MachineCode.Index, e.RowIndex).Value = Me.objMachine.GetCode(Me.DgvMachine.Item(Me.ColMachine_PrimaryMachine.Index, e.RowIndex).Value)
                                Me.DgvMachine.Item(Me.ColMachine_Desc.Index, e.RowIndex).Value = Me.objMachine.GetDesc(Me.DgvMachine.Item(Me.ColMachine_PrimaryMachine.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvMachine.Item(Me.ColMachine_MachineCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvMachine)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmMasterMachineDataCari
                            ij.txtCari.Text = Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.Status = "Only Active"
                            ij.ShowDialog()
                            If ij.PrimaryMachineData <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                index = Me.GetIndexItemMachine(ij.PrimaryMachineData, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvMachine.Item(Me.ColMachine_MachineCode.Index, e.RowIndex).Value = Me.objMachine.GetCode(ij.PrimaryMachineData)
                                    Me.DgvMachine.Item(Me.ColMachine_Desc.Index, e.RowIndex).Value = Me.objMachine.GetDesc(ij.PrimaryMachineData)
                                    Me.DgvMachine.Item(Me.ColMachine_PrimaryMachine.Index, e.RowIndex).Value = ij.PrimaryMachineData
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvMachine)
                                        Exit Sub
                                    Else
                                        Me.DgvMachine.Item(Me.ColMachine_MachineCode.Index, e.RowIndex).Value = Me.objMachine.GetCode(ij.PrimaryMachineData)
                                        Me.DgvMachine.Item(Me.ColMachine_Desc.Index, e.RowIndex).Value = Me.objMachine.GetDesc(ij.PrimaryMachineData)
                                        Me.DgvMachine.Item(Me.ColMachine_PrimaryMachine.Index, e.RowIndex).Value = ij.PrimaryMachineData
                                    End If
                                End If
                            ElseIf ij.PrimaryMachineData = -1 Then
                                If Me.DgvMachine.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvMachine)
                    End If
                Else
                    If e.ColumnIndex = Me.ColMachine_Remarks.Index Then
                        Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColMachine_MachineCode.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            If Me.DgvMachine.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        Else
                            Me.DgvMachine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvMachine.Columns(Me.ColMachine_No.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvMachine.Columns(Me.ColMachine_MachineCode.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvMachine.Columns(Me.ColMachine_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvMachine.Columns(Me.ColMachine_Remarks.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub DgvMachine_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvMachine.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.TxtUnitDesc.Enabled = True Then
                rowDGV = Me.DgvMachine.Rows(Me.DgvMachine.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColMachine_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColMachine_MachineCode.Index).Value = ""
                    .Cells(Me.ColMachine_Desc.Index).Value = ""
                    .Cells(Me.ColMachine_Remarks.Index).Value = ""
                    .Cells(Me.ColMachine_PrimaryMachine.Index).Value = "-1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMachine_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvMachine.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvMachine)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMachine_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvMachine.UserDeletedRow
        Try
            'panggil fungsi set number
            Me.ObjCD.SetNumber(Me.DgvMachine)
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
            Me.DgvMachine.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvMachine)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindGroup.Click
        Dim FrmCariGroup As New FrmMasterWorkUnitGroup_Cari
        Dim ObjGroup As New ClsMasterWorkUnitGroup

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryWorkUnitGroup.Text = "-1" And Me.txtGroupCode.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.txtGroupCode.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimaryWorkUnitGroup) > 0 Then
                Me.TxtPrimaryWorkUnitGroup.Text = FrmCariGroup.PrimaryWorkUnitGroup
                Me.txtGroupCode.Text = ObjGroup.GetCode(FrmCariGroup.PrimaryWorkUnitGroup)
                Me.txtGroupDesc.Text = ObjGroup.GetDesc(FrmCariGroup.PrimaryWorkUnitGroup)
                Me.txtCodeLocationIn.Focus() ''Me.btnOK.Focus() -r9
            Else
                Me.TxtPrimaryWorkUnitGroup.Text = -1
                Me.txtGroupCode.Text = ""
                Me.txtGroupDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub txtGroupCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGroupCode.KeyDown
        Dim ObjGroup As New ClsMasterWorkUnitGroup
        Try
            If e.KeyValue = 13 Then
                If ObjGroup.Apakahsudahada(Me.txtGroupCode.Text) = True Then
                    Me.TxtPrimaryWorkUnitGroup.Text = ObjGroup.GetPrimary(Me.txtGroupCode.Text)
                    Me.txtGroupDesc.Text = ObjGroup.GetDesc(Me.TxtPrimaryWorkUnitGroup.Text)
                    Me.txtCodeLocationIn.Focus() ''Me.btnOK.Focus() -r9
                Else
                    Me.btnFindGroup.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.Leave
        Dim ObjGroup As New ClsMasterWorkUnitGroup
        Try
            If Me.txtGroupCode.Text <> "" Then
                If ObjGroup.Apakahsudahada(Me.txtGroupCode.Text) = True Then
                    Me.TxtPrimaryWorkUnitGroup.Text = ObjGroup.GetPrimary(Me.txtGroupCode.Text)
                    Me.txtGroupDesc.Text = ObjGroup.GetDesc(Me.TxtPrimaryWorkUnitGroup.Text)
                Else
                    Me.btnFindGroup.PerformClick()
                End If
            Else
                Me.TxtPrimaryWorkUnitGroup.Text = -1
                Me.txtGroupCode.Text = ""
                Me.txtGroupDesc.Text = ""
            End If
            Me.txtGroupCode.BackColor = Me.TxtPrimaryWorkUnitGroup.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSetuptime_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSetuptime.CellBeginEdit

        Try
            If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
                If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                    Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Else
                    Me.buffer = ""
                End If
            Else
                If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) = False Then
                    If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_CodeNo1.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_ProcessCode2.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_CodeNo2.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_Code.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_SetupTimeDuration.Index Then
                        If Me.ObjCD.Is_Time_String(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = True Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = Me.ObjCD.Convert_ToString(0)
                        End If
                    ElseIf e.ColumnIndex = Me.ColSetupTime_Remarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            Me.buffer = Me.dgvSetuptime.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                        Else
                            Me.buffer = Me.ObjCD.Convert_ToString(0)
                        End If
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

    Private Sub dgvSetuptime_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSetuptime.CellEndEdit
        Dim objPro As New Core_Module.ClsMasterProses
        Dim objCodeNo As New Core_Module.clsMasterProduct
        Dim IJ As New Core_Module.FrmMasterProsesCari
        Dim IK As New Core_Module.frmCariMasterProduct
        Dim IL As New Core_Module.FrmMasterSetupTime_Cari
        Try
            If e.ColumnIndex = Me.ColSetupTime_ProcessCode1.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) Then
                    If objPro.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value) = True Then
                        Dim primary As Double = -1
                        primary = objPro.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value)
                        If Me.ApakahDataMengulang(e.RowIndex, primary, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then

                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value = _
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value.ToString.ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = primary
                        Else
                            'MsgBox("double input", MsgBoxStyle.Information, "Information")
                            MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                End If
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    Else
                        IJ.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value

                        IJ.ShowDialog()
                        If IJ.mPrimary > -1 Then
                            If Me.ApakahDataMengulang(e.RowIndex, IJ.mPrimary, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = IJ.mPrimary
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode1.Index).Value = objPro.GetCode(IJ.mPrimary).ToString.ToUpper
                            Else
                                MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    End If
                                Else
                                    Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                End If
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        If Me.dgvSetuptime.Rows.Count - 1 <> e.RowIndex Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        End If
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = objPro.GetPrimary(Me.buffer)
                    End If
                End If
            End If


            If e.ColumnIndex = Me.ColSetupTime_ProcessCode2.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value) Then
                    If objPro.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value) = True Then
                        Dim primary As Double = -1
                        primary = objPro.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value)

                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                                                 primary, _
                                                 Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                                                 Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then

                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value = objPro.GetCode(primary).ToString.ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = primary
                        Else
                            MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    Else
                        IJ.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value
                        IJ.ShowDialog()
                        If IJ.mPrimary > -1 Then
                            If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                                                                           IJ.mPrimary, _
                                                                           Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                                                                           Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = IJ.mPrimary
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_ProcessCode2.Index).Value = objPro.GetCode(IJ.mPrimary).ToString.ToUpper
                            Else
                                MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                                Else
                                    Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = objPro.GetPrimary(Me.buffer)
                    End If
                End If
            End If






            If e.ColumnIndex = Me.ColSetupTime_CodeNo1.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value) Then
                    If objCodeNo.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value) = True Then
                        Dim primary2 As Double = -1
                        primary2 = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value)

                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                         Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                         primary2, _
                         Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = objCodeNo.GetProductCode(primary2).ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = primary2
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(primary2)
                        Else
                            MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                            Else
                                Dim primary As Double = -1
                                primary = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = primary
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(primary)
                            End If
                        End If

                    Else
                        IK.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value
                        IK.ShowDialog()
                        If IK.PrimaryProduct > -1 Then
                            If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                               IK.PrimaryProduct, _
                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = IK.PrimaryProduct
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = objCodeNo.GetProductCode(IK.PrimaryProduct).ToString.ToUpper
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(IK.PrimaryProduct)
                            Else
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                                Else
                                    Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Dim PrimaryProd As Double = -1
                                    PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = PrimaryProd
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                            End If
                        End If
                    End If

                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo1.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc1.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                    End If
                End If
            End If







            If e.ColumnIndex = Me.ColSetupTime_CodeNo2.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value) Then
                    If objCodeNo.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value.ToString.ToUpper
                        Dim primary As Double = -1
                        primary = objCodeNo.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value)
                        If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                       Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                       Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                       primary) = False Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primary)
                        Else
                            MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                            Else
                                Dim primaryDet As Double = -1
                                primaryDet = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primaryDet
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primaryDet)
                            End If
                        End If
                    Else

                        IK.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value
                        IK.ShowDialog()
                        If IK.PrimaryProduct > -1 Then
                            If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
                                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
                                               Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
                                               IK.PrimaryProduct) = False Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = IK.PrimaryProduct
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = objCodeNo.GetProductCode(IK.PrimaryProduct).ToString.ToUpper
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(IK.PrimaryProduct)
                            Else
                                MsgBox("This work unit data has double code no, please change one of product code", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                                Else
                                    Dim primaryDet As Double = -1
                                    primaryDet = objCodeNo.GetPrimary(Me.buffer)
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = primaryDet
                                    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(primaryDet)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objCodeNo.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_CodeNo2.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Desc2.Index).Value = objCodeNo.GetDescription(PrimaryProd)
                    End If
                End If
            End If




            If e.ColumnIndex = Me.ColSetupTime_Code.Index Then
                If Not String.IsNullOrEmpty(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value) Then
                    If Me.objSetup.ApakahSudahAda(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value.ToString.ToUpper
                        Dim primary As Double = -1
                        primary = objSetup.GetPrimary(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = primary
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(primary))
                    Else
                        IL.txtCari.Text = Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value
                        IL.ShowDialog()
                        If IL.Primary > -1 Then
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = IL.Primary
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.objSetup.GetCode(IL.Primary).ToString.ToUpper
                            Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(IL.Primary))
                        Else
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = ""
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = ""

                            Else
                                Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Dim PrimaryProd As Double = -1
                                PrimaryProd = objSetup.GetPrimary(Me.buffer)
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = PrimaryProd
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.buffer
                                Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(PrimaryProd))
                            End If
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = -1
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = ""
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = ""

                    Else
                        Me.dgvSetuptime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Dim PrimaryProd As Double = -1
                        PrimaryProd = objSetup.GetPrimary(Me.buffer)
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = PrimaryProd
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_Code.Index).Value = Me.buffer
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(objSetup.GetDuration(PrimaryProd))
                    End If
                End If
            End If

            If e.ColumnIndex = Me.ColSetupTime_SetupTimeDuration.Index Then
                If Me.ObjCD.Is_Time_String(Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value) = False Then
                    MsgBox("setup time only can fill by time data, please correct your data", MsgBoxStyle.Information, "Information")
                    If Me.ObjCD.Is_Time_String(Me.buffer) = True Then
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.buffer
                    Else
                        Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(0)
                    End If
                End If
            End If


            'If Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value <> "-1" AndAlso _
            ' Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value <> "-1" Then

            '    If Me.ApakahDataMengulang(e.RowIndex, Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryProses2.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value, _
            '    Me.dgvSetuptime.Rows(e.RowIndex).Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value) = True Then
            '        MsgBox("Double input ", MsgBoxStyle.Information, "information")

            '    End If
            'End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objPro = Nothing
            objCodeNo = Nothing
            IJ.Dispose()
            IK.Dispose()
            IL.Dispose()
        End Try
    End Sub

    Public Function ApakahDataMengulang(ByVal vIndex As Integer, ByVal vPrimaryProses1 As Double, ByVal vPrimaryProses2 As Double, _
                                        ByVal vPrimaryProduct1 As Double, ByVal vprimaryProduct2 As Double) As Boolean
        Dim rowDGV As New DataGridViewRow

        Try
            If Me.dgvSetuptime.RowCount > 1 Then
                For intA As Integer = 0 To Me.dgvSetuptime.Rows.Count - 2
                    rowDGV = Me.dgvSetuptime.Rows(intA)
                    If intA <> vIndex Then
                        If vPrimaryProses1 = rowDGV.Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value AndAlso _
                        vPrimaryProses2 = rowDGV.Cells(Me.ColSetupTime_PrimaryProses2.Index).Value AndAlso _
                        vPrimaryProduct1 = rowDGV.Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value AndAlso _
                        vprimaryProduct2 = rowDGV.Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value Then
                            Return True
                        End If
                    End If
                Next
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Sub dgvSetuptime_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSetuptime.RowsAdded

        Dim rowDGV As DataGridViewRow
        Try
            If Me.TxtUnitDesc.Enabled = True Then
                rowDGV = Me.dgvSetuptime.Rows(Me.dgvSetuptime.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColSetupTime_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColSetupTime_PrimaryCodeNo1.Index).Value = "-1"
                    .Cells(Me.ColSetupTime_PrimaryCodeNo2.Index).Value = "-1"
                    .Cells(Me.ColSetupTime_PrimaryProcess1.Index).Value = "-1"
                    .Cells(Me.ColSetupTime_PrimaryProses2.Index).Value = "-1"
                    .Cells(Me.ColSetupTime_PrimarySetupTime.Index).Value = "-1"
                    .Cells(Me.ColSetupTime_ProcessCode1.Index).Value = ""
                    .Cells(Me.ColSetupTime_ProcessCode2.Index).Value = ""
                    .Cells(Me.ColSetupTime_CodeNo1.Index).Value = ""
                    .Cells(Me.ColSetupTime_CodeNo2.Index).Value = ""
                    .Cells(Me.ColSetupTime_Desc1.Index).Value = ""
                    .Cells(Me.ColSetupTime_Desc2.Index).Value = ""
                    .Cells(Me.ColSetupTime_Code.Index).Value = ""
                    .Cells(Me.ColSetupTime_SetupTimeDuration.Index).Value = Me.ObjCD.Convert_ToString(0)
                    .Cells(Me.ColSetupTime_Remarks.Index).Value = ""
                End With
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub dgvSetuptime_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSetuptime.Sorted

        Try
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub dgvSetuptime_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvSetuptime.UserDeletedRow

        Try
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.dgvSetuptime.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvSetuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '|------------------------------------------|
    '| Ahdar - Penambahaan Find Location IN&OUT |
    '|------------------------------------------|                                          
    Private Sub BtnFindLocationIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLocationIn.Click
        Dim FrmCari As New FrmCariMasterStorageArea
        Dim ObjS As New clsMasterStorageArea

        Try
            FrmCari.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryLocationIn.Text = "-1" And Me.txtCodeLocationIn.Text = "" Then
                FrmCari.ShowDialog()
            Else
                FrmCari.txtCari.Text = Me.txtCodeLocationIn.Text
                FrmCari.ShowDialog()
            End If
            If CDbl(FrmCari.PrimaryStorageArea) > 0 Then
                Me.txtPrimaryLocationIn.Text = FrmCari.PrimaryStorageArea
                Me.txtCodeLocationIn.Text = ObjS.GetStorageAreaCode(FrmCari.PrimaryStorageArea)
                Me.txtDescLocationIn.Text = ObjS.GetDescription(FrmCari.PrimaryStorageArea)
                Me.txtCodeLocationOut.Focus()
            Else
                Me.txtPrimaryLocationIn.Text = -1
                Me.txtCodeLocationIn.Text = ""
                Me.txtDescLocationIn.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCari.Dispose()
        End Try
    End Sub

    Private Sub txtCodeLocationIn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeLocationIn.KeyDown
        Dim ObjS As New clsMasterStorageArea
        Try
            If e.KeyValue = 13 Then
                If ObjS.ApakahSudahAda(Me.txtCodeLocationIn.Text) = True Then
                    Me.txtPrimaryLocationIn.Text = ObjS.GetPrimary(Me.txtCodeLocationIn.Text)
                    Me.txtDescLocationIn.Text = ObjS.GetDescription(Me.txtPrimaryLocationIn.Text)
                    Me.txtCodeLocationOut.Focus() ''Me.btnOK.Focus() -r9
                Else
                    Me.BtnFindLocationIn.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeLocationIn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeLocationIn.Leave
        Dim ObjS As New clsMasterStorageArea
        Try
            If Me.txtCodeLocationIn.Text <> "" Then
                If ObjS.ApakahSudahAda(Me.txtCodeLocationIn.Text) = True Then
                    Me.txtPrimaryLocationIn.Text = ObjS.GetPrimary(Me.txtCodeLocationIn.Text)
                    Me.txtDescLocationIn.Text = ObjS.GetDescription(Me.txtPrimaryLocationIn.Text)
                Else
                    Me.BtnFindLocationIn.PerformClick()
                End If
            Else
                Me.txtPrimaryLocationIn.Text = -1
                Me.txtCodeLocationIn.Text = ""
                Me.txtDescLocationIn.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindLocationOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLocationOut.Click
        Dim FrmCari As New FrmCariMasterStorageArea
        Dim ObjS As New clsMasterStorageArea

        Try
            FrmCari.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryLocationOut.Text = "-1" And Me.txtCodeLocationOut.Text = "" Then
                FrmCari.ShowDialog()
            Else
                FrmCari.txtCari.Text = Me.txtCodeLocationOut.Text
                FrmCari.ShowDialog()
            End If
            If CDbl(FrmCari.PrimaryStorageArea) > 0 Then
                Me.txtPrimaryLocationOut.Text = FrmCari.PrimaryStorageArea
                Me.txtCodeLocationOut.Text = ObjS.GetStorageAreaCode(FrmCari.PrimaryStorageArea)
                Me.txtDescLocationOut.Text = ObjS.GetDescription(FrmCari.PrimaryStorageArea)
                Me.TxtProductionLineCode.Focus()
            Else
                Me.txtPrimaryLocationOut.Text = -1
                Me.txtCodeLocationOut.Text = ""
                Me.txtDescLocationOut.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCari.Dispose()
        End Try
    End Sub

    Private Sub txtCodeLocationOut_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeLocationOut.KeyDown
        Dim ObjS As New clsMasterStorageArea
        Try
            If e.KeyValue = 13 Then
                If ObjS.ApakahSudahAda(Me.txtCodeLocationOut.Text) = True Then
                    Me.txtPrimaryLocationOut.Text = ObjS.GetPrimary(Me.txtCodeLocationOut.Text)
                    Me.txtDescLocationOut.Text = ObjS.GetDescription(Me.txtPrimaryLocationOut.Text)
                    Me.TxtProductionLineCode.Focus()
                Else
                    Me.BtnFindLocationOut.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeLocationOut_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeLocationOut.Leave
        Dim ObjS As New clsMasterStorageArea
        Try
            If Me.txtCodeLocationOut.Text <> "" Then
                If ObjS.ApakahSudahAda(Me.txtCodeLocationOut.Text) = True Then
                    Me.txtPrimaryLocationOut.Text = ObjS.GetPrimary(Me.txtCodeLocationOut.Text)
                    Me.txtDescLocationOut.Text = ObjS.GetDescription(Me.txtPrimaryLocationOut.Text)
                Else
                    Me.BtnFindLocationOut.PerformClick()
                End If
            Else
                Me.txtPrimaryLocationOut.Text = -1
                Me.txtCodeLocationOut.Text = ""
                Me.txtDescLocationOut.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindProductionLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindProductionLine.Click
        Dim objLine As New Core_Module.ClsMasterMachineLine
        Dim frmFind As New Core_Module.FrmMasterMachineLine_Cari
        Try
            Me.TxtProductionLineCode.Text = Me.TxtProductionLineCode.Text.Trim
            frmFind.txtCari.Text = Me.TxtProductionLineCode.Text
            frmFind.BackColorAll = Me.BackColorAll
            frmFind.ShowDialog()
            If frmFind.Primary > 0 Then
                Me.TxtPrimaryProductionLine.Text = frmFind.Primary
                Me.TxtProductionLineCode.Text = objLine.GetCode(frmFind.Primary)
                Me.TxtProductionLineDesc.Text = objLine.GetDescription(frmFind.Primary)
                Me.ChkActive.Focus()
            Else
                Me.TxtPrimaryProductionLine.Text = -1
                Me.TxtProductionLineCode.Text = String.Empty
                Me.TxtProductionLineDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtProductionLineCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductionLineCode.GotFocus
        Try
            If Me.btnOK.Enabled = True Then
                If TxtProductionLineCode.Enabled = True Then
                    If TxtProductionLineCode.ReadOnly = True Then
                        TxtProductionLineCode.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        TxtProductionLineCode.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtProductionLineCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductionLineCode.KeyDown
        Dim objLine As New Core_Module.ClsMasterMachineLine
        Try
            If e.KeyCode = 13 Then
                Me.TxtProductionLineCode.Text = Me.TxtProductionLineCode.Text.Trim
                If objLine.ApakahSudahAda(Me.TxtProductionLineCode.Text) Then
                    Me.TxtPrimaryProductionLine.Text = objLine.GetPrimary(Me.TxtProductionLineCode.Text)
                    Me.TxtProductionLineDesc.Text = objLine.GetDescription(Me.TxtPrimaryProductionLine.Text)
                    Me.ChkActive.Focus()
                Else
                    Me.BtnFindProductionLine.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtProductionLineCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductionLineCode.Leave
        Dim objLine As New Core_Module.ClsMasterMachineLine
        Try
            Me.TxtProductionLineCode.Text = Me.TxtProductionLineCode.Text.Trim
            If String.IsNullOrEmpty(Me.TxtProductionLineCode.Text) = False Then
                If objLine.ApakahSudahAda(Me.TxtProductionLineCode.Text) Then
                    Me.TxtPrimaryProductionLine.Text = objLine.GetPrimary(Me.TxtProductionLineCode.Text)
                    Me.TxtProductionLineDesc.Text = objLine.GetDescription(Me.TxtPrimaryProductionLine.Text)
                Else
                    Me.BtnFindProductionLine.PerformClick()
                End If
            Else
                Me.TxtPrimaryProductionLine.Text = -1
                Me.TxtProductionLineCode.Text = String.Empty
                Me.TxtProductionLineDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtProductionLineCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductionLineCode.LostFocus
        Try
            If Me.btnOK.Enabled = True Then
                If TxtProductionLineCode.Enabled = True Then
                    If TxtProductionLineCode.ReadOnly = True Then
                        TxtProductionLineCode.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        TxtProductionLineCode.BackColor = Me.TxtPrimaryUnit.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class