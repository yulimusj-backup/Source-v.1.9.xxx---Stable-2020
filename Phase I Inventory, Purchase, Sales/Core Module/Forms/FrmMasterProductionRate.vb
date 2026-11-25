Option Explicit On
Option Strict Off
Public Class FrmMasterProductionRate
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim obj As New ClsMasterProductionRate
    Dim objCd As New Core_Module.ClsCekData
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim ObjWorkUnit As New ClsMasterWorkUnit
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim baru As Boolean = True
    Dim buffer As String = ""
    Dim mEraseRowIndex As Integer = 0
    Dim EraseRowIndex As Integer = 0
    Dim mIndex As Integer = -1
    Dim DummyIndex As Integer = -1
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    Private Sub AktifkanDetail()
        Try
            Me.objCd.SetBackColor(Me.GboDetail.Controls, Me.txtPrimary.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.GboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.DgvWorkUnit.ReadOnly = False
            Me.dgvItemProductGroup.ReadOnly = False
            Me.DgvWorkUnit.AllowUserToAddRows = True
            Me.dgvItemProductGroup.AllowUserToAddRows = True
            Me.DgvWorkUnit.AllowUserToDeleteRows = True
            Me.dgvItemProductGroup.AllowUserToDeleteRows = True
            Me.CancelButton = Me.BtnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.cboDependency.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.GboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.DgvWorkUnit.ReadOnly = True
            Me.dgvItemProductGroup.ReadOnly = True
            Me.DgvWorkUnit.AllowUserToAddRows = False
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            Me.DgvWorkUnit.AllowUserToDeleteRows = False
            Me.dgvItemProductGroup.AllowUserToDeleteRows = False
            Me.CancelButton = Me.BtnExit
            Me.objCd.SetBackColor(Me.GboDetail.Controls, Me.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.cboDependency.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.GboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.DgvWorkUnit.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.txtPrimary.Text = "-1"
            Me.txtPrimaryProcess.Text = "-1"
            Me.TslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.TslModifiedDate.Text = "Modified Date: "
            Me.txtDependency.Text = "00.00"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            KoleksiNamaData.Add("Process")
            KoleksiTextBox.Add(Me.txtProcessCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.baru = True Then
                If Me.obj.ApakahSudahAda(True, Me.TxtDocumentNo.Text) = True Then
                    Me.TxtDocumentNo.Text = Me.obj.getNoBaru(Me.dtpDate.Value)
                End If
            Else
                Dim mDocNo As String = Me.obj.GetDocumentNo(Me.txtPrimary.Text)
                If mDocNo <> Me.TxtDocumentNo.Text Then
                    If Me.obj.ApakahSudahAda(False, TxtDocumentNo.Text) = True Then
                        Me.TxtDocumentNo.Text = Me.obj.getNoBaru(Me.dtpDate.Value)
                    End If
                End If
                Dim OldTimeStamp As Date = Me.obj.GetTimestamp(Me.txtPrimary.Text)
                If Me.obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.obj.Primary, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.ApakahDataItemBenar() = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Dim row_ As DataGridViewRow
        Try
            If Me.DgvWorkUnit.AllowUserToAddRows = True Then
                If Me.DgvWorkUnit.RowCount <= 1 Then
                    Me.TabControl1.SelectedTab = Me.TabPage1
                    row_ = Me.DgvWorkUnit.Rows(0)
                    Me.DgvWorkUnit.CurrentCell = row_.Cells(Me.ColWorkUnit_Code.Index)
                    MsgBox("Input item work unit at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.DgvWorkUnit.RowCount <= 0 Then
                    Me.TabControl1.SelectedTab = Me.TabPage1
                    row_ = Me.DgvWorkUnit.Rows(0)
                    Me.DgvWorkUnit.CurrentCell = row_.Cells(Me.ColWorkUnit_Code.Index)
                    MsgBox("Input item work unit at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.dgvItemProductGroup.AllowUserToAddRows = True Then
                If Me.dgvItemProductGroup.RowCount <= 1 Then
                    Me.TabControl1.SelectedTab = Me.TabPage2
                    row_ = Me.dgvItemProductGroup.Rows(0)
                    Me.dgvItemProductGroup.CurrentCell = row_.Cells(Me.ColProductGroup_CodeNo.Index)
                    MsgBox("Input item product at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If Me.dgvItemProductGroup.RowCount <= 0 Then
                    Me.TabControl1.SelectedTab = Me.TabPage2
                    row_ = Me.dgvItemProductGroup.Rows(0)
                    Me.dgvItemProductGroup.CurrentCell = row_.Cells(Me.ColProductGroup_CodeNo.Index)
                    MsgBox("Input item product at least 1 item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Dim level As Double = -1
            For Each row As DataGridViewRow In Me.dgvItemProductGroup.Rows
                If IsNothing(row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value) = False Then
                    If String.IsNullOrEmpty(row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString) = False Then
                        If row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> "-1" Then
                            Select Case row.Cells(Me.ColProductGroup_Level.Index).Value
                                Case "Product Group"
                                    level = 0
                                Case "Brand"
                                    level = 1
                                Case "Sub Brand"
                                    level = 2
                                Case "Category"
                                    level = 3
                                Case "Product"
                                    level = 4
                            End Select
                            If Me.obj.IsExistProcessAndProduct(Me.txtPrimaryProcess.Text, level, row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value, Me.txtPrimary.Text) Then
                                Me.TabControl1.SelectedTab = Me.TabPage2
                                Me.dgvItemProductGroup.CurrentCell = row.Cells(Me.ColProductGroup_CodeNo.Index)
                                MsgBox(row.Cells(Me.ColProductGroup_Level.Index).Value & ": " & row.Cells(Me.ColProductGroup_CodeNo.Index).Value & " with process: " & Me.txtProcessCode.Text & " already exist in other no document.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.obj
                .Primary = Me.txtPrimary.Text
                .DocumentNo = Me.TxtDocumentNo.Text
                ._Date = Me.dtpDate.Value.Date
                .PrimaryProcess = Me.txtPrimaryProcess.Text
                .Remarks = Me.txtRemarks.Text
                .TipeDependency = Me.cboDependency.SelectedIndex
                .DelayDependency = Me.objCd.Convert_ToMinutes(Me.txtDependency.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.obj
                Me.txtPrimary.Text = .Primary
                Me.TxtDocumentNo.Text = .DocumentNo
                Me.dtpDate.Value = ._Date
                Me.txtPrimaryProcess.Text = .PrimaryProcess
                Me.txtProcessCode.Text = Me.objProcess.GetCode(.PrimaryProcess)
                Me.txtProcessDesc.Text = Me.objProcess.GetDesc(.PrimaryProcess)
                Me.txtRemarks.Text = .Remarks
                Me.cboDependency.SelectedIndex = .TipeDependency
                'If .TipeDependency = clsMasterProduct.EnumTipeDependency.End_Start Then
                '    Me.txtDependency.ReadOnly = True
                '    Me.txtDependency.BackColor = Me.txtReadOnly.BackColor
                'Else
                '    Me.txtDependency.ReadOnly = False
                '    Me.txtDependency.BackColor = Me.txtPrimary.BackColor
                'End If
                Me.txtDependency.Text = Me.objCd.Convert_ToString(.DelayDependency)
                Me.TslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.TslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.TslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.TslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKota_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtRemarks.Enabled = True And (Me.txtPrimaryProcess.Text <> "-1" Or Me.DgvWorkUnit.RowCount > 1 Or Me.dgvItemProductGroup.RowCount > 1) Then
                Dim hasil As MsgBoxResult = MsgBox("Save data before closing?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanBaru(Me.UserCode, Me) = True Then
                                MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.BtnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.obj.SimpanPerubahan(Me.UserCode, Me) = True Then
                                MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.BtnEdit.Focus()
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

    Private Sub FrmMasterKota_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.txtDependency.ReadOnly = False
            Me.txtDependency.BackColor = Me.txtPrimary.BackColor

            Me.DgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.SplitContainer1.TabStop = False
            Me.dgvItemProductGroup.ColumnHeadersHeight = 25
            Me.DgvWorkUnit.ColumnHeadersHeight = 25
            Me.dgvItemProductGroup.Rows.Clear()

            'dependency
            Me.cboDependency.Items.Clear()
            Me.cboDependency.Items.Add("End-Start")
            Me.cboDependency.Items.Add("Start-Start")
            Me.cboDependency.Items.Add("Start-Start (Strict)")
            Me.cboDependency.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanBaru(Me.UserCode, Me) Then
                            MsgBox("Data Saved.", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.BtnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanPerubahan(Me.UserCode, Me) Then
                            MsgBox("Data change saved.", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.BtnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtRemarks.Enabled = True And (Me.txtPrimaryProcess.Text <> "-1" Or Me.DgvWorkUnit.RowCount > 1 Or Me.dgvItemProductGroup.RowCount > 1) Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        If obj.CariDetail(Me.txtPrimary.Text, Me) = True Then
                            Me.SalinKeField()
                        End If
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        Try
            Me.txtPrimary.Text = "-1"
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtDocumentNo.Text = Me.obj.getNoBaru(Me.dtpDate.Value)
            Me.txtProcessCode.Focus()
            Me.TslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.TslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.TslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFind.Click
        Dim hj As New FrmMasterProductionRate_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                If Me.obj.CariDetail(hj.Primary, Me) = True Then
                    Me.SalinKeField()
                    Me.BtnEdit.Focus()
                Else
                    Me.BtnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.BtnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.obj.ApakahBolehEdit(True, Me.UserCode, Me.txtPrimary.Text, True) = False Then
                    Me.BtnFind.Focus()
                Else
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.TxtDocumentNo.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.BtnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.obj.ApakahBolehHapus(True, Me.UserCode, Me.txtPrimary.Text, True) = False Then
                    Me.BtnFind.Focus()
                Else
                    hasil = MsgBox("Are you sure to delete this data?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.obj.Hapus(Me.txtPrimary.Text) = True Then
                            Me.KosongkanField()
                            Me.BtnNew.Focus()
                            MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub txtKode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.GotFocus, txtRemarks.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False Then
                txt.BackColor = BackColorAll
            Else
                txt.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.Leave
        Try
            txtRemarks.Text = txtRemarks.Text.Trim.Replace("'", "").Replace("*", "").Replace("%", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtKode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProcessCode.LostFocus, txtRemarks.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            If txt.ReadOnly = False And txt.Enabled = True Then
                txt.BackColor = Me.txtPrimary.BackColor
            Else
                txt.BackColor = Me.txtReadOnly.BackColor
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.obj._Date.Year = Me.dtpDate.Value.Year And Me.obj._Date.Month = Me.dtpDate.Value.Month Then
                        Me.TxtDocumentNo.Text = Me.obj.DocumentNo
                    Else
                        Me.TxtDocumentNo.Text = Me.obj.getNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtDocumentNo.Text = Me.obj.getNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWorkUnit.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColWorkUnit_Code.Index Or e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index _
                Or e.ColumnIndex = Me.ColWorkUnit_Remarks.Index) Or e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                    Me.buffer = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                    Me.buffer = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWorkUnit.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvWorkUnit.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            'kemungkinan
                            If Me.ObjWorkUnit.ApakahSudahAdaOnlyActive(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Primary = Me.ObjWorkUnit.GetPrimary(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemWorkUnit(Primary, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetPrimary(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value)
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                        Else
                                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                                ij.txtCari.Text = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.Status = "Only Active"
                                ij.ShowDialog()
                                If ij.PrimaryWorkUnit <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    index = Me.GetIndexItemWorkUnit(ij.PrimaryWorkUnit, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                                Me.EraseRowIndex = e.RowIndex
                                                Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                            Else
                                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                                            Exit Sub
                                        Else
                                            Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                            Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                            Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                        End If
                                    End If
                                ElseIf ij.PrimaryWorkUnit = -1 Then
                                    If Me.DgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                        Else
                                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                    End If
                                End If
                                ij = Nothing
                            End If
                        ElseIf e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index Then
                            'digit grouping
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.objCd.DigitGrouping(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                            'harus bilangan positif
                            If Me.objCd.CekBilanganPositif(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'maximal nilainya 100 %
                            If CDec(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                            Me.DgvWorkUnit.Item(Me.ColWorkUnit_EffectiveRate.Index, e.RowIndex).Value = Me.objCd.BuangNolDibelakangDecimal(Me.objCd.DigitGrouping(Me.DgvWorkUnit.Item(Me.ColWorkUnit_Efficiency.Index, e.RowIndex).Value * Me.DgvWorkUnit.Item(Me.ColWorkUnit_Rate.Index, e.RowIndex).Value / 100, 6))
                        ElseIf e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                            If IsNumeric(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'digit grouping
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.objCd.BuangNolDibelakangDecimal(Me.objCd.DigitGrouping(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value, 6))
                            'harus bilangan positif
                            If Me.objCd.CekBilanganPositif(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.DgvWorkUnit.Item(Me.ColWorkUnit_EffectiveRate.Index, e.RowIndex).Value = Me.objCd.BuangNolDibelakangDecimal(Me.objCd.DigitGrouping(Me.DgvWorkUnit.Item(Me.ColWorkUnit_Efficiency.Index, e.RowIndex).Value * Me.DgvWorkUnit.Item(Me.ColWorkUnit_Rate.Index, e.RowIndex).Value / 100, 6))
                        ElseIf e.ColumnIndex = Me.ColWorkUnit_Remarks.Index Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim().Replace("'", "").Replace("%", "").Replace("*", "")
                        End If
                    Else
                        If e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index Or e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        ElseIf e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                If Me.DgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                End If
                            Else
                                Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        Else
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index Or e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    ElseIf e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            If Me.DgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                            End If
                        Else
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Function GetIndexItemWorkUnit(ByVal vPrimaryWorkUnit As Double, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvWorkUnit.Rows.Count - 2
                RowDGV = Me.DgvWorkUnit.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColWorkUnit_PrimaryWorkUnit.Index).Value = vPrimaryWorkUnit Then
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

    Private Sub DgvWorkUnit_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvWorkUnit.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.BtnOk.Enabled = True Then
                rowDGV = Me.DgvWorkUnit.Rows(Me.DgvWorkUnit.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColWorkUnit_Priority.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColWorkUnit_Code.Index).Value = ""
                    .Cells(Me.ColWorkUnit_Efficiency.Index).Value = "100.00"
                    .Cells(Me.ColWorkUnit_Rate.Index).Value = "1"
                    .Cells(Me.ColWorkUnit_Remarks.Index).Value = ""
                    .Cells(Me.ColWorkUnit_PrimaryWorkUnit.Index).Value = "-1"
                    .Cells(Me.ColWorkUnit_EffectiveRate.Index).Value = "1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvWorkUnit.Sorted
        Try
            Me.objCd.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvWorkUnit.UserDeletedRow
        Try
            Me.objCd.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelDgvWorkUnit_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDelDgvWorkUnit.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelDgvWorkUnit_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDelDgvWorkUnit.RunWorkerCompleted
        Try
            Me.DgvWorkUnit.Rows.RemoveAt(Me.EraseRowIndex)
            Me.objCd.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProcessCode.Leave
        Try
            If Me.txtProcessCode.Text.Trim.Length > 0 Then
                If Me.objProcess.ApakahSudahAda(Me.txtProcessCode.Text) Then
                    Me.txtPrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txtProcessCode.Text)
                    Me.txtProcessDesc.Text = Me.objProcess.GetDesc(Me.txtPrimaryProcess.Text)
                Else
                    Me.btnFindProcess.PerformClick()
                End If
            Else
                Me.txtProcessDesc.Clear()
                Me.txtPrimaryProcess.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProcessCode.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If Me.objProcess.ApakahSudahAda(Me.txtProcessCode.Text) Then
                    Me.txtPrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txtProcessCode.Text)
                    Me.txtProcessDesc.Text = Me.objProcess.GetDesc(Me.txtPrimaryProcess.Text)
                    Me.txtRemarks.Focus()
                Else
                    Me.btnFindProcess.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProcess.Click
        Dim frm As New FrmMasterProsesCari
        Try
            frm.txtCari.Text = Me.txtProcessCode.Text
            frm.ShowDialog()
            If frm.mPrimary > -1 Then
                Me.txtPrimaryProcess.Text = frm.mPrimary
                Me.txtProcessCode.Text = Me.objProcess.GetCode(frm.mPrimary)
                Me.txtProcessDesc.Text = Me.objProcess.GetDesc(frm.mPrimary)
                Me.txtRemarks.Focus()
            Else
                Me.txtProcessCode.Clear()
                Me.txtProcessDesc.Clear()
                Me.txtPrimaryProcess.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        Dim f As New FrmMasterProductionRate_Apply()
        Try
            If Me.txtPrimary.Text = "-1" Or Me.TxtDocumentNo.Text = "" Then
                Me.BtnFind.PerformClick()
            End If
            If Me.txtPrimary.Text <> "-1" And Me.TxtDocumentNo.Text <> "" Then
                If MsgBox("Are you sure to apply this rate settings?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                f.TipeDependency = Me.cboDependency.SelectedIndex
                f.DelayDependency = Me.objCd.Convert_ToMinutes(Me.txtDependency.Text)
                f.mProductionRate = Me.txtPrimary.Text
                f.mPrimaryProcess = Me.txtPrimaryProcess.Text
                f.UserCode = Me.UserCode
                f.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As ClsMasterProductionRate.ProductLevel = ClsMasterProductionRate.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.obj.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = ClsMasterProductionRate.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = ClsMasterProductionRate.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = ClsMasterProductionRate.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = ClsMasterProductionRate.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.obj.GetProductLevelStr(ClsMasterProductionRate.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGV_ItemGroup()
        Dim ObjSettingPG As New Core_Module.ClsTransaksiSetting_ProdGroup

        Try
            ObjSettingPG.IsikanDGV(Me.UserCode, Core_Module.ClsTransaksiSetting_ProdGroup.EnumTipeTransaksi.LaporanPurchaseRequest, Me.dgvItemProductGroup, _
            Me.ColProductGroup_No.Index, Me.ColProductGroup_Level.Index, Me.ColProductGroup_CodeNo.Index, Me.ColProductGroup_Desc.Index, _
            Me.ColProductGroup_Remarks.Index, Me.ColProductGroup_PrimaryAcuan.Index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSettingPG = Nothing
        End Try
    End Sub

    Private Sub cboDependency_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDependency.SelectedIndexChanged
        'Try
        '    If Me.cboDependency.SelectedIndex = clsMasterProduct.EnumTipeDependency.End_Start Then
        '        Me.txtDependency.Text = Me.ObjCD.Convert_ToString(0)
        '        Me.txtDependency.ReadOnly = True
        '        Me.txtDependency.BackColor = Me.TxtReadOnly.BackColor
        '    Else
        '        Me.txtDependency.ReadOnly = False
        '        If Me.txtDependency.Enabled = True Then
        '            Me.txtDependency.BackColor = Me.txtPrimary.BackColor
        '        Else
        '            Me.txtDependency.BackColor = Me.TxtReadOnly.BackColor
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub txtDependency_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDependency.Enter
        Try
            Me.buffer = Me.txtDependency.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.GotFocus
        Try
            If Me.txtDependency.Enabled = True Then
                If Me.txtDependency.ReadOnly = True Then
                    Me.txtDependency.BackColor = Me.txtDependency.BackColor
                Else
                    Me.txtDependency.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.Leave
        Try
            If Me.ObjCD.Is_Time_String(Me.txtDependency.Text) = False Then
                MsgBox("Please input valid format (HH.mm).", MsgBoxStyle.Information, "Information")
                Me.txtDependency.Text = Me.buffer
                Me.txtDependency.Focus()
                Exit Sub
            End If
            If Me.ObjCD.Convert_ToMinutes(Me.txtDependency.Text) < 0 Then
                MsgBox("Delay dependency duration cannot negative value.", MsgBoxStyle.Information, "Information")
                Me.txtDependency.Text = Me.buffer
                Me.txtDependency.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.LostFocus
        Try
            If Me.txtDependency.Enabled = True Then
                If Me.txtDependency.ReadOnly = True Then
                    Me.txtDependency.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.txtDependency.BackColor = Me.txtPrimary.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class