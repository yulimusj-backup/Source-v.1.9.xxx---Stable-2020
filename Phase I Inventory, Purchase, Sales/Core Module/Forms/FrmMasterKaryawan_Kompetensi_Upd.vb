'@hdar
Option Explicit On
Option Strict Off
Imports System.Threading
Imports System.Globalization
Public Class FrmMasterKaryawan_Kompetensi_Upd
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public LocationFileSumber As String = ""
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim IndexSalFor As Integer = 0
    Dim Buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjSalFor As New Core_Module.ClsMasterSalaryFormula
    Dim ObjKomp As New Core_Module.ClsMasterKompetensi
    Dim DefPrec As Decimal = 0
    Dim TBLEmp As New DseMasterKaryawan_Kompetensi_Upd.DataTable1DataTable
    Dim TBLEmp2 As New DseMasterKaryawan_Kompetensi_Upd.DataTable2DataTable
    Dim TBLRec As New DseMasterKaryawan_Kompetensi_Upd.GRecordDataTable

    Private Sub LlbAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbAdvance.LinkClicked
        Me.ContextMenuStrip1.Show(Me.LlbAdvance, 0, 20)
        Me.TxtPrimary.Visible = False
    End Sub

    Private Sub ApplyToAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyToAllToolStripMenuItem.Click
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim i As New FrmMasterKaryawan_Kompetensi_Upd_Opt
        Dim StrValue As String = ""
        Dim ICount As Integer = 0
        Dim DataBenar As Boolean = False
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(Me.TxtCode.Text) = True Then
                MsgBox("Code cannot empty string", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Select()
                Exit Sub
            End If

            If Me.DgvItem.Rows.Count <= 0 Then
                MsgBox("Data item not found", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                MsgBox("Data item with competence code invalid please reload again", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtCode.Text <> Me.ObjKomp.GetDataCode(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed code at master competence, please reload code competence", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.cboDataType.SelectedIndex <> Me.ObjKomp.GetDataType(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed type value at master competence, please reload again", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 1
                If Me.DgvItem.Item(Me.ColItem_Modify.Index, ij).Value = True Then
                    DataBenar = True
                    Exit For
                End If
            Next

            If DataBenar = False Then
                MsgBox("No data selected", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.DataType = Me.cboDataType.SelectedIndex
            i.DefPrec = Me.DefPrec
            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                For Each rowDGV As DataGridViewRow In Me.DgvItem.Rows
                    With rowDGV
                        Dim NextCertificateD As Date = #1/1/1950#

                        If i.DataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            StrValue = i.cboAlfabetic.SelectedItem.ToString
                        ElseIf i.DataType = ClsMasterKompetensi.EnumDataType.Number Then
                            StrValue = i.TxtDefValueDigit.Text
                        ElseIf i.DataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            StrValue = Format(i.DtpDefValue.Value.Date, "dd MMM yyy")
                        ElseIf i.DataType = ClsMasterKompetensi.EnumDataType.Text Then
                            StrValue = i.TxtDefValueText.Text
                        ElseIf i.DataType = ClsMasterKompetensi.EnumDataType.YN Then
                            StrValue = i.cboDefValue.SelectedItem.ToString
                        End If
                        If i.DtpNextCertDate.Enabled = True Then
                            NextCertificateD = i.DtpNextCertDate.Value.Date
                        End If
                        If .Cells(Me.ColItem_Modify.Index).Value = True Then
                            'new 
                            If String.IsNullOrEmpty(.Cells(Me.ColItem_Value.Index).Value) = True Then
                                DTA.InsertQuery(.Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text, StrValue, i.txtRemarks.Text, Format(NextCertificateD, "dd MMM yyy"))
                                DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, .Cells(Me.ColItem_PrimaryEmp.Index).Value)
                            Else
                                'update
                                DTA.UpdateQuery(StrValue, i.txtRemarks.Text, Format(NextCertificateD, "dd MMM yyy"), .Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text)
                                DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, .Cells(Me.ColItem_PrimaryEmp.Index).Value)
                            End If
                            ICount += 1
                        End If
                    End With
                Next
                If ICount > 0 Then
                    Me.IsikanData()
                    MsgBox("Successed", MsgBoxStyle.Information, "Information")
                    Me.LlbChkAll_Click(sender, e)
                    Me.LlbChkAllExp_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DeleteToAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToAllToolStripMenuItem.Click
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim Result As MsgBoxResult
        Dim DataBenar As Boolean = False
        Dim ICount As Integer = 0
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(Me.TxtCode.Text) = True Then
                MsgBox("Code cannot empty string", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Select()
                Exit Sub
            End If

            If Me.DgvItem.Rows.Count <= 0 Then
                MsgBox("Data item not found.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                MsgBox("Data item with competence code invalid please reload again", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtCode.Text <> Me.ObjKomp.GetDataCode(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed code at master competence, please reload code competence", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 1
                If Me.DgvItem.Item(Me.ColItem_Modify.Index, ij).Value = True Then
                    DataBenar = True
                    Exit For
                End If
            Next

            If DataBenar = False Then
                MsgBox("No data selected", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Result = MsgBox("Are you sure delete this data?", MsgBoxStyle.OkCancel, "Confirm")
            If Result = MsgBoxResult.Ok Then
                For Each rowDGV As DataGridViewRow In Me.DgvItem.Rows
                    With rowDGV
                        If String.IsNullOrEmpty(.Cells(Me.ColItem_Value.Index).Value) = False Then
                            If .Cells(Me.ColItem_Modify.Index).Value = True Then
                                DTA.DeleteQuery(.Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text)
                                DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, .Cells(Me.ColItem_PrimaryEmp.Index).Value)
                                ICount += 1
                            End If
                        End If
                    End With
                Next
                If ICount > 0 Then
                    Me.IsikanData()
                    MsgBox("Successed", MsgBoxStyle.Information, "Information")
                    Me.LlbChkAll_Click(sender, e)
                    Me.LlbChkAllExp_Click(sender, e)
                Else
                    MsgBox("Data value not found for delete", MsgBoxStyle.Information, "Information")
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_Kompetensi_Upd_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.TxtCode.Select()
    End Sub

    Private Sub FrmMasterKaryawan_Kompetensi_Upd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.DgvSalFormula.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.TxtCode.Focus()
            Me.TxtPrimary.Visible = False
            Me.ChkShowAll.Checked = True
            Me.cboDataType.Enabled = False
            Me.cboDefValue.Enabled = False
            Me.cboAlfabetic.Enabled = False
            Me.TxtDefValueDigit.ReadOnly = True
            Me.TxtDefValueText.ReadOnly = True
            Me.DtpDefValue.Enabled = False
            Me.DgvItem.MultiSelect = False
            Me.dgvEmployee.MultiSelect = False
            Me.dgvJobPosition.MultiSelect = False
            Me.dgvDepartment.MultiSelect = False
            Me.DgvSalFormula.MultiSelect = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwEmployee_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwEmployee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwEmployee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwEmployee.RunWorkerCompleted
        Try
            If dgvEmployee.RowCount > 0 Then
                Me.dgvEmployee.Rows.RemoveAt(Me.IndexEmp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDepartment.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDepartment.RunWorkerCompleted
        Try
            If dgvDepartment.RowCount > 0 Then
                Me.dgvDepartment.Rows.RemoveAt(Me.IndexDept)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJobPosition_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwJobPosition.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJobPosition_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwJobPosition.RunWorkerCompleted
        Try
            If dgvJobPosition.RowCount > 0 Then
                Me.dgvJobPosition.Rows.RemoveAt(Me.IndexJobPost)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSalFor_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSalFor.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSalFor_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSalFor.RunWorkerCompleted
        Try
            If DgvSalFormula.RowCount > 0 Then
                Me.DgvSalFormula.Rows.RemoveAt(Me.IndexSalFor)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCompetace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCompetace.Click
        Dim i As New FrmMasterKompetensi_Cari
        Dim obj As New ClsMasterKompetensi

        Try
            i.BackColorAll = Me.BAckColorAll
            If Me.TxtCode.Text <> "" Then
                i.txtCari.Text = Me.TxtCode.Text
            End If
            i.ShowDialog()
            If i.mPrimary > 0 Then
                i.txtprimarygroup.Text = i.mPrimary
                Me.TxtPrimary.Text = i.mPrimary
                Me.TxtCode.Text = obj.GetDataCode(i.mPrimary)
                Me.TxtDesc.Text = obj.GetDesc(i.mPrimary)
            Else
                Me.TxtCode.Clear()
                Me.TxtDesc.Clear()
                Me.TxtPrimary.Text = "-1"
                Me.TxtCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
            Me.AktifkanCbo()
        End Try
    End Sub

    Private Sub TxtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus
        Me.TxtCode.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCode.Leave
        Dim i As New FrmMasterKompetensi_Cari
        Dim obj As New ClsMasterKompetensi

        Try
            i.BackColorAll = Me.BAckColorAll
            If Me.TxtCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    TxtPrimary.Text = obj.GetPrimary(Me.TxtCode.Text)
                    Me.TxtDesc.Text = obj.GetDesc(TxtPrimary.Text)
                Else
                    If Me.TxtCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCode.Text
                    End If

                    i.mPrimary = Me.TxtPrimary.Text
                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        TxtPrimary.Text = i.mPrimary
                        Me.TxtCode.Text = obj.GetDataCode(TxtPrimary.Text)
                        Me.TxtDesc.Text = obj.GetDesc(TxtPrimary.Text)
                    Else
                        Me.TxtCode.Clear()
                        Me.TxtDesc.Clear()
                        TxtPrimary.Text = "-1"
                        Me.TxtCode.Focus()
                    End If
                End If
            Else
                TxtPrimary.Text = "-1"
                Me.TxtDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
            Me.AktifkanCbo()
        End Try
    End Sub

    Private Sub TxtCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCode.KeyDown
        Dim i As New FrmMasterKompetensi_Cari
        Dim Obj As New ClsMasterKompetensi

        Try
            i.BackColorAll = Me.BAckColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    TxtPrimary.Text = Obj.GetPrimary(Me.TxtCode.Text)
                    Me.TxtDesc.Text = Obj.GetDesc(Me.TxtPrimary.Text)
                    Me.BtnShow.Focus()
                Else
                    If Me.TxtCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimary > 0 Then
                        TxtPrimary.Text = i.mPrimary
                        Me.TxtCode.Text = Obj.GetDataCode(TxtPrimary.Text)
                        Me.TxtDesc.Text = Obj.GetDesc(Me.TxtPrimary.Text)
                        Me.BtnShow.Focus()
                    Else
                        Me.TxtCode.Clear()
                        Me.TxtDesc.Clear()
                        TxtPrimary.Text = "-1"
                        Me.TxtCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
            Me.AktifkanCbo()
        End Try
    End Sub

    Private Sub AktifkanCbo()
        Dim obj As New ClsMasterKompetensi
        Dim DataType As Integer = 0
        Try
            DataType = obj.GetDataType(Me.TxtPrimary.Text)
            Me.cboDataType.SelectedIndex = DataType
            If DataType = ClsMasterKompetensi.EnumDataType.Number Then
                Me.DefPrec = obj.GetDefPrec(Me.TxtPrimary.Text)
                Me.TxtDefValueDigit.Text = ObjCD.DigitGrouping(obj.GetValue(Me.TxtPrimary.Text), Me.DefPrec)
            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                Me.DtpDefValue.Value = CDate(obj.GetValue(Me.TxtPrimary.Text))
            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Text Then
                Me.TxtDefValueText.Text = obj.GetValue(Me.TxtPrimary.Text)
            ElseIf DataType = ClsMasterKompetensi.EnumDataType.YN Then
                Me.cboDefValue.SelectedIndex = CInt(obj.GetValue(Me.TxtPrimary.Text))
            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                Me.cboAlfabetic.SelectedIndex = CInt(obj.GetValue(Me.TxtPrimary.Text))
            End If

            If String.IsNullOrEmpty(Me.TxtCode.Text) = False Then
                'Me.cboDataType.Enabled = True
                'Me.cboDefValue.Enabled = True
            Else
                Me.cboDataType.Enabled = False
                Me.cboDataType.SelectedIndex = -1
                Me.cboDefValue.Enabled = False
                Me.cboDefValue.SelectedIndex = -1
                Me.cboAlfabetic.Enabled = False
                Me.cboAlfabetic.SelectedIndex = -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus
        Me.TxtCode.BackColor = Me.TxtPrimary.BackColor
    End Sub

    Private Sub LlbChkAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbChkAll.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Me.DgvItem.Rows.Count - 1
                Me.DgvItem.Item(Me.ColItem_Modify.Index, i).Value = True
            Next

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUnchkAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbUnchkAll.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Me.DgvItem.Rows.Count - 1
                Me.DgvItem.Item(Me.ColItem_Modify.Index, i).Value = False
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemDepartment(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvDepartment.Rows.Count - 1
                RowDGV = Me.dgvDepartment.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColDeptPrimary.Index).Value = vPrimary Then
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

    Private Sub dgvDepartment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDepartment.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDepartment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDepartment.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvDepartment.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDept As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.dgvDepartment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDepartment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColDeptCode.Index Then
                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjDept.ApakahSudahAda(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryDept = Me.ObjDept.GetPrimary(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemDepartment(PrimaryDept, e.RowIndex)
                            If index < 0 Then
                                Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = PrimaryDept
                                Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(PrimaryDept)
                                Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(PrimaryDept)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterDepartementCari
                            ij.txtCari.Text = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.PrimaryDepartement <> -1 Then
                                index = Me.GetIndexItemDepartment(ij.PrimaryDepartement, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                    Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                    Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                Else
                                    If Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexDept = e.RowIndex
                                            Me.BgwDepartment.RunWorkerAsync()
                                        Else
                                            Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvDepartment)
                                        Exit Sub
                                    Else
                                        Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                        Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                        Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                    End If
                                End If
                            ElseIf ij.PrimaryDepartement = -1 Then
                                If Me.dgvDepartment.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                                If Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvDepartment)
                    End If
                Else
                    If e.RowIndex <> Me.dgvDepartment.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.IndexDept = e.RowIndex
                            Me.BgwDepartment.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvDepartment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvDepartment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub dgvDepartment_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvDepartment.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvDepartment, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJobPosition_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvJobPosition.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvJobPosition.Item(Me.ColDeptCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJobPosition_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvJobPosition.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvJobPosition.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryJobPost As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.dgvJobPosition.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvJobPosition.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColJobPosisitonCode.Index Then
                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjJobPost.ApakahSudahAda(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryJobPost = Me.ObjJobPost.GetPrimary(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemJobPosition(PrimaryJobPost, e.RowIndex)
                            If index < 0 Then
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = PrimaryJobPost
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(PrimaryJobPost)
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(PrimaryJobPost)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterJobPosition_Cari
                            ij.txtCari.Text = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemJobPosition(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                Else
                                    If Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexJobPost = e.RowIndex
                                            Me.BgwJobPosition.RunWorkerAsync()
                                        Else
                                            Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                        Exit Sub
                                    Else
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvJobPosition.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                                If Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvJobPosition)
                    End If
                Else
                    If e.RowIndex <> Me.dgvJobPosition.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.IndexJobPost = e.RowIndex
                            Me.BgwJobPosition.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvJobPosition.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvJobPosition.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Function GetIndexItemJobPosition(ByVal vPrimary As Double, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvJobPosition.Rows.Count - 1
                RowDGV = Me.dgvJobPosition.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColJobPosisitonPrimary.Index).Value = vPrimary Then
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

    Private Sub dgvJobPosition_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvJobPosition.RowsRemoved
        Try
            Me.ObjCD.SetNumber(dgvJobPosition, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                    Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                    Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryEmployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryEmployee = Me.ObjEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemEmployee(PrimaryEmployee, e.RowIndex)
                            If index < 0 Then
                                Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = PrimaryEmployee
                                Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(PrimaryEmployee)
                                Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(PrimaryEmployee)
                                Me.dgvEmployee.Item(Me.ColEmployee_NIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(PrimaryEmployee)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmployee_NIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(ij.PrimaryEmploye)
                                Else
                                    If Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")

                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexEmp = e.RowIndex
                                            BgwEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColEmployee_NIK.Index, e.RowIndex).Value = Me.ObjEmp.GetNIK(ij.PrimaryEmploye)
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        Me.BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                                If Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        Me.BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvEmployee)
                    End If
                Else
                    If e.RowIndex <> Me.dgvEmployee.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.IndexEmp = e.RowIndex
                            Me.BgwEmployee.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Function GetIndexItemEmployee(ByVal vPrimary As Double, _
        ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvEmployee.Rows.Count - 1
                RowDGV = Me.dgvEmployee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColEmployee_primary.Index).Value = vPrimary Then
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

    Private Sub dgvEmployee_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvEmployee.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvEmployee, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalFormula_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalFormula.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColSalFor_Code.Index Then
                    Me.Buffer = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColSalFor_Code.Index Then
                    Me.Buffer = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalFormula_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalFormula.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalFormula.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryEmployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                Me.DgvSalFormula.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvSalFormula.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If IsNothing(Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColSalFor_Code.Index Then
                        Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjSalFor.ApakahSudahAda(Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryEmployee = Me.ObjSalFor.GetPrimaryByCode(Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexDgvSalaryFor(PrimaryEmployee, e.RowIndex)
                            If index < 0 Then
                                Me.DgvSalFormula.Item(Me.ColSalFor_Primary.Index, e.RowIndex).Value = PrimaryEmployee
                                Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value = Me.ObjSalFor.GetCodeByPrimary(PrimaryEmployee)
                                Me.DgvSalFormula.Item(Me.ColSalFor_Desc.Index, e.RowIndex).Value = Me.ObjSalFor.GetDescription(PrimaryEmployee)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexSalFor = e.RowIndex
                                        BgwSalFor.RunWorkerAsync()
                                    Else
                                        Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalFormula)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterSalaryFormula_Cari
                            ij.txtCari.Text = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BAckColorAll
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                index = Me.GetIndexDgvSalaryFor(ij.Primary, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvSalFormula.Item(Me.ColSalFor_Primary.Index, e.RowIndex).Value = ij.Primary
                                    Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value = Me.ObjSalFor.GetCodeByPrimary(ij.Primary)
                                    Me.DgvSalFormula.Item(Me.ColSalFor_Desc.Index, e.RowIndex).Value = Me.ObjSalFor.GetDescription(ij.Primary)
                                Else
                                    If Me.Buffer = Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")

                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexSalFor = e.RowIndex
                                            BgwSalFor.RunWorkerAsync()
                                        Else
                                            Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalFormula)
                                        Exit Sub
                                    Else
                                        Me.DgvSalFormula.Item(Me.ColSalFor_Primary.Index, e.RowIndex).Value = ij.Primary
                                        Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value = Me.ObjSalFor.GetCodeByPrimary(ij.Primary)
                                        Me.DgvSalFormula.Item(Me.ColSalFor_Desc.Index, e.RowIndex).Value = Me.ObjSalFor.GetDescription(ij.Primary)
                                    End If
                                End If
                            ElseIf ij.Primary = -1 Then
                                If Me.DgvSalFormula.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexSalFor = e.RowIndex
                                        Me.BgwSalFor.RunWorkerAsync()
                                    Else
                                        Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalFormula)
                                    Exit Sub
                                End If
                                If Me.DgvSalFormula.Item(Me.ColSalFor_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexSalFor = e.RowIndex
                                        Me.BgwSalFor.RunWorkerAsync()
                                    Else
                                        Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalFormula)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalFormula)
                    End If
                Else
                    If e.RowIndex <> Me.DgvSalFormula.Rows.Count - 1 Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.IndexSalFor = e.RowIndex
                            Me.BgwSalFor.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.DgvSalFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvSalFormula.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvSalFormula.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub DgvSalFormula_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalFormula.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvSalFormula, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexDgvSalaryFor(ByVal vPrimary As Double, _
            ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalFormula.Rows.Count - 1
                RowDGV = Me.DgvSalFormula.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSalFor_Primary.Index).Value = vPrimary Then
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

    Private Sub CboTypeValue_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDataType.SelectedIndexChanged

        Try
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Number) Then
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueDigit.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = True
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Tanggal) Then
                Me.ClearObjectDefaultValue(True)
                Me.DtpDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = True
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Text) Then
                Me.ClearObjectDefaultValue(True)
                Me.TxtDefValueText.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = True
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.YN) Then
                Me.ClearObjectDefaultValue(True)
                Me.cboDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = True
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = False

                Exit Sub
            End If

            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                Me.ClearObjectDefaultValue(True)
                Me.cboDefValue.BringToFront()
                Me.TxtDefValueText.TabStop = False
                Me.TxtDefValueDigit.TabStop = False
                Me.DtpDefValue.TabStop = False
                Me.cboDefValue.TabStop = False
                Me.cboAlfabetic.TabStop = False

                Me.TxtDefValueDigit.Visible = False
                Me.TxtDefValueText.Visible = False
                Me.cboDefValue.Visible = False
                Me.DtpDefValue.Visible = False
                Me.cboAlfabetic.Visible = True

                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClearObjectDefaultValue(ByVal Semua As Boolean, _
       Optional ByVal IndexCBOData As Integer = -1)
        Me.DefPrec = Me.ObjKomp.GetDefPrec(Me.TxtPrimary.Text)
        Try
            If Semua = True Then
                Me.TxtDefValueDigit.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
                Me.TxtDefValueText.Clear()
                Me.DtpDefValue.Value = Now
                Me.cboDefValue.SelectedIndex = 0
                Me.cboAlfabetic.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.TxtLastPrimary.Text = Me.TxtPrimary.Text
            Me.GetRefreshTypeValue()
            If String.IsNullOrEmpty(Me.TxtCode.Text) = True Then
                MsgBox("Code cannot empty string.", MsgBoxStyle.Information, "information")
                Me.TxtCode.Focus()
                Exit Sub
            End If

            IsikanData()
            Me.LlbChkAll_Click(sender, e)
            Me.LlbChkAllExp_Click(sender, e)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Click
        Try
            Me.ObjCD.Number_Click(Me.TxtDefValueDigit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDefValueDigit, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefValueDigit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefValueDigit.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDefValueDigit, sender, e, Me.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanData()
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseMasterKaryawan_Kompetensi_Upd.DataTable1DataTable
        Dim Row As DseMasterKaryawan_Kompetensi_Upd.DataTable1Row
        Dim DTA2 As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.DataTable2TableAdapter
        Dim TBL2 As New DseMasterKaryawan_Kompetensi_Upd.DataTable2DataTable
        Dim Row2 As DseMasterKaryawan_Kompetensi_Upd.DataTable2Row
        Dim Filter As String = ""
        Dim RowDGV As DataGridViewRow

        Try
            Me.ColItem_NextCertDate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            If Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                Me.ColItem_Value.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Number) Then
                Me.ColItem_Value.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal) Then
                Me.ColItem_Value.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Text) Then
                Me.ColItem_Value.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.YN) Then
                Me.ColItem_Value.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If
            Me.ColItem_NextCertDate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvItem.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkShowAll.Checked = False Then
                DTA.Fill(TBL, Me.TxtPrimary.Text, Core_Module.ClsMasterKaryawan.EnumAktif.Active)
                'Me.FilterData1(TBL)
                Me.Filtering(TBL)
                If TBL.Rows.Count > 0 Then
                    Filter = Filter
                    TBL.DefaultView.RowFilter = Filter
                    TBL.DefaultView.Sort = "[Code] ASC, [Full Name] ASC"
                    If TBL.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TBL.DefaultView.Count - 1
                            Row = TBL.DefaultView(ij).Row
                            Me.DgvItem.Rows.Add()
                            RowDGV = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                            With RowDGV
                                .Cells(Me.ColItem_PrimaryEmp.Index).Value = Row.primary_karyawan
                                .Cells(Me.ColItem_EmpCode.Index).Value = Row.Code
                                .Cells(Me.ColItem_EmpName.Index).Value = Row.Full_Name
                                .Cells(Me.ColItem_EmpNIK.Index).Value = Row.NIK
                                .Cells(Me.ColItem_PrimaryJobPost.Index).Value = Row.Primary_Job_Post
                                .Cells(Me.ColItem_JobPos.Index).Value = Row.Job_Post_Desc
                                .Cells(Me.ColItem_Dept.Index).Value = Row.Dept_Desc

                                If Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                                    .Cells(Me.ColItem_Value.Index).Value = Row.Value_Karyawan
                                ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Number) Then
                                    If String.IsNullOrEmpty(Row.Value_Karyawan) = False Then
                                        .Cells(Me.ColItem_Value.Index).Value = Me.ObjCD.DigitGrouping(Row.Value_Karyawan, Me.DefPrec)
                                    End If
                                ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal) Then
                                    If String.IsNullOrEmpty(Row.Value_Karyawan) = False Then
                                        .Cells(Me.ColItem_Value.Index).Value = Format(CDate(Row.Value_Karyawan).Date, "dd MMM yyyy")
                                    End If
                                ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Text) Then
                                    .Cells(Me.ColItem_Value.Index).Value = Row.Value_Karyawan
                                ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.YN) Then
                                    .Cells(Me.ColItem_Value.Index).Value = Row.Value_Karyawan
                                End If

                                If Row.NextCertf_Karyawan <= #1/1/1950# Then
                                    .Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                Else
                                    .Cells(Me.ColItem_NextCertDate.Index).Value = Format(Row.NextCertf_Karyawan, "dd MMM yyyy")
                                End If
                                .Cells(Me.ColItem_Remarks.Index).Value = Row.Remarks_Karyawan
                                '.Cells(Me.ColItem_Modify.Index).Value = False
                            End With
                        Next
                        Me.ObjCD.SetNumber(Me.DgvItem, False)
                    End If
                End If
                Me.TBLEmp = TBL.Copy
            Else
                DTA.Fill(TBL, Me.TxtPrimary.Text, Core_Module.ClsMasterKaryawan.EnumAktif.Active)
                'Me.FilterData1(TBL)
                DTA2.FillByALL(TBL2, Core_Module.ClsMasterKaryawan.EnumAktif.Active)

                For i As Integer = 0 To TBL.Rows.Count - 1
                    Row = TBL.Rows(i)
                    TBL2.DefaultView.RowFilter = "[primary karyawan]=" & Row.primary_karyawan
                    If TBL2.DefaultView.Count > 0 Then
                        Row2 = TBL2.DefaultView(0).Row
                        Row2.Value_Karyawan = Row.Value_Karyawan
                        Row2.NextCertf_Karyawan = Row.NextCertf_Karyawan
                        Row2.Remarks_Karyawan = Row.Remarks_Karyawan
                    End If
                Next

                Me.Filtering(TBL2)

                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For ij As Integer = 0 To TBL2.Rows.Count - 1
                    Row2 = TBL2.Rows(ij)
                    Me.DgvItem.Rows.Add()
                    RowDGV = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                    With RowDGV
                        .Cells(Me.ColItem_PrimaryEmp.Index).Value = Row2.primary_karyawan
                        .Cells(Me.ColItem_EmpCode.Index).Value = Row2.Code
                        .Cells(Me.ColItem_EmpName.Index).Value = Row2.Full_Name
                        .Cells(Me.ColItem_EmpNIK.Index).Value = Row2.NIK
                        .Cells(Me.ColItem_PrimaryJobPost.Index).Value = Row2.Primary_Job_Post
                        .Cells(Me.ColItem_JobPos.Index).Value = Row2.Job_Post_Desc
                        .Cells(Me.ColItem_Dept.Index).Value = Row2.Dept_Desc

                        If Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                            .Cells(Me.ColItem_Value.Index).Value = Row2.Value_Karyawan
                        ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Number) Then
                            If String.IsNullOrEmpty(Row2.Value_Karyawan) = False Then
                                .Cells(Me.ColItem_Value.Index).Value = Me.ObjCD.DigitGrouping(Row2.Value_Karyawan, Me.DefPrec)
                            End If
                        ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal) Then
                            If String.IsNullOrEmpty(Row2.Value_Karyawan) = False Then
                                .Cells(Me.ColItem_Value.Index).Value = Format(CDate(Row2.Value_Karyawan).Date, "dd MMM yyyy")
                            End If
                        ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.Text) Then
                            .Cells(Me.ColItem_Value.Index).Value = Row2.Value_Karyawan
                        ElseIf Me.cboDataType.SelectedIndex = CInt(Core_Module.ClsMasterKompetensi.EnumDataType.YN) Then
                            .Cells(Me.ColItem_Value.Index).Value = Row2.Value_Karyawan
                        End If

                        If Row2.NextCertf_Karyawan <= #1/1/1950# Then
                            .Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                        Else
                            .Cells(Me.ColItem_NextCertDate.Index).Value = Format(Row2.NextCertf_Karyawan, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColItem_Remarks.Index).Value = Row2.Remarks_Karyawan
                        '.Cells(Me.ColItem_Modify.Index).Value = False
                    End With
                Next
                Me.ObjCD.SetNumber(Me.DgvItem, False)
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.TBLEmp2 = TBL2.Copy
            End If

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default

        End Try
    End Sub

    Private Function FilterData1(ByRef vTbl As Data.DataTable) As Boolean
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim Filter As String = ""
        Try
            tblDummy = vTbl.Copy
            Filter = "[Primary]=Primary"
            Filter &= " AND [Data Type]='" & cboDataType.SelectedIndex & "'"
            If cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Number) Then
                Filter &= " AND [Standar Value]='" & Me.TxtDefValueDigit.Text.ToString & "'"
            ElseIf cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Tanggal) Then
                Filter &= " AND [Standar Value]='" & Me.DtpDefValue.Value.ToString & "'"
            ElseIf cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Text) Then
                Filter &= " AND [Standar Value]='" & Me.TxtDefValueText.Text.ToString & "'"
            ElseIf cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.YN) Then
                Filter &= " AND [Standar Value]='" & Me.cboDefValue.SelectedItem.ToString & "'"
            ElseIf cboDataType.SelectedIndex = CInt(ClsMasterKompetensi.EnumDataType.Alfabetic) Then
                Filter &= " AND [Standar Value]='" & Me.cboAlfabetic.SelectedItem.ToString & "'"
            End If

            tblDummy.DefaultView.RowFilter = Filter
            tblDummy.DefaultView.Sort = "[Code] ASC, [Full Name] ASC"
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(ij).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetFilterByEmployee() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.dgvEmployee.Rows
                With dr
                    If dr.Cells(Me.ColEmployee_primary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColEmployee_Code.Index).Value) = False Then
                        If Filter = "" Then
                            Filter = "[primary karyawan] = " & dr.Cells(Me.ColEmployee_primary.Index).Value
                        Else
                            Filter += " OR [primary karyawan] = " & dr.Cells(Me.ColEmployee_primary.Index).Value
                        End If
                    End If
                End With
            Next
            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Filter
        End Try
    End Function

    Private Function GetFilterByDepartemen() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.dgvDepartment.Rows
                With dr
                    If dr.Cells(Me.ColDeptPrimary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColDeptCode.Index).Value) = False Then
                        If Filter = "" Then
                            Filter = "[primary departemen] = " & .Cells(Me.ColDeptPrimary.Index).Value.ToString
                        Else
                            Filter += " OR [primary departemen] = " & .Cells(Me.ColDeptPrimary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Filter
        End Try
    End Function

    Private Function GetFilterByPosition() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.dgvJobPosition.Rows
                With dr
                    If dr.Cells(Me.ColJobPosisitonPrimary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColJobPosisitonCode.Index).Value) = False Then
                        If Filter = "" Then
                            Filter = "[Primary Job Post] = " & .Cells(Me.ColJobPosisitonPrimary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary Job Post] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Filter
        End Try
    End Function

    Private Function GetFilterBySalaryFormula() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.DgvSalFormula.Rows
                With dr
                    If dr.Cells(Me.ColSalFor_Primary.Index).Value <> -1 And String.IsNullOrEmpty(dr.Cells(Me.ColSalFor_Code.Index).Value) = False Then
                        If Filter = "" Then
                            Filter = "[primary salary formula] = " & .Cells(Me.ColSalFor_Primary.Index).Value.ToString
                        Else
                            Filter += " OR [primary salary formula] = " & .Cells(Me.ColSalFor_Primary.Index).Value.ToString
                        End If
                    End If
                End With
            Next
            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Filter
        End Try
    End Function

    Sub Filtering(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim FilterAll As String = ""
        Dim FilterEmployee As String = ""
        Dim FilterDepartemen As String = ""
        Dim FilterPosisi As String = ""
        Dim FilterSalfor As String = ""
        Try
            tblDummy = vTbl.Copy
            FilterEmployee = GetFilterByEmployee()
            If FilterEmployee <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterEmployee & ")"
                Else
                    FilterAll &= "(" & FilterEmployee & ")"
                End If
            End If
            FilterDepartemen = GetFilterByDepartemen()
            If FilterDepartemen <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterDepartemen & ")"
                Else
                    FilterAll &= "(" & FilterDepartemen & ")"
                End If
            End If
            FilterPosisi = GetFilterByPosition()
            If FilterPosisi <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterPosisi & ")"
                Else
                    FilterAll &= "(" & FilterPosisi & ")"
                End If
            End If
            FilterSalfor = GetFilterBySalaryFormula()
            If FilterSalfor <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterSalfor & ")"
                Else
                    FilterAll &= "(" & FilterSalfor & ")"
                End If
            End If


            tblDummy.DefaultView.RowFilter = FilterAll
            tblDummy.DefaultView.Sort = "[Code] ASC, [Full Name] ASC"
            vTbl.Rows.Clear()
            If tblDummy.DefaultView.Count > 0 Then
                For ij As Integer = 0 To tblDummy.DefaultView.Count - 1
                    rowDummy = tblDummy.DefaultView(ij).Row
                    rowTbl = vTbl.NewRow
                    For ii As Integer = 0 To tblDummy.Columns.Count - 1
                        rowTbl.Item(ii) = rowDummy.Item(ii)
                    Next
                    vTbl.Rows.Add(rowTbl)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick
        Dim MsgR As MsgBoxResult
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim Icount As Integer = 0
        Try
            If e.RowIndex < 0 Then
                Exit Sub
            End If
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.DgvItem.Rows.Count > 0 And e.ColumnIndex = Me.ColItem_Del.Index Then
                If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                    MsgBox("Data item with competence code invalid please reload again", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                'Me.Rec()
                If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                    MsgR = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confrm")
                    If MsgR = MsgBoxResult.Yes Then
                        DTA.DeleteQuery(DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                        DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                        Icount += 1
                    End If
                Else
                    MsgBox("Data value not found for delete.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Icount > 0 Then
                    Me.IsikanData()
                    Me.LlbChkAll_Click(sender, e)
                    Me.LlbChkAllExp_Click(sender, e)
                    'Me.UnRec(e.RowIndex)
                    MsgBox("Successed", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                MsgBox("Data item with competence code invalid please reload again", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If
            If Me.cboDataType.SelectedIndex <> Me.ObjKomp.GetDataType(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed type value at master competence, please reload again", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColItem_Value.Index Or e.ColumnIndex = Me.ColItem_NextCertDate.Index Then

                If e.ColumnIndex = Me.ColItem_Value.Index Then
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
                If e.ColumnIndex = Me.ColItem_NextCertDate.Index Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                        Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColItem_Remarks.Index Then
                If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = True Then
                    e.Cancel = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColItem_Value.Index Or e.ColumnIndex = Me.ColItem_Remarks.Index Or e.ColumnIndex = Me.ColItem_NextCertDate.Index Then
                    If Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColItem_Value.Index Then
                        If Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "A" Or _
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "B" Or _
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "C" Or _
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "D" Or _
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "E" Or _
                                                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "F" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            Else
                                MsgBox("Must input A , B , C ,D , E ,F", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Number Then
                            If Me.ObjCD.CekNumerik(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjKomp.SetFormatDataCustom(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.cboDataType.SelectedIndex, Me.DefPrec, Me.ObjCD)
                            End If

                            If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for data.", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjKomp.SetFormatDataCustom(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.cboDataType.SelectedIndex, DefPrec, Me.ObjCD)
                            End If
                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Text Then
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                                          Me.ObjKomp.SetFormatDataCustom(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.cboDataType.SelectedIndex, DefPrec, Me.ObjCD)
                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.YN Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "YES" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                            ElseIf Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "NO" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                            ElseIf Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "Y" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                            ElseIf Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "N" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                            Else
                                MsgBox("Must input Yes / No .", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            'new
                            Dim Value As String = ""
                            Dim NextCert As Date = #1/1/1950#
                            Dim Remarks As String = ""

                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                                Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                            Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                                NextCert = #1/1/1950#
                            Else
                                NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                                Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                            End If

                            DTA.InsertQuery(Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text, Value, Remarks, NextCert)
                            DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                            Exit Sub
                        Else
                            'update
                            Dim Value As String = ""
                            Dim NextCert As String = #1/1/1950#
                            Dim Remarks As String = ""

                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                                Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                            Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                                NextCert = #1/1/1950#
                            Else
                                NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                                Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                            End If

                            DTA.UpdateQuery_Value(Value, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                            DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                            Exit Sub
                        End If
                    End If

                    If e.ColumnIndex = Me.ColItem_NextCertDate.Index Then
                        If Me.ObjJobPost.GetPeriodic(Me.DgvItem.Item(Me.ColItem_PrimaryJobPost.Index, e.RowIndex).Value, Me.TxtPrimary.Text) = 1 Then
                            'jika di master job position periodic update nya di ceklis maka harus input tanggal di next certification

                            ''set region ke indonesia
                            'My.Application.ChangeCulture("id-ID")
                            'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                            End If
                            'kembalikan ke region defaultnya
                            My.Application.ChangeCulture("en-US")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                        ElseIf Me.ObjJobPost.GetPeriodic(Me.DgvItem.Item(Me.ColItem_PrimaryJobPost.Index, e.RowIndex).Value, Me.TxtPrimary.Text) = -1 Then
                            'set region ke indonesia
                            'My.Application.ChangeCulture("id-ID")
                            'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                                ElseIf Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> "-" Then
                                    MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                            End If
                            'kembalikan ke region defaultnya
                            My.Application.ChangeCulture("en-US")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                        End If

                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            'new
                            Dim Value As String = ""
                            Dim NextCert As Date = #1/1/1950#
                            Dim Remarks As String = ""

                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                                Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                                 Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                                NextCert = #1/1/1950#
                            Else
                                NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                                Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                            End If

                            DTA.InsertQuery(Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text, Value, Remarks, NextCert)
                            DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                            Exit Sub
                        Else
                            Dim Value As String = ""
                            Dim NextCert As Date = #1/1/1950#
                            Dim Remarks As String = ""

                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                                Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                                 Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                                NextCert = #1/1/1950#
                            Else
                                NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                            End If
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                                Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                            End If

                            DTA.UpdateQuery_NextCertf(NextCert, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                            DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                            Exit Sub
                        End If
                    End If
                End If

                If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                    Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value.ToString.Replace("'", "").Replace("*", "").Replace("^", "").Replace("%", "").ToString.Trim
                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                        'new
                        Dim Value As String = ""
                        Dim NextCert As Date = #1/1/1950#
                        Dim Remarks As String = ""

                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                            Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                        End If
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                           Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                            NextCert = #1/1/1950#
                        Else
                            NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                        End If
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                            Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                        End If
                        DTA.InsertQuery(Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text, Value, Remarks, NextCert)
                        DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                        Exit Sub
                    Else
                        Dim Value As String = ""
                        Dim NextCert As Date = #1/1/1950#
                        Dim Remarks As String = ""

                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value) = False Then
                            Value = Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value
                        End If
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value) = True Or _
                           Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value = "-" Then
                            NextCert = #1/1/1950#
                        Else
                            NextCert = CDate(Me.DgvItem.Item(Me.ColItem_NextCertDate.Index, e.RowIndex).Value)
                        End If
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value) = False Then
                            Remarks = Me.DgvItem.Item(Me.ColItem_Remarks.Index, e.RowIndex).Value
                        End If

                        DTA.UpdateQuery_Remarks(Remarks, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                        DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                        Exit Sub
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColItem_Value.Index Then
                    If String.IsNullOrEmpty(Me.Buffer) = False Then
                        Me.DgvItem.Item(Me.ColItem_Value.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColItem_NextCertDate.Index Then
                    If String.IsNullOrEmpty(Me.Buffer) = False Then
                        DTA.UpdateQuery_NextCertf(#1/1/1950#, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                        DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                    End If
                ElseIf e.ColumnIndex = Me.ColItem_Remarks.Index Then
                    If String.IsNullOrEmpty(Me.Buffer) = False Then
                        DTA.UpdateQuery_Remarks("", Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value, Me.TxtPrimary.Text)
                        DTA.UpdateQueryModifyKary(Me.UserCode, Now, Now, Me.DgvItem.Item(Me.ColItem_PrimaryEmp.Index, e.RowIndex).Value)
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.DataTable1TableAdapter
        Dim DTA2 As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseMasterKaryawan_Kompetensi_Upd.DataTable1DataTable
        Dim TBL2 As New DseMasterKaryawan_Kompetensi_Upd.DataTable2DataTable
        Dim RowEmp As DseMasterKaryawan_Kompetensi_Upd.DataTable1Row
        Dim Row As DseMasterKaryawan_Kompetensi_Upd.DataTable1Row
        Dim Row2 As DseMasterKaryawan_Kompetensi_Upd.DataTable2Row
        Dim listEmployee As New List(Of String)
        Dim ix As Integer = 0
        Dim FileExtStr As String = ""
        Dim hj As New FrmMassageCSV

        Try
            If Me.ChkShowAll.Checked = False Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.Fill(TBL, Me.TxtPrimary.Text, Core_Module.ClsMasterKaryawan.EnumAktif.Active)
                Me.Filtering(TBL)
                TBLEmp = TBL.Copy
            Else
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA2.Connection.ConnectionString = Core_Module.GetConstring

                DTA.Fill(TBL, Me.TxtPrimary.Text, Core_Module.ClsMasterKaryawan.EnumAktif.Active)
                DTA2.FillByALL(TBL2, Core_Module.ClsMasterKaryawan.EnumAktif.Active)

                For i As Integer = 0 To TBL.Rows.Count - 1
                    Row = TBL.Rows(i)
                    TBL2.DefaultView.RowFilter = "[primary karyawan]=" & Row.primary_karyawan
                    If TBL2.DefaultView.Count > 0 Then
                        Row2 = TBL2.DefaultView(0).Row
                        Row2.Value_Karyawan = Row.Value_Karyawan
                        Row2.NextCertf_Karyawan = Row.NextCertf_Karyawan
                        Row2.Remarks_Karyawan = Row.Remarks_Karyawan
                    End If
                Next

                Me.Filtering(TBL2)
                TBLEmp2 = TBL2.Copy
            End If

            If String.IsNullOrEmpty(Me.TxtCode.Text) = True Then
                MsgBox("Code cannot empty string", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Select()
                Exit Sub
            End If

            If Me.DgvItem.RowCount < 1 Then
                MsgBox("Data item not found", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                MsgBox("Data item with competence code invalid please reload again", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtCode.Text <> Me.ObjKomp.GetDataCode(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed code at master competence, please reload code competence", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim vBenar As Boolean = False
            For i As Integer = 0 To Me.DgvItem.RowCount - 1
                If Me.DgvItem.Item(Me.ColItem_Export.Index, i).Value = True Then
                    vBenar = True
                End If
            Next
            If vBenar = False Then
                MsgBox("Modify not found selection.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.GetFrmDia(FileExtStr) = False Then
                Exit Sub
            End If

            'Me.IsikanData()
            For Each rowd As DataGridViewRow In Me.DgvItem.Rows
                If rowd.Cells(Me.ColEmployee_Code.Index).Value <> "" Then
                    TBLEmp.DefaultView.RowFilter = "[primary karyawan]=" & rowd.Cells(Me.ColItem_PrimaryEmp.Index).Value
                    If TBLEmp.DefaultView.Count > 0 Then
                        RowEmp = TBLEmp.DefaultView(0).Row
                        RowEmp.Urutan = CInt(rowd.Cells(Me.ColItem_No.Index).Value)
                    End If
                End If
            Next

            'While ix <= TBLEmp.Rows.Count - 1
            '    RowEmp = TBLEmp.Rows(ix)
            '    If RowEmp.Amount = 0 Then
            '        TBLEmp.Rows.Remove(RowEmp)
            '    Else
            '        ix += 1
            '    End If
            'End While

            If FileExtStr.StartsWith(".csv") Then
                If hj.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    If hj.Rbo = 0 Then
                        Me.ExportFileCsv(TBLEmp)
                    Else
                        Me.ExportFileCsv2(TBLEmp)
                    End If
                End If

                'ElseIf FileExtStr.StartsWith(".xlsx") Then
                '    Me.ExportFileExcel(TBLEmp)
                'ElseIf FileExtStr.StartsWith(".xls") Then
                '    Me.ExportFileExcel(TBLEmp)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function GetFrmDia(ByRef vFileExtStr As String) As Boolean
        Try
            Me.frmDia.Title = "Please specify file"
            Me.frmDia.Filter = "csv files (*.csv)|*.csv"
            'Me.FrmDia.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.frmDia.FilterIndex = 1
            'Me.FrmDia.DefaultExt = "xls"
            Me.frmDia.DefaultExt = "csv"
            Me.frmDia.FileName = ""
            If Me.frmDia.ShowDialog <> Windows.Forms.DialogResult.OK And Me.frmDia.FileName.Trim.Length = 0 Then
                Return False
            End If

            vFileExtStr = System.IO.Path.GetExtension(Me.frmDia.FileName)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ExportFileCsv(ByVal TBLEmp As DseMasterKaryawan_Kompetensi_Upd.DataTable1DataTable)
        Dim RowEmp As DseMasterKaryawan_Kompetensi_Upd.DataTable1Row
        Dim RowEmp2 As DseMasterKaryawan_Kompetensi_Upd.DataTable2Row
        Dim csvVal As String = ""
        'jika komputer menggunakan regional indonesia
        Try

            If Me.DgvItem.Rows.Count - 1 <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "Competence;Employee Code;Employee NIK;Employee Name;Job Position;Department;value;Next Certificate Date;Remarks"

                If Me.ChkShowAll.Checked = False Then
                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                        If rowD.Cells(Me.ColItem_Export.Index).Value = True Then
                            TBLEmp.DefaultView.RowFilter = Nothing
                            TBLEmp.DefaultView.RowFilter = "[primary karyawan]=" & rowD.Cells(Me.ColItem_PrimaryEmp.Index).Value
                            TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                            If TBLEmp.DefaultView.Count > 0 Then
                                RowEmp = TBLEmp.DefaultView(0).Row
                                With RowEmp
                                    csvVal += vbNewLine & Me.TxtCode.Text.ToString.Replace(",", "") & ";"
                                    csvVal += .Code.ToString.Replace(",", "") & ";"
                                    csvVal += .NIK.ToString.Replace(",", "") & ";"
                                    csvVal += .Full_Name.ToString.Replace(",", "") & ";"
                                    csvVal += .Job_Post_Code.ToString.Replace(",", "") & ";"
                                    csvVal += .Dept_Code.ToString.Replace(",", "") & ";"

                                    If String.IsNullOrEmpty(.Value_Karyawan) = False Then
                                        If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                            csvVal += Me.ObjCD.DigitGrouping(.Value_Karyawan, .Default_Precision).ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                            csvVal += Format(CDate(.Value_Karyawan).Date, "dd MMM yyyy").ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        End If
                                    Else
                                        csvVal += ("").ToString.Replace(",", "") & ";"
                                    End If

                                    If .NextCertf_Karyawan <= #1/1/1950# Then
                                        csvVal += ("-").ToString.Replace(",", "") & ";"
                                    Else
                                        csvVal += Format(CDate(.NextCertf_Karyawan).Date, "dd MMM yyyy").ToString.Replace(",", "") & ";"
                                    End If
                                    csvVal += .Remarks_Karyawan.ToString.Replace(",", "") & ";"
                                End With
                            End If
                        End If
                    Next

                Else
                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                        If rowD.Cells(Me.ColItem_Export.Index).Value = True Then
                            TBLEmp2.DefaultView.RowFilter = Nothing
                            TBLEmp2.DefaultView.RowFilter = "[primary karyawan]=" & rowD.Cells(Me.ColItem_PrimaryEmp.Index).Value
                            TBLEmp2.DefaultView.Sort = "[Urutan] ASC"
                            If TBLEmp2.DefaultView.Count > 0 Then
                                RowEmp2 = TBLEmp2.DefaultView(0).Row
                                With RowEmp2
                                    csvVal += vbNewLine & TxtCode.Text.ToString.Replace(",", "") & ";"
                                    csvVal += .Code.ToString.Replace(",", "") & ";"
                                    csvVal += .NIK.ToString.Replace(",", "") & ";"
                                    csvVal += .Full_Name.ToString.Replace(",", "") & ";"
                                    csvVal += .Job_Post_Code.ToString.Replace(",", "") & ";"
                                    csvVal += .Dept_Code.ToString.Replace(",", "") & ";"

                                    If String.IsNullOrEmpty(.Value_Karyawan) = False Then
                                        If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                            csvVal += Me.ObjCD.DigitGrouping(.Value_Karyawan, Me.DefPrec).ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                            csvVal += Format(CDate(.Value_Karyawan).Date, "dd MMM yyyy").ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                            csvVal += .Value_Karyawan.ToString.Replace(",", "") & ";"
                                        End If
                                    Else
                                        csvVal += ("").ToString.Replace(",", "") & ";"
                                    End If

                                    If .NextCertf_Karyawan <= #1/1/1950# Then
                                        csvVal += ("-").ToString.Replace(",", "") & ";"
                                    Else
                                        csvVal += Format(CDate(.NextCertf_Karyawan).Date, "dd MMM yyyy").ToString.Replace(",", "") & ";"
                                    End If
                                    csvVal += .Remarks_Karyawan.ToString.Replace(",", "") & ";"
                                End With
                            End If
                        End If
                    Next

                    'TBLEmp2.DefaultView.RowFilter = Nothing
                    'TBLEmp2.DefaultView.Sort = "[Urutan] ASC"
                    'For i As Integer = 0 To TBLEmp2.DefaultView.Count - 1

                    'Next

                End If

                My.Computer.FileSystem.WriteAllText(Me.frmDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '
    Private Sub ExportFileCsv2(ByVal TBLEmp As DseMasterKaryawan_Kompetensi_Upd.DataTable1DataTable)
        Dim RowEmp As DseMasterKaryawan_Kompetensi_Upd.DataTable1Row
        Dim RowEmp2 As DseMasterKaryawan_Kompetensi_Upd.DataTable2Row
        Dim csvVal As String = ""
        'jika komputer menggunakan regional english
        Try
            If Me.DgvItem.Rows.Count - 1 <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "Competence,Employee Code,Employee NIK,Employee Name,Job Position,Department,value,Next Certificate Date,Remarks"

                If Me.ChkShowAll.Checked = False Then
                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                        If rowD.Cells(Me.ColItem_Export.Index).Value = True Then
                            TBLEmp.DefaultView.RowFilter = Nothing
                            TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                            TBLEmp.DefaultView.RowFilter = "[primary karyawan]=" & rowD.Cells(Me.ColItem_PrimaryEmp.Index).Value
                            If TBLEmp.DefaultView.Count > 0 Then
                                RowEmp = TBLEmp.DefaultView(0).Row
                                With RowEmp
                                    csvVal += vbNewLine & TxtCode.Text.ToString.Replace(";", "") & ","
                                    csvVal += .Code.ToString.Replace(";", "") & ","
                                    csvVal += .NIK.ToString.Replace(";", "") & ","
                                    csvVal += .Full_Name.ToString.Replace(";", "") & ","
                                    csvVal += .Job_Post_Code.ToString.Replace(";", "") & ","
                                    csvVal += .Dept_Code.ToString.Replace(";", "") & ","

                                    If String.IsNullOrEmpty(.Value_Karyawan) = False Then
                                        If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                            csvVal += Me.ObjCD.DigitGrouping(.Value_Karyawan, .Default_Precision).ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                            csvVal += Format(CDate(.Value_Karyawan).Date, "dd MMM yyyy").ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        End If
                                    Else
                                        csvVal += ("").Replace(";", "") & ","
                                    End If

                                    If .NextCertf_Karyawan <= #1/1/1950# Then
                                        csvVal += ("-").ToString.Replace(";", "") & ","
                                    Else
                                        csvVal += Format(CDate(.NextCertf_Karyawan).Date, "dd MMM yyyy").ToString.Replace(";", "") & ","
                                    End If
                                    csvVal += .Remarks_Karyawan.ToString.Replace(";", "") & ","
                                End With
                            End If
                        End If
                    Next

                Else
                    For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                        If rowD.Cells(Me.ColItem_Export.Index).Value = True Then
                            TBLEmp2.DefaultView.RowFilter = Nothing
                            TBLEmp2.DefaultView.Sort = "[Urutan] ASC"
                            TBLEmp2.DefaultView.RowFilter = "[primary karyawan]=" & rowD.Cells(Me.ColItem_PrimaryEmp.Index).Value
                            If TBLEmp2.DefaultView.Count > 0 Then
                                RowEmp2 = TBLEmp2.DefaultView(0).Row
                                With RowEmp2
                                    csvVal += vbNewLine & TxtCode.Text.ToString.Replace(";", "") & ","
                                    csvVal += .Code.ToString.Replace(";", "") & ","
                                    csvVal += .NIK.ToString.Replace(";", "") & ","
                                    csvVal += .Full_Name.ToString.Replace(";", "") & ","
                                    csvVal += .Job_Post_Code.ToString.Replace(";", "") & ","
                                    csvVal += .Dept_Code.ToString.Replace(";", "") & ","

                                    If String.IsNullOrEmpty(.Value_Karyawan) = False Then
                                        If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                            csvVal += Me.ObjCD.DigitGrouping(.Value_Karyawan, Me.DefPrec).ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                            csvVal += Format(CDate(.Value_Karyawan).Date, "dd MMM yyyy").ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                            csvVal += .Value_Karyawan.ToString.Replace(";", "") & ","
                                        End If
                                    Else
                                        csvVal += ("").Replace(";", "") & ","
                                    End If

                                    If .NextCertf_Karyawan <= #1/1/1950# Then
                                        csvVal += ("-").ToString.Replace(";", "") & ","
                                    Else
                                        csvVal += Format(CDate(.NextCertf_Karyawan), "dd MMM yyyy").ToString.Replace(";", "") & ","
                                    End If
                                    csvVal += .Remarks_Karyawan.ToString.Replace(";", "")
                                End With
                            End If
                        End If
                    Next
                End If

                My.Computer.FileSystem.WriteAllText(Me.frmDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ImportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ImportToolStripMenuItem.Click
        Dim FileExtStr As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If String.IsNullOrEmpty(Me.TxtCode.Text) = True Then
                MsgBox("Code cannot empty string", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Select()
                Exit Sub
            End If

            If Me.TxtPrimary.Text <> Me.TxtLastPrimary.Text Then
                MsgBox("Data item with competence code invalid please reload item", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtCode.Text <> Me.ObjKomp.GetDataCode(Me.TxtPrimary.Text) Then
                MsgBox("Other user have changed code at master competence, please reload code competence", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.Openfile(FileExtStr) = False Then
                Exit Sub
            End If

            If FileExtStr.StartsWith(".csv") Then
                Me.ImportFile()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub ImportFile()
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.LocationFileSumber)
        Dim line As String = SR.ReadLine()
        Dim Tbl As DataTable = New DataTable
        Dim row As DataRow

        Dim NamaFilter As String = ""
        Dim JmlKaryawan As Integer = 0
        Dim ListNama_InDGV As New List(Of String)
        Dim ListNama_1 As New List(Of String) 'list karyawan di dgv yang ga ada di csv
        Dim ListNama_2 As New List(Of String) 'list karyawan di csv yang ga ada di dgv
        Dim ListCompetenceNotSame As New List(Of String)
        Dim ListKaryNotFound As New List(Of String)
        Dim ListTypeNotFound As New List(Of String)
        Dim StrMsgBox_1 As String = ""
        Dim StrMsgBox_2 As String = ""
        Dim StrMsgBox_3 As String = ""
        Dim StrMsgBox_4 As String = ""
        Dim StrMsgBox_5 As String = ""
        Dim StrMsgBox_6 As String = ""
        Dim ApaExitSub As Boolean = False
        Dim iCOunt As String() = line.Split(","c)
        Dim Icount2 As String() = line.Split(";"c)
        Dim strArray As String()

        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim StrValue As String = ""
        Dim ListKary As New List(Of String)
        Dim ApakahImport As Integer = 0
        Dim FrmImp As New FrmMasterKaryawan_Kompetensi_Upd_ImportOpt
        Dim DgvImp As DataGridView = FrmImp.DgvImport
        Dim DgvR As DataGridViewRow
        Dim vPrimaryKary, vCode, vNIK, vFullName, vJobPost, vDept As String
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If iCOunt.Length > Icount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If
            'pindahkan data csv ke dalam data table
            For Each s As String In strArray
                Tbl.Columns.Add(New DataColumn())
            Next
            Tbl.Columns.Add(New DataColumn())

            If Tbl.Columns.Count <> 6 Then
                MsgBox("Make sure csv file consist of competence, employee name, value, next certificate date and remarks.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Do
                line = SR.ReadLine
                If Not line = String.Empty Then
                    If iCOunt.Length > Icount2.Length Then
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(","c)
                        row.Item(1) = row.Item(1)
                        Tbl.Rows.Add(row)
                    Else
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(";"c)
                        row.Item(1) = row.Item(1)
                        Tbl.Rows.Add(row)
                    End If
                Else
                    Exit Do
                End If
            Loop

            ''cek dimaster by code
            'If Tbl.Rows.Count > 0 Then
            '    row = Tbl.Rows(0)
            '    If Me.ObjKomp.ApakahSudahAda(row.Item(0)) = True Then
            '        Me.TxtCode.Text = row.Item(0)
            '        Me.TxtPrimary.Text = Me.ObjKomp.GetPrimary(Me.TxtCode.Text)
            '        Me.TxtDesc.Text = Me.ObjKomp.GetDesc(Me.TxtPrimary.Text)
            '        Me.AktifkanCbo()
            '    Else
            '        MsgBox("File csv invalid, Please make sure your data.", MsgBoxStyle.Information, "Information")
            '        Exit Sub
            '    End If
            'End If

            'hapus item dari .csv yang tidak valid/ tidak lengkap
            For i As Integer = Tbl.Rows.Count - 1 To 0 Step -1
                row = Tbl.Rows(i)
                If String.IsNullOrEmpty(row.Item(0)) Or String.IsNullOrEmpty(row.Item(1)) = True Then
                    Tbl.Rows.Remove(row)
                ElseIf String.IsNullOrEmpty(row.Item(2)) = True Then
                    ListKary.Add(row.Item(1))
                    Tbl.Rows.Remove(row)
                End If
            Next

            If Tbl.Rows.Count <= 0 Then
                MsgBox("File csv not found item, Please make sure your data.", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Focus()
                Exit Sub
            End If

            'list yang competence dari csv beda dengan yang di form
            If Tbl.Rows.Count > 0 Then
                For i As Integer = Tbl.Rows.Count - 1 To 0 Step -1
                    row = Tbl.Rows(i)
                    If row.Item(0) <> Me.TxtCode.Text Then
                        If ListCompetenceNotSame.Contains(row.Item(1)) = False Then
                            ListCompetenceNotSame.Add(row.Item(1))
                            Tbl.Rows.Remove(row)
                        End If
                    End If
                Next
            End If

            'kemudian cek juga karyawan yang salah/ tidak aktif
            'jika ada munculkan peringatan dan hapus aja dari tbl 
            For i As Integer = Tbl.Rows.Count - 1 To 0 Step -1
                row = Tbl.Rows(i)
                If Me.ApakahKaryawanAda(row.Item(1)) = False Then
                    If ListKaryNotFound.Contains(row.Item(1)) = False Then
                        ListKaryNotFound.Add(row.Item(1))
                    End If
                    Tbl.Rows.Remove(row)
                End If
            Next

            'cek juga kebenaran type nya
            For i As Integer = Tbl.Rows.Count - 1 To 0 Step -1
                row = Tbl.Rows(i)
                If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                    If String.IsNullOrEmpty(row(2)) = False Then
                        If row(2) = "A" Or row(2) = "B" Or row(2) = "C" Or row(2) = "D" Or row(2) = "E" Or row(2) = "F" Then
                        Else
                            If ListTypeNotFound.Contains(row(1)) = False Then
                                ListTypeNotFound.Add(row(1))
                            End If
                            Tbl.Rows.Remove(row)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                        If IsNumeric(row.Item(2)) = False Then
                            If ListTypeNotFound.Contains(row(1)) = False Then
                                ListTypeNotFound.Add(row(1))
                            End If
                            Tbl.Rows.Remove(row)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                        If IsDate(row(2)) = False Then
                            If ListTypeNotFound.Contains(row(1)) = False Then
                                ListTypeNotFound.Add(row(1))
                            End If
                            Tbl.Rows.Remove(row)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                        'biarin aja karna string/text
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                        If row.Item(2).ToString.ToUpper <> "YES" And row.Item(2).ToString.ToUpper <> "NO" Then
                            If ListTypeNotFound.Contains(row(1)) = False Then
                                ListTypeNotFound.Add(row(1))
                            End If
                            Tbl.Rows.Remove(row)
                        End If
                    End If
                End If
            Next

            If Tbl.Rows.Count > 0 Then
                IsikanDgvImpOpt(FrmImp, Tbl)
                If FrmImp.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    For Each rowD As DataGridViewRow In DgvImp.Rows
                        If DgvImp.Item(FrmImp.ColImp_Import.Index, rowD.Index).Value = True And String.IsNullOrEmpty(rowD.Cells(FrmImp.ColImp_EmpoloyeeName.Index).Value) = False Then
                            Tbl.DefaultView.RowFilter = "[Column2]='" & rowD.Cells(FrmImp.ColImp_EmpoloyeeName.Index).Value & "'"
                            If Tbl.DefaultView.Count > 0 Then
                                row = Tbl.DefaultView(0).Row
                                row.Item(5) = True
                            End If
                        End If
                    Next

                    If Me.ChkShowAll.Checked = True Then
                        For Each rowd As DataGridViewRow In Me.DgvItem.Rows
                            NamaFilter = rowd.Cells(Me.ColItem_EmpName.Index).Value
                            If ListNama_InDGV.Contains(NamaFilter) = False Then
                                ListNama_InDGV.Add(NamaFilter)

                                Tbl.DefaultView.RowFilter = "[Column2] = '" & NamaFilter & "' AND [Column6]= True"
                                If Tbl.DefaultView.Count > 0 Then
                                    row = Tbl.DefaultView(0).Row
                                    If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                        rowd.Cells(Me.ColItem_Value.Index).Value = row(2).ToString
                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                        If IsNumeric(row(2)) = True Then
                                            rowd.Cells(Me.ColItem_Value.Index).Value = Me.ObjCD.DigitGrouping(row(2), Me.DefPrec)
                                        End If
                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                        If IsDate(row(2)) = True Then
                                            rowd.Cells(Me.ColItem_Value.Index).Value = Format(CDate(row.Item(2)), "dd MMM yyyy")
                                        End If
                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                        rowd.Cells(Me.ColItem_Value.Index).Value = row(2).ToString
                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                        rowd.Cells(Me.ColItem_Value.Index).Value = row(2).ToString
                                    End If
                                    If String.IsNullOrEmpty(row(3)) = True Or row(3) = "-" Then
                                        rowd.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                    Else

                                        'cek apakah benar2 tangal
                                        If IsDate(row(3)) = True Then
                                            rowd.Cells(Me.ColItem_NextCertDate.Index).Value = Format(CDate(row(3)), "dd MMM yyyy")
                                        Else
                                            rowd.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                        End If
                                    End If
                                    If String.IsNullOrEmpty(row(4)) = False Then
                                        rowd.Cells(Me.ColItem_Remarks.Index).Value = row(4).ToString
                                    Else
                                        rowd.Cells(Me.ColItem_Remarks.Index).Value = ""
                                    End If
                                    rowd.Cells(Me.ColItem_Modify.Index).Value = row.Item(5)
                                Else
                                    If ListNama_1.Contains(rowd.Cells(Me.ColItem_EmpName.Index).Value) = False Then
                                        ListNama_1.Add(rowd.Cells(Me.ColItem_EmpName.Index).Value)
                                    End If
                                End If
                            End If
                        Next
                    Else
                        'kondisi jika tidak ada item = 0
                        If FrmImp.DgvImport.Rows.Count > 0 Then
                            For Each RowD As DataGridViewRow In FrmImp.DgvImport.Rows
                                With RowD
                                    If .Cells(FrmImp.ColImp_Competence.Index).Value = Me.TxtCode.Text And .Cells(FrmImp.ColImp_Import.Index).Value = True Then
                                        If Me.GetPropetryKary(.Cells(FrmImp.ColImp_EmpoloyeeName.Index).Value, vPrimaryKary, vCode, vNIK, vFullName, vJobPost, vDept) = True Then
                                            If GetJmlDgvBy(vPrimaryKary) = 0 Then
                                                Me.DgvItem.Rows.Add()
                                                DgvR = DgvItem.Rows(DgvItem.Rows.Count - 1)
                                                With DgvR
                                                    .Cells(Me.ColItem_PrimaryEmp.Index).Value = vPrimaryKary
                                                    .Cells(Me.ColItem_EmpCode.Index).Value = vCode
                                                    .Cells(Me.ColItem_EmpName.Index).Value = RowD.Cells(FrmImp.ColImp_EmpoloyeeName.Index).Value
                                                    .Cells(Me.ColItem_EmpNIK.Index).Value = vNIK
                                                    .Cells(Me.ColItem_JobPos.Index).Value = vJobPost
                                                    .Cells(Me.ColItem_Dept.Index).Value = vDept

                                                    If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                                        If IsNumeric(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_Value.Index).Value = Me.ObjCD.DigitGrouping(RowD.Cells(FrmImp.ColImp_Value.Index).Value, Me.DefPrec)
                                                        End If
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                                        If IsDate(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_Value.Index).Value = Format(CDate(RowD.Cells(FrmImp.ColImp_Value.Index).Value), "dd MMM yyyy")
                                                        End If
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    End If
                                                    If String.IsNullOrEmpty(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value) = True Or RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value = "-" Then
                                                        DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                                    Else

                                                        'cek apakah benar2 tangal
                                                        If IsDate(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = Format(CDate(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value), "dd MMM yyyy")
                                                        Else
                                                            DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                                        End If
                                                    End If
                                                    If String.IsNullOrEmpty(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = False Then
                                                        DgvR.Cells(Me.ColItem_Remarks.Index).Value = RowD.Cells(FrmImp.ColImp_Remarks.Index).Value
                                                    Else
                                                        DgvR.Cells(Me.ColItem_Remarks.Index).Value = ""
                                                    End If
                                                    DgvR.Cells(Me.ColItem_Modify.Index).Value = RowD.Cells(FrmImp.ColImp_Import.Index).Value
                                                End With
                                            Else
                                                DgvR = DgvItem.Rows(DgvItem.Rows.Count - 1)
                                                With DgvR
                                                    If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                                                        If IsNumeric(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_Value.Index).Value = Me.ObjCD.DigitGrouping(RowD.Cells(FrmImp.ColImp_Value.Index).Value, Me.DefPrec)
                                                        End If
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                                                        If IsDate(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_Value.Index).Value = Format(RowD.Cells(FrmImp.ColImp_Value.Index).Value, "dd MMM yyyy")
                                                        End If
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                                                        DgvR.Cells(Me.ColItem_Value.Index).Value = RowD.Cells(FrmImp.ColImp_Value.Index).Value
                                                    End If
                                                    If String.IsNullOrEmpty(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value) = True Or RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value = "-" Then
                                                        DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                                    Else

                                                        'cek apakah benar2 tangal
                                                        If IsDate(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value) = True Then
                                                            DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = Format(CDate(RowD.Cells(FrmImp.ColImp_NextCertificateDate.Index).Value), "dd MMM yyyy")
                                                        Else
                                                            DgvR.Cells(Me.ColItem_NextCertDate.Index).Value = "-"
                                                        End If
                                                    End If
                                                    If String.IsNullOrEmpty(RowD.Cells(FrmImp.ColImp_Value.Index).Value) = False Then
                                                        DgvR.Cells(Me.ColItem_Remarks.Index).Value = RowD.Cells(FrmImp.ColImp_Remarks.Index).Value
                                                    Else
                                                        DgvR.Cells(Me.ColItem_Remarks.Index).Value = ""
                                                    End If
                                                    DgvR.Cells(Me.ColItem_Modify.Index).Value = RowD.Cells(FrmImp.ColImp_Import.Index).Value
                                                End With
                                            End If

                                        End If
                                    End If
                                End With
                            Next
                        End If
                    End If

                    Me.ObjCD.SetNumber(Me.DgvItem, False)

                    Dim MsgR As MsgBoxResult = MsgBoxResult.Yes
                    Dim Ic As Integer = 0
                    MsgR = MsgBox("File import success, Do you want continue to save?", MsgBoxStyle.YesNo, "Confirm")
                    If MsgR = MsgBoxResult.Yes Then
                        For Each rowDGV As DataGridViewRow In Me.DgvItem.Rows
                            Tbl.DefaultView.RowFilter = ""
                            Tbl.DefaultView.RowFilter = "[Column2]='" & rowDGV.Cells(Me.ColItem_EmpName.Index).Value & "' AND [Column6]=true"
                            If Tbl.DefaultView.Count > 0 Then
                                If rowDGV.Cells(Me.ColItem_Modify.Index).Value = True Then
                                    With rowDGV
                                        StrValue = ""
                                        If Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                            StrValue = .Cells(Me.ColItem_Value.Index).Value
                                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Number Then
                                            StrValue = .Cells(Me.ColItem_Value.Index).Value
                                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Tanggal Then
                                            StrValue = .Cells(Me.ColItem_Value.Index).Value
                                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.Text Then
                                            StrValue = .Cells(Me.ColItem_Value.Index).Value
                                        ElseIf Me.cboDataType.SelectedIndex = ClsMasterKompetensi.EnumDataType.YN Then
                                            StrValue = .Cells(Me.ColItem_Value.Index).Value
                                        End If
                                        Dim DateNextCert As Date = #1/1/1950#
                                        If IsDate(.Cells(Me.ColItem_NextCertDate.Index).Value) = True Then
                                            DateNextCert = .Cells(Me.ColItem_NextCertDate.Index).Value
                                        End If

                                        If .Cells(Me.ColItem_Modify.Index).Value = True Then
                                            'new 
                                            If Me.GetJmlKomp(rowDGV.Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text) <= 0 Then
                                                DTA.InsertQuery(.Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text, StrValue, .Cells(Me.ColItem_Remarks.Index).Value, DateNextCert)
                                                Ic += 1
                                            Else
                                                'update
                                                DTA.UpdateQuery(StrValue, .Cells(Me.ColItem_Remarks.Index).Value, DateNextCert, .Cells(Me.ColItem_PrimaryEmp.Index).Value, Me.TxtPrimary.Text)
                                                Ic += 1
                                            End If
                                        End If
                                    End With
                                End If
                            End If
                        Next
                        If Ic > 0 Then
                            MsgBox("Save successed", MsgBoxStyle.Information, "Information")
                        End If
                    Else
                        Me.IsikanData()
                        Exit Sub
                    End If

                    'Exit Sub
                Else
                    Exit Sub
                End If
            End If

            '' __. INCLUDE BATASAN IMPORT DI CODING // 
            '1 SAAT BEDA TYPE 
            '2 SAAT BEDA KOMPETENCE
            '3 SAAT DATA CSV TIDAK LENGKAP
            '4 SAAT NAMA KARAYWAN SAMA
            '5 SAAT FILE CSV KOSONG ATAU CORUPT
            '6 MSGBOX PERINGATAN ATAU INFORMASI MENGENAI : 
            '- IMPORT SUKSES APAKAH AKAN DISIMPAN
            '- POINT2 YANG SEBELUMNYA 1-5

            If ListNama_1.Count > 0 Then
                ListNama_1.Sort()
                For Each NamaKary As String In ListNama_1
                    If StrMsgBox_1 = "" Then
                        StrMsgBox_1 = "Employee not found in csv file." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_1 &= ", " & NamaKary
                    End If
                Next
            End If

            If ListNama_2.Count > 0 Then
                ListNama_2.Sort()
                For Each NamaKary As String In ListNama_2
                    If StrMsgBox_2 = "" Then
                        StrMsgBox_2 = "Employee not found in item" & vbCrLf & NamaKary
                    Else
                        StrMsgBox_2 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_2 &= vbCrLf
            End If

            If ListKary.Count > 0 Then
                ListKary.Sort()
                For Each NamaKary As String In ListKary
                    If StrMsgBox_3 = "" Then
                        StrMsgBox_3 = "Value not found in csv file." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_3 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_3 &= vbCrLf
            End If

            If ListCompetenceNotSame.Count > 0 Then
                ListCompetenceNotSame.Sort()
                For Each NamaKary As String In ListCompetenceNotSame
                    If StrMsgBox_4 = "" Then
                        StrMsgBox_4 = "Employee not same data competence with the form." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_4 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_4 &= vbCrLf
            End If

            If ListKaryNotFound.Count > 0 Then
                ListKaryNotFound.Sort()
                For Each NamaKary As String In ListKaryNotFound
                    If StrMsgBox_5 = "" Then
                        StrMsgBox_5 = "Employee Name invalid." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_5 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_5 &= vbCrLf
            End If

            If ListTypeNotFound.Count > 0 Then
                ListTypeNotFound.Sort()
                For Each NamaKary As String In ListTypeNotFound
                    If StrMsgBox_6 = "" Then
                        StrMsgBox_6 = "Employee with data type invalid, can't import." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_6 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_6 &= vbCrLf
            End If

            If StrMsgBox_1 <> "" Or StrMsgBox_2 <> "" Or StrMsgBox_3 <> "" Or StrMsgBox_4 <> "" Or StrMsgBox_5 <> "" Then
                If StrMsgBox_1 <> "" Then
                    StrMsgBox_1 &= vbCrLf & vbCrLf
                End If
                If StrMsgBox_3 <> "" Then
                    StrMsgBox_3 &= vbCrLf & vbCrLf
                End If
                If StrMsgBox_4 <> "" Then
                    StrMsgBox_4 &= vbCrLf & vbCrLf
                End If
                MsgBox(StrMsgBox_1 & StrMsgBox_2 & StrMsgBox_3 & StrMsgBox_4 & StrMsgBox_5 & StrMsgBox_6, MsgBoxStyle.Information, "Information")
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Openfile(ByRef vFileExtStr As String) As Boolean
        Dim FileExtStr As String = ""

        Try
            '1. cari file csv format infinity
            Me.OpenDia.Title = "Please specify file"
            Me.OpenDia.Filter = "csv files (*.csv)|*.csv"
            'Me.OpenFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.OpenDia.FilterIndex = 1
            Me.OpenDia.DefaultExt = "csv"
            Me.OpenDia.CheckFileExists = True
            Me.OpenDia.Multiselect = False
            Me.OpenDia.FileName = ""
            If Me.OpenDia.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Return False
            End If
            If Me.OpenDia.FileName = "" Then
                Return False
            End If

            Me.LocationFileSumber = Me.OpenDia.FileName
            vFileExtStr = System.IO.Path.GetExtension(Me.OpenDia.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub Rec()
    '    Dim Row As DseMasterKaryawan_Kompetensi_Upd.GRecordRow
    '    Try
    '        TBLRec.Rows.Clear()
    '        For Each rowD As DataGridViewRow In Me.DgvItem.Rows
    '            If rowD.Cells(Me.ColItem_Modify.Index).Value = True Then
    '                Row = TBLRec.NewGRecordRow
    '                With rowD
    '                    Row.Primary_Emp = .Cells(Me.ColItem_PrimaryEmp.Index).Value
    '                    Row.Modify = .Cells(Me.ColItem_Modify.Index).Value
    '                End With
    '                TBLRec.AddGRecordRow(Row)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub UnRec(ByVal ie As Integer)
    '    Dim Row As DseMasterKaryawan_Kompetensi_Upd.GRecordRow
    '    Try
    '        For Each rowD As DataGridViewRow In Me.DgvItem.Rows
    '            If rowD.Index <> ie Then
    '                TBLRec.DefaultView.RowFilter = "[Primary Emp]=" & rowD.Cells(Me.ColItem_PrimaryEmp.Index).Value
    '                If TBLRec.DefaultView.Count > 0 Then
    '                    Row = TBLRec.DefaultView(0).Row
    '                    rowD.Cells(Me.ColItem_Modify.Index).Value = Row.Modify
    '                End If
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Function GetRefreshTypeValue() As Boolean
        Dim TypeValueDB As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
        Try
            TypeValueDB = Me.ObjKomp.GetDataType(Me.TxtPrimary.Text)
            If TypeValueDB <> Me.cboDataType.SelectedIndex Then
                Me.AktifkanCbo()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbChkAllExp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbChkAllExp.Click
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                rowD.Cells(Me.ColItem_Export.Index).Value = True
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUnchkAllExp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbUnchkAllExp.Click
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                rowD.Cells(Me.ColItem_Export.Index).Value = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDgvImpOpt(ByVal Hj As FrmMasterKaryawan_Kompetensi_Upd_ImportOpt, ByVal Tbl As DataTable)
        Dim Dgv As DataGridView = Hj.DgvImport
        Dim row As DataRow
        Dim rowD As DataGridViewRow
        Try
            Dgv.Rows.Clear()
            For Each row In Tbl.Rows
                Dgv.Rows.Add()
                rowD = Dgv.Rows(Dgv.Rows.Count - 1)
                With rowD
                    .Cells(Hj.ColImp_Competence.Index).Value = row.Item(0)
                    .Cells(Hj.ColImp_EmpoloyeeName.Index).Value = row.Item(1) 'ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper

                    If Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Alfabetic Then
                        If String.IsNullOrEmpty(row.Item(2)) = False Then
                            .Cells(Hj.ColImp_Value.Index).Value = row.Item(2)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Number Then
                        If IsNumeric(row.Item(2)) = True Then
                            If Me.ObjCD.CekBilanganPositif(row.Item(2)) = False Then
                                .Cells(Hj.ColImp_Value.Index).Value = ""
                            Else
                                .Cells(Hj.ColImp_Value.Index).Value = Me.ObjCD.DigitGrouping(row.Item(2), Me.DefPrec)
                            End If
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Tanggal Then
                        If IsDate(row.Item(2)) = True Then
                            .Cells(Hj.ColImp_Value.Index).Value = row.Item(2)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.Text Then
                        If String.IsNullOrEmpty(row.Item(2)) = False Then
                            .Cells(Hj.ColImp_Value.Index).Value = row.Item(2)
                        End If
                    ElseIf Me.cboDataType.SelectedIndex = Core_Module.ClsMasterKompetensi.EnumDataType.YN Then
                        If row.Item(2).ToString.ToUpper = "YES" Or row.Item(2).ToString.ToUpper = "NO" Then
                            .Cells(Hj.ColImp_Value.Index).Value = row.Item(2)
                        End If
                    End If

                    If IsDate(row.Item(3)) = False Then
                        .Cells(Hj.ColImp_NextCertificateDate.Index).Value = "-"
                    Else
                        If IsDate(row.Item(3)) = True Then
                            rowD.Cells(Hj.ColImp_NextCertificateDate.Index).Value = Format(CDate(row.Item(3)), "dd MMM yyyy")
                        End If
                    End If
                    .Cells(Hj.ColImp_Remarks.Index).Value = row.Item(4)
                    .Cells(Hj.ColImp_Import.Index).Value = True
                End With
            Next
            Me.ObjCD.SetNumber(Dgv)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPropetryKary(ByVal Karyawan As String, ByRef vPrimaryKary As Double, ByRef Code As String, ByRef NIK As String, ByRef Name As String, ByRef JobPost As String, ByRef Dept As String) As Boolean
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.DataTable3TableAdapter
        Dim TblPro As New DseMasterKaryawan_Kompetensi_Upd.DataTable3DataTable
        Dim Row As DseMasterKaryawan_Kompetensi_Upd.DataTable3Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TblPro, Karyawan)
            For Each Row In TblPro
                vPrimaryKary = Row.Primary
                Code = Row.Code
                NIK = Row.NIK
                Name = Row.Full_Name
                JobPost = Row.Post_Desc
                Dept = Row.Dept_Desc
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlDgvBy(ByVal vPrimary As String) As Integer
        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                With rowD
                    If .Cells(Me.ColItem_PrimaryEmp.Index).Value = vPrimary Then
                        Return 1
                    End If
                End With
            Next
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlKomp(ByVal vPrimaryKary As Double, ByVal vPrimaryKomp As Double) As Integer
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByKaryKomp(vPrimaryKary, vPrimaryKomp)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return Retval.Value
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahKaryawanAda(ByVal vfullName As String) As Boolean
        Dim DTA As New DseMasterKaryawan_Kompetensi_UpdTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim DataLong As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DataLong = DTA.ScalarQuery_GetJmlByKaryawan(vfullName)
            If DataLong.HasValue = True Then
                If DataLong.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class