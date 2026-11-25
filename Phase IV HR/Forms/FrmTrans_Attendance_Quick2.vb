Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms
Public Class FrmTrans_Attendance_Quick2
    Public User As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim TBL_JP As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable
    Dim TBL_Dept As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable
    Dim TBL_Formula As New Core_Module.DseMasterWorkshift_Allocation.tblmaster_salary_formula1DataTable
    Dim ObjQuick As New ClsTrans_Attendance_quick
    Dim EraseRowIndex As Integer = 0
    Dim AllData As Boolean = True
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Private Function GetIndexItemInput(ByVal vPrimary As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvEmployee.Rows
                If rowDGV.Cells(Me.ColPrimaryCode.Index).Value = vPrimary Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim TBLMsgBox As New DseTrans_Attendance.TBLMessageBoxDataTable
        Try
            Me.dtpPeriod.Value = Me.dtpPeriod.Value.Date
            Me.DtpTo.Value = Me.DtpTo.Value.Date

            If Me.dtpPeriod.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Start date should not be more than end date", MsgBoxStyle.Information, "Error")
            Else

                Dim ObjLeave As New Core_Module.ClsMasterCuti
                Dim ObjCd As New Core_Module.ClsCekData
                If ObjLeave.IsAlreadyDefaultLeave(ObjCd.GetFisrtDateOfTheYear(Me.DtpTo.Value)) = False Then
                    MsgBox("There is no default leave settings in this period, please define first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Dim IsImportFromMachine As Boolean = False
                If Me.CboImportFromMachine.Checked = True Then
                    IsImportFromMachine = True
                End If

                'aul tambahan
                Dim IsContinueFromPrevAtt As Boolean = False
                If Me.ChkAttPrev.Checked = True Then
                    IsContinueFromPrevAtt = True
                End If
                Me.ObjQuick.Simpan(Me, TBLMsgBox, Me.TBL_JP, Me.TBL_Dept, IsImportFromMachine, IsContinueFromPrevAtt)
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLMsgBox.Dispose()
        End Try
    End Sub

    Private Sub IsiListBoxJobPosition(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As Core_Module.DseMasterWorkshift_Allocation.tblmaster_jobpositionDataTable)
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Position"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub IsiListBoxDepartemen(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As Core_Module.DseMasterWorkshift_Allocation.tblmaster_departemenDataTable)
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_departemenTableAdapter
        Try
            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Departement"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiListBoxFormula(ByVal Lbo As System.Windows.Forms.CheckedListBox, ByVal TBL As Core_Module.DseMasterWorkshift_Allocation.tblmaster_salary_formula1DataTable)
        Dim DTA As New Core_Module.DseMasterWorkshift_AllocationTableAdapters.tblmaster_salary_formula1TableAdapter
        Try
            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            Lbo.DataSource = Nothing
            Lbo.DataSource = TBL.DefaultView

            'Isikan
            Lbo.DisplayMember = "Formula"
            For i As Integer = 0 To Lbo.Items.Count - 1
                Lbo.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Quick2_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Quick2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            Me.IsiListBoxJobPosition(Me.ClbPosition, Me.TBL_JP)
            Me.IsiListBoxDepartemen(Me.ClbDepartement, Me.TBL_Dept)
            Me.IsiListBoxFormula(Me.clbSalaryFormula, Me.TBL_Formula)
            Me.RboAll.Checked = True
            Me.ClbPosition.Enabled = False
            Me.clbSalaryFormula.Enabled = False
            Me.ClbDepartement.Enabled = False
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvEmployee.Enabled = False
            Me.SplitContainer4.Panel1Collapsed = True
            Me.AllData = True
            'Tambahan koding quick
            Me.btnCancel.Enabled = False
            Me.btnCancel.Visible = False
            Me.BtnPause.Enabled = False
            Me.BtnPause.Visible = False
            Me.LblProgress.Text = "......................................................"
            Me.CheckForIllegalCrossThreadCalls = False

            Dim UserSettingStr As String = ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, ClsTransaksiSettingUser.EnumNamaParameter.HR_Attendance_MergeAttendance, Me.User)
            If UserSettingStr <> "" Then
                If UserSettingStr = "True" Then
                    Me.ChkAttPrev.Checked = True
                Else
                    Me.ChkAttPrev.Checked = False
                End If
            Else
                Me.ChkAttPrev.Checked = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboAll.CheckedChanged
        Try
            Me.chkJobPosition.Enabled = False
            Me.chkSalaryFormula.Enabled = False
            Me.ChkDepartment.Enabled = False
            Me.ChkEmployee.Enabled = False

            Me.chkJobPosition.Checked = False
            Me.chkSalaryFormula.Checked = False
            Me.ChkDepartment.Checked = False
            Me.ChkEmployee.Checked = False
            If Me.RboAll.Checked = True Then
                Me.AllData = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboSelected_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboSelected.CheckedChanged
        Try
            Me.chkJobPosition.Enabled = True
            Me.chkSalaryFormula.Enabled = True
            Me.ChkDepartment.Enabled = True
            Me.ChkEmployee.Enabled = True
            If Me.RboSelected.Checked = True Then
                Me.AllData = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkJobPosition_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkJobPosition.CheckedChanged
        Try
            If Me.chkJobPosition.Checked = True Then
                Me.ClbPosition.Enabled = True
            Else
                Me.ClbPosition.Enabled = False
                Dim i As Integer = 1
                For i = 1 To Me.ClbPosition.Items.Count
                    Me.ClbPosition.SetItemChecked(i - 1, True)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSalaryFormula_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSalaryFormula.CheckedChanged
        Try
            If Me.chkSalaryFormula.Checked = True Then
                Me.clbSalaryFormula.Enabled = True
            Else
                Me.clbSalaryFormula.Enabled = False
                Dim i As Integer = 1
                For i = 1 To Me.clbSalaryFormula.Items.Count
                    Me.clbSalaryFormula.SetItemChecked(i - 1, True)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkDepartment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDepartment.CheckedChanged
        Try
            If Me.ChkDepartment.Checked = True Then
                Me.ClbDepartement.Enabled = True
            Else
                Me.ClbDepartement.Enabled = False
                Dim i As Integer = 1
                For i = 1 To Me.ClbDepartement.Items.Count
                    Me.ClbDepartement.SetItemChecked(i - 1, True)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkEmployee_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkEmployee.CheckedChanged
        Try
            If Me.ChkEmployee.Checked = True Then
                Me.dgvEmployee.Enabled = True
            Else
                Me.dgvEmployee.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If Me.ChkEmployee.Enabled = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim dta As New DseTrans_AttendanceTableAdapters.DataTable9TableAdapter
        Dim tbl As New DseTrans_Attendance.DataTable9DataTable
        Dim row As DseTrans_Attendance.DataTable9Row
        Dim primary As Double = -1
        Dim index As Integer = 0
        Try
            If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColCode.Index Then
                    If String.IsNullOrEmpty(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primary = Me.ObjEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemInput(primary)
                            If index < 0 Then
                                dta.FillBy_PrimaryEmp(tbl, primary)
                                For Each row In tbl.Rows
                                    Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value = primary
                                    Me.dgvEmployee.Item(Me.ColCode.Index, e.RowIndex).Value = row.Emp_Code
                                    Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = row.Emp_Name
                                    Me.dgvEmployee.Item(Me.ColNIK.Index, e.RowIndex).Value = row.NIK
                                    Me.dgvEmployee.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = row.Job_Primary
                                    Me.dgvEmployee.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = row.Dept_Primary
                                    Me.dgvEmployee.Item(Me.ColPrimarySF.Index, e.RowIndex).Value = row.SF_Primary
                                Next
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            index = Me.GetIndexItemInput(ij.PrimaryEmploye)
                            If ij.PrimaryEmploye = -1 Then
                                'hapus baris
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                If index < 0 Then
                                    dta.FillBy_PrimaryEmp(tbl, ij.PrimaryEmploye)
                                    For Each row In tbl.Rows
                                        Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.dgvEmployee.Item(Me.ColCode.Index, e.RowIndex).Value = row.Emp_Code
                                        Me.dgvEmployee.Item(Me.ColName.Index, e.RowIndex).Value = row.Emp_Name
                                        Me.dgvEmployee.Item(Me.ColNIK.Index, e.RowIndex).Value = row.NIK
                                        Me.dgvEmployee.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = row.Job_Primary
                                        Me.dgvEmployee.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = row.Dept_Primary
                                        Me.dgvEmployee.Item(Me.ColPrimarySF.Index, e.RowIndex).Value = row.SF_Primary
                                    Next
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                        'hapus baris
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        'hapus baris
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            Else
                If IsNothing(Me.dgvEmployee.Item(Me.ColPrimaryCode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCode.Index Then
                        Me.EraseRowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvEmployee.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_Position_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_checked.Click
        Try
            For i As Integer = 0 To Me.ClbPosition.Items.Count - 1
                Me.ClbPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Position_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_JP_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbPosition.Items.Count - 1
                Me.ClbPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_Departemen_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_Dept_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_SF_Checked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_SF_Checked.Click
        Try
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub Lbl_SF_UnChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lbl_SF_Unchecked.Click
        Try
            For i As Integer = 0 To Me.ClbSalaryFormula.Items.Count - 1
                Me.ClbSalaryFormula.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub
End Class