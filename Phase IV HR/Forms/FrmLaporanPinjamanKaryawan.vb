Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class FrmLaporanPinjamanKaryawan
    Public UserCode As String = ""
    Public BackColorAll As color = Core_Module.GetSMESSBackColorAll
    Dim Buffer As String = ""
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjLoan As New HR.ClsTransaksiLoan
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim IndexLoan As Integer = 0
    Dim DefPrec As Decimal = 0
    Dim vStart As Date = #1/1/1950#
    Dim vEnd As Date = #1/1/2200#

    Private Sub FrmLaporanPinjamanKaryawan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboGrouping.SelectedIndex = 0
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvDocNo.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.cboFormat.SelectedIndex = 1
            Me.CancelButton = Me.btnClose
            Me.ChkFilterDate.Checked = False
            Me.dtpFrom.Enabled = False
            Me.dtpTo.Enabled = False
            Me.chkOnlyShowOutStandingBalance.Checked = True
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

    Private Function GetIndexItemEmployee(ByVal vPrimary As Double, _
           ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvEmployee.Rows.Count - 1
                RowDGV = Me.dgvEmployee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColEmployee_Primary.Index).Value = vPrimary Then
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

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryEmployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryEmployee = Me.ObjEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemEmployee(PrimaryEmployee, e.RowIndex)
                            If index < 0 Then
                                Me.dgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = PrimaryEmployee
                                Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(PrimaryEmployee)
                                Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(PrimaryEmployee)
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
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexEmp = e.RowIndex
                                            BgwEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.dgvEmployee.Item(Me.ColEmployee_Primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
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
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                                If Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        Me.BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                    End If
                Else
                    If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.IndexEmp = e.RowIndex
                        Me.BgwEmployee.RunWorkerAsync()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvEmployee.Columns(Me.ColEmployee_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvEmployee.Columns(Me.ColEmployee_Name.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvEmployee.Rows(Me.dgvEmployee.Rows.Count - 1)
            dr.Cells(Me.ColEmployee_Primary.Index).Value = -1
            dr.Cells(Me.ColEmployee_Code.Index).Value = ""
            dr.Cells(Me.ColEmployee_Name.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub dgvDepartment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDepartment.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvDepartment.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDept As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
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
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterDepartementCari
                            ij.txtCari.Text = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryDepartement <> -1 Then
                                index = Me.GetIndexItemDepartment(ij.PrimaryDepartement, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                    Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                    Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexDept = e.RowIndex
                                            Me.BgwDepartment.RunWorkerAsync()
                                        Else
                                            Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvDepartment)
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
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                                If Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvDepartment)
                    End If
                Else
                    If e.ColumnIndex = Me.ColDeptCode.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.IndexDept = e.RowIndex
                        Me.BgwDepartment.RunWorkerAsync()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvDepartment.Columns(Me.ColDeptCode.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvDepartment.Columns(Me.ColDeptDesc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub dgvDepartment_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvDepartment.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvDepartment.Rows(Me.dgvDepartment.Rows.Count - 1)
            dr.Cells(Me.ColDeptPrimary.Index).Value = -1
            dr.Cells(Me.ColDeptCode.Index).Value = ""
            dr.Cells(Me.ColDeptDesc.Index).Value = ""
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

    Private Sub dgvJobPosition_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvJobPosition.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvJobPosition.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryJobPost As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
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
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterJobPosition_Cari
                            ij.txtCari.Text = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemJobPosition(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexJobPost = e.RowIndex
                                            Me.BgwJobPosition.RunWorkerAsync()
                                        Else
                                            Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvJobPosition)
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
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                                If Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                    End If
                Else
                    If e.ColumnIndex = Me.ColJobPosisitonCode.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.IndexJobPost = e.RowIndex
                        Me.BgwJobPosition.RunWorkerAsync()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvJobPosition.Columns(Me.ColJobPosisitonCode.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.dgvJobPosition.Columns(Me.ColDeptDesc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub dgvJobPosition_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvJobPosition.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvJobPosition.Rows(Me.dgvJobPosition.Rows.Count - 1)
            dr.Cells(Me.ColJobPosisitonPrimary.Index).Value = -1
            dr.Cells(Me.ColJobPosisitonCode.Index).Value = ""
            dr.Cells(Me.ColJobPosisitonDesc.Index).Value = ""
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

    Private Sub DgvDocNo_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvDocNo.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                    Me.Buffer = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                    Me.Buffer = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryByDoc(ByVal vDoc As String) As Double
        Dim DTA As New DseTransaksiLoanTableAdapters.tbltrans_loanTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryByDocNo(vDoc)
            If Retval.HasValue = True Then
                If Retval.Value <> -1 Then
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

    Private Sub DgvDocNo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDocNo.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvDocNo.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDocNo As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjLoan.ApakahSudahAda(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryDocNo = Me.GetPrimaryByDoc(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemDocNo(PrimaryDocNo, e.RowIndex)
                            If index < 0 Then
                                Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = PrimaryDocNo
                                Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjLoan.GetDocNoByPrimary(PrimaryDocNo)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication document, please input another document.", MsgBoxStyle.Information, "Information")
                                    Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexLoan = e.RowIndex
                                        Me.BgwDoc.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDocNo)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New HR.FrmTransaksiLoan_Cari
                            ij.AllFilterData = True
                            ij.txtCari.Text = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryTransaksi <> -1 Then
                                index = Me.GetIndexItemDocNo(ij.PrimaryTransaksi, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = ij.PrimaryTransaksi
                                    Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjLoan.GetDocNoByPrimary(ij.PrimaryTransaksi)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication document, please input another document.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexLoan = e.RowIndex
                                            Me.BgwDoc.RunWorkerAsync()
                                        Else
                                            Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvDocNo)
                                        Exit Sub
                                    Else
                                        Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = ij.PrimaryTransaksi
                                        Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjLoan.GetDocNoByPrimary(ij.PrimaryTransaksi)
                                    End If
                                End If
                            ElseIf ij.PrimaryTransaksi = -1 Then
                                If Me.DgvDocNo.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexLoan = e.RowIndex
                                        Me.BgwDoc.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDocNo)
                                    Exit Sub
                                End If
                                If Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexLoan = e.RowIndex
                                        Me.BgwDoc.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvDocNo)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvDocNo)
                    End If
                Else
                    If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.IndexLoan = e.RowIndex
                        Me.BgwDoc.RunWorkerAsync()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvDocNo.Columns(Me.ColDocNo_Document.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvDocNo.Columns(Me.ColDeptDesc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub DgvDocNo_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvDocNo.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvDocNo.Rows(Me.DgvDocNo.Rows.Count - 1)
            dr.Cells(Me.ColDocNo_Primary.Index).Value = -1
            dr.Cells(Me.ColDocNo_Document.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemDocNo(ByVal vPrimary As Double, _
        ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvDocNo.Rows.Count - 1
                RowDGV = Me.DgvDocNo.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColDocNo_Primary.Index).Value = vPrimary Then
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

    Private Sub BgwDoc_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDoc.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDoc_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDoc.RunWorkerCompleted
        Try
            If DgvDocNo.RowCount > 0 Then
                Me.DgvDocNo.Rows.RemoveAt(Me.IndexLoan)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DefPrec = Me.ObjCurrency.GetDefaultPrecission(1)

            If Me.dtpFrom.Enabled = True Then
                If Me.dtpTo.Value.Date < Me.dtpFrom.Value.Date Then
                    MsgBox("To date must be > with from date.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Me.vStart = Me.dtpFrom.Value.Date
                Me.vEnd = Me.dtpTo.Value.Date
            Else
                Me.vStart = #1/1/1950#
                Me.vEnd = #1/1/5000#
            End If
            If Me.cboFormat.SelectedIndex = 0 Then
                CetakLaporan_Detail(Me.crvLaporanKontrakKaryawan, vStart, vEnd)
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                CetakLaporan_Rekap(Me.crvLaporanKontrakKaryawan, vStart, vEnd)
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function CetakLaporan_Rekap(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim Ds As New DseLaporanPinjamanKaryawan
        Dim Report1 As New CrpLaporanPinjamanKaryawan_Rekap1
        Dim DtaRekap As New DseLaporanPinjamanKaryawanTableAdapters.TblRekapTableAdapter
        Dim TblRekap As DseLaporanPinjamanKaryawan.TblRekapDataTable = Ds.TblRekap
        Dim RowRekap As DseLaporanPinjamanKaryawan.TblRekapRow
        Dim IntRpt As Integer = 0
        Try
            TblRekap.Rows.Clear()
            DtaRekap.Connection.ConnectionString = Core_Module.GetConstring
            DtaRekap.Fill(TblRekap, Me.vStart, Me.vEnd)
            Me.Filtering(TblRekap)
            For Each RowRekap In TblRekap
                With RowRekap
                    If .Loan_Type = 0 Then
                        .StrType = "No Rate"
                    Else
                        IntRpt += 1
                        Dim StrLoanType As String = ""
                        If .Loan_Type = ClsTransaksiLoan.EnumLoanType.Annuity Then
                            StrLoanType = "Annuity"
                        ElseIf .Loan_Type = ClsTransaksiLoan.EnumLoanType.Flat Then
                            StrLoanType = "Flat"
                        ElseIf .Loan_Type = ClsTransaksiLoan.EnumLoanType.Effective Then
                            StrLoanType = "Effective"
                        End If
                        .StrType = StrLoanType & "(" & Me.ObjCD.RemoveDecimalZero(.Persen_Rate) & "%)"
                    End If
                    .IntPayToGain_OrLoss = .Payment_Amount - .Installment_Amount2
                    If IsNumeric(.Installment_Period_OF) = False Then
                        .Installment_Period_OF = 0
                    End If
                    If String.IsNullOrEmpty(CStr(.Installment_Period_OF)) = True Then
                        .Installment_Period_OF = 0
                    End If
                    .StrInstallment = .Installment_Period_OF & " Of " & RowRekap.Installment_To
                    .Balance_Base = .Loan_Amount - .Base2
                    .Balance = .Loan_Amount + .Interest - .Installment_Amount2
                    .DefPrec = Me.DefPrec
                    If Me.CboGrouping.SelectedIndex = 1 Then
                        .Group_By = .Departemen_Code
                    End If
                End With
            Next

            If Me.chkOnlyShowOutStandingBalance.Checked = True Then
                TblRekap.DefaultView.RowFilter = "[Balance Base]=0"
                For ij As Integer = TblRekap.DefaultView.Count - 1 To 0 Step -1
                    RowRekap = TblRekap.DefaultView(ij).Row
                    TblRekap.Rows.Remove(RowRekap)
                Next
            End If

            Me.GetSorting1(TblRekap)
            Report1.SetDataSource(Ds)
            vCRV.ReportSource = Report1
            vCRV.Show()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CetakLaporan_Detail(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim Ds As New DseLaporanPinjamanKaryawan
        Dim Report As New CrpLaporanPinjamanKaryawan_Detail1
        Dim DtaDetail As New DseLaporanPinjamanKaryawanTableAdapters.TblDetailTableAdapter
        Dim TblDetail As DseLaporanPinjamanKaryawan.TblDetailDataTable = Ds.TblDetail
        Dim RowDetail As DseLaporanPinjamanKaryawan.TblDetailRow
        Try
            TblDetail.Rows.Clear()
            DtaDetail.Connection.ConnectionString = Core_Module.GetConstring
            DtaDetail.Fill(TblDetail, Me.vStart, Me.vEnd)
            Me.Filtering(TblDetail)
            For Each RowDetail In TblDetail
                With RowDetail
                    If .Loan_Type = 0 Then
                        .StrType = "No Rate"
                    Else
                        Dim StrLoanType As String = ""
                        If .Loan_Type = ClsTransaksiLoan.EnumLoanType.Annuity Then
                            StrLoanType = "Annuity"
                        ElseIf .Loan_Type = ClsTransaksiLoan.EnumLoanType.Flat Then
                            StrLoanType = "Flat"
                        ElseIf .Loan_Type = ClsTransaksiLoan.EnumLoanType.Effective Then
                            StrLoanType = "Effective"
                        End If
                        .StrType = StrLoanType & "(" & Me.ObjCD.RemoveDecimalZero(.Persen_Rate) & "%)"
                    End If

                    If .Payment_Amount > 0 Then
                        .IntPayToGain_OrLoss = .Payment_Amount - .Installment_Amount
                    End If

                    If .Payment_Date.Date = CDate("1/1/1950") Then
                        .Payment_Date = .Date_Schedule
                    End If
                    .Balance = .Loan_Amount + .Interest2 - .Installment_Amount2
                    .Total_Balance_Base = .Loan_Amount - .Base2
                    If Me.CboGrouping.SelectedIndex = 1 Then
                        .Group_By = .Departemen_Code
                    End If
                    .DefPrec = Me.DefPrec
                End With
            Next

            If Me.chkOnlyShowOutStandingBalance.Checked = True Then
                TblDetail.DefaultView.RowFilter = "[Total Balance Base]=0"
                For ij As Integer = TblDetail.DefaultView.Count - 1 To 0 Step -1
                    RowDetail = TblDetail.DefaultView(ij).Row
                    TblDetail.Rows.Remove(RowDetail)
                Next
            End If
            Me.GetSorting(TblDetail)
            Report.SetDataSource(Ds)
            vCRV.ReportSource = Report
            vCRV.Show()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaDetail.Dispose()
        End Try
    End Function

    Public Function GetSorting(ByRef TBL As DseLaporanPinjamanKaryawan.TblDetailDataTable) As Boolean
        Dim row As DseLaporanPinjamanKaryawan.TblDetailRow
        Dim TBLSort As New DseLaporanPinjamanKaryawan.TblDetailDataTable
        Dim ROWSort As DseLaporanPinjamanKaryawan.TblDetailRow
        Dim SortBy As String
        Try
            If Me.CboGrouping.SelectedIndex = 0 Then
                SortBy = "[Kode Karyawan] ASC, [Date] ASC, [Document No] ASC, [Installment Period] ASC"
            Else
                SortBy = "[Departemen Code] ASC, [Kode Karyawan] ASC, [Date] ASC, [Document No] ASC, [Installment Period] ASC"
            End If
            For Each row In TBL
                ROWSort = TBLSort.NewTblDetailRow
                For i As Integer = 0 To TBL.Columns.Count - 1
                    ROWSort.Item(i) = row.Item(i)
                Next
                TBLSort.Rows.Add(ROWSort)
            Next

            TBL.Rows.Clear()
            TBLSort.DefaultView.Sort = SortBy
            For i As Integer = 0 To TBLSort.DefaultView.Count - 1
                ROWSort = TBLSort.DefaultView(i).Row
                row = TBL.NewTblDetailRow
                For i2 As Integer = 0 To TBLSort.Columns.Count - 1
                    row.Item(i2) = ROWSort.Item(i2)
                Next
                TBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Public Function GetSorting1(ByRef TBL As DseLaporanPinjamanKaryawan.TblRekapDataTable) As Boolean
        Dim row As DseLaporanPinjamanKaryawan.TblRekapRow
        Dim TBLSort As New DseLaporanPinjamanKaryawan.TblRekapDataTable
        Dim ROWSort As DseLaporanPinjamanKaryawan.TblRekapRow
        Dim SortBy As String
        Try
            If Me.CboGrouping.SelectedIndex = 0 Then
                SortBy = "[Kode Karyawan] ASC, [Date] ASC, [Document No] ASC, [Installment Period] ASC"
            Else
                SortBy = "[Departemen Code] ASC, [Kode Karyawan] ASC, [Date] ASC, [Document No] ASC, [Installment Period] ASC"
            End If
            For Each row In TBL
                ROWSort = TBLSort.NewTblRekapRow
                For i As Integer = 0 To TBL.Columns.Count - 1
                    ROWSort.Item(i) = row.Item(i)
                Next
                TBLSort.Rows.Add(ROWSort)
            Next

            TBL.Rows.Clear()
            TBLSort.DefaultView.Sort = SortBy
            For i As Integer = 0 To TBLSort.DefaultView.Count - 1
                ROWSort = TBLSort.DefaultView(i).Row
                row = TBL.NewTblRekapRow
                For i2 As Integer = 0 To TBLSort.Columns.Count - 1
                    row.Item(i2) = ROWSort.Item(i2)
                Next
                TBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Sub Filtering(ByRef vTbl As Data.DataTable)
        Dim tblDummy As New Data.DataTable
        Dim rowDummy, rowTbl As Data.DataRow
        Dim FilterAll As String = ""
        Dim FilterEmployee As String = ""
        Dim FilterDepartemen As String = ""
        Dim FilterPosisi As String = ""
        Dim FilterDocNo As String = ""
        Try
            tblDummy = vTbl.Copy
            FilterDocNo = GetFilterByDocNo()
            If FilterDocNo <> "" Then
                If FilterAll <> "" Then
                    FilterAll &= "AND (" & FilterDocNo & ")"
                Else
                    FilterAll &= "(" & FilterDocNo & ")"
                End If
            End If
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
            tblDummy.DefaultView.RowFilter = FilterAll
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

    Private Function GetFilterByEmployee() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.dgvEmployee.Rows
                With dr
                    If dr.Cells(Me.ColEmployee_primary.Index).Value <> -1 Then
                        If Filter = "" Then
                            Filter = "[Primary Karyawan] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary Karyawan] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
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
                    If dr.Cells(Me.ColDeptPrimary.Index).Value <> -1 Then
                        If Filter = "" Then
                            Filter = "[Primary Dept] = " & .Cells(Me.ColDeptPrimary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary Dept] = " & .Cells(Me.ColDeptPrimary.Index).Value.ToString
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
                    If dr.Cells(Me.ColJobPosisitonPrimary.Index).Value <> -1 Then
                        If Filter = "" Then
                            Filter = "[Primary Posisi] = " & .Cells(Me.ColJobPosisitonPrimary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary Posisi] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
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

    Private Function GetFilterByDocNo() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.DgvDocNo.Rows
                With dr
                    If dr.Cells(Me.ColDocNo_Primary.Index).Value <> -1 Then
                        If Filter = "" Then
                            Filter = "[Primary] = " & .Cells(Me.ColDocNo_Primary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary] = " & .Cells(Me.ColDocNo_Primary.Index).Value.ToString
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

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkFilterDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFilterDate.CheckedChanged
        Try
            If Me.ChkFilterDate.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class