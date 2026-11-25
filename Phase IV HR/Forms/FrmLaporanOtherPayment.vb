Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class FrmLaporanOtherPayment

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Tipe As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.SalaryOverride
    Dim Buffer As String = ""

    Dim TBLDOCNo As New DseLaporanOtherPayment.tbl_DocumentDataTable
    Dim rowDocNo As DseLaporanOtherPayment.tbl_DocumentRow

    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjComSalary As New Core_Module.ClsMasterSalaryFormula
    Dim ObjFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
    Dim ObjDocNo As New ClsTransaksiSalary_Override
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim IndexComSalary As Integer = 0
    Dim IndexDocNo As Integer = 0



    Private Sub FrmLaporanOtherPayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSalaryFormula.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvDocNo.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboReport.SelectedIndex = 0
            Me.CboGrouping.SelectedIndex = 0

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
                                    Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
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
                                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
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
                            Me.IndexEmp = e.RowIndex
                            Me.BgwEmployee.RunWorkerAsync()
                        End If
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
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
    Private Function GetIndexItemComponentSalary(ByVal vIdentifierComSalary As String, _
          ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalaryFormula.Rows.Count - 1
                RowDGV = Me.DgvSalaryFormula.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSC_IdentifierComponent.Index).Value = vIdentifierComSalary Then
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

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvEmployee.Rows(Me.dgvEmployee.Rows.Count - 1)
            dr.Cells(Me.ColEmployee_primary.Index).Value = -1
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
                            Me.IndexDept = e.RowIndex
                            Me.BgwDepartment.RunWorkerAsync()
                        End If
                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
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
                            Me.IndexJobPost = e.RowIndex
                            Me.BgwJobPosition.RunWorkerAsync()
                        End If
                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
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

    Private Sub DgvSalaryFormula_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalaryFormula.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvSalaryFormula.Item(Me.ColSC_ComponentName.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColSC_ComponentName.Index Then
                    Me.Buffer = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColSC_ComponentName.Index Then
                    Me.Buffer = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalaryFormula_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalaryFormula.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalaryFormula.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim IdenComSalary As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                Me.DgvSalaryFormula.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvSalaryFormula.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If IsNothing(Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColSC_ComponentName.Index Then
                        Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        IdenComSalary = Me.ObjFormulaItem.GetIdentifierFormula(Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value)
                        If IdenComSalary <> "-1-1" Then
                            index = Me.GetIndexItemComponentSalary(IdenComSalary, e.RowIndex)
                            If index < 0 Then
                                'Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value = me.obj
                                Me.DgvSalaryFormula.Item(Me.ColSC_ComponentName.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetDescComponent(IdenComSalary)
                                Me.DgvSalaryFormula.Item(Me.ColSC_IdentifierComponent.Index, e.RowIndex).Value = IdenComSalary
                                Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetPrimaryFormula_Identifier(IdenComSalary)
                                Me.DgvSalaryFormula.Item(Me.ColSC_SalaryFormula.Index, e.RowIndex).Value = Me.ObjComSalary.GetDescription(Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalaryFormula.Item(Me.ColSC_ComponentName.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexComSalary = e.RowIndex
                                        BgwSalaryComponent.RunWorkerAsync()
                                    Else
                                        Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New FrmTrans_Gaji_Override_Cari_Formula
                            ij.txtCari.Text = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mIdentifierFormula <> "-1-1" Then
                                index = Me.GetIndexItemComponentSalary(ij.mIdentifierFormula, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvSalaryFormula.Item(Me.ColSC_ComponentName.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetDescComponent(ij.mIdentifierFormula)
                                    Me.DgvSalaryFormula.Item(Me.ColSC_IdentifierComponent.Index, e.RowIndex).Value = ij.mIdentifierFormula
                                    Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetPrimaryFormula_Identifier(ij.mIdentifierFormula)
                                    Me.DgvSalaryFormula.Item(Me.ColSC_SalaryFormula.Index, e.RowIndex).Value = Me.ObjComSalary.GetDescription(Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value)
                                Else
                                    If Me.Buffer = Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")

                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexComSalary = e.RowIndex
                                            BgwSalaryComponent.RunWorkerAsync()
                                        Else
                                            Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.DgvSalaryFormula.Item(Me.ColSC_ComponentName.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetDescComponent(ij.mIdentifierFormula)
                                        Me.DgvSalaryFormula.Item(Me.ColSC_IdentifierComponent.Index, e.RowIndex).Value = ij.mIdentifierFormula
                                        Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value = Me.ObjFormulaItem.GetPrimaryFormula_Identifier(ij.mIdentifierFormula)
                                        Me.DgvSalaryFormula.Item(Me.ColSC_SalaryFormula.Index, e.RowIndex).Value = Me.ObjComSalary.GetDescription(Me.DgvSalaryFormula.Item(Me.ColSC_PrimarySalary.Index, e.RowIndex).Value)
                                    End If
                                End If
                            ElseIf ij.mIdentifierFormula = "-1-1" Then
                                If Me.DgvSalaryFormula.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexComSalary = e.RowIndex
                                        Me.BgwSalaryComponent.RunWorkerAsync()
                                    Else
                                        Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                                If Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexComSalary = e.RowIndex
                                        Me.BgwSalaryComponent.RunWorkerAsync()
                                    Else
                                        Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                            'ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                    End If
                Else
                    If e.ColumnIndex = Me.ColSC_ComponentName.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.IndexComSalary = e.RowIndex
                            Me.BgwSalaryComponent.RunWorkerAsync()
                        End If
                        Me.DgvSalaryFormula.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvSalaryFormula.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvSalaryFormula.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
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

    Private Sub BgwSalaryComponent_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSalaryComponent.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSalaryComponent_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSalaryComponent.RunWorkerCompleted
        Try
            If DgvSalaryFormula.RowCount > 0 Then
                Me.DgvSalaryFormula.Rows.RemoveAt(Me.IndexComSalary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDocNo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDocNo.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvDocNo.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDoc As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                Me.DgvDocNo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvDocNo.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If IsNothing(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjDocNo.ApakahSudahAda(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryDoc = Me.ObjDocNo.GetPrimary(Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemDocNo(PrimaryDoc, e.RowIndex)
                            If index < 0 Then
                                Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = PrimaryDoc
                                Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjDocNo.GetDocumentNo(PrimaryDoc)

                                Me.TBLDOCNo.DefaultView.RowFilter = Nothing
                                Me.TBLDOCNo.DefaultView.RowFilter = "[Primary] = " & PrimaryDoc
                                If Me.TBLDOCNo.DefaultView.Count > 0 Then
                                    Me.rowDocNo = Me.TBLDOCNo.DefaultView(0).Row
                                    Me.rowDocNo.No_Doc = Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value
                                Else
                                    Me.rowDocNo = Me.TBLDOCNo.Newtbl_DocumentRow
                                    With Me.rowDocNo
                                        .Primary = PrimaryDoc
                                        .No_Doc = Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value
                                        If Me.CboReport.SelectedIndex = 0 Then
                                            .Type = 0
                                        Else
                                            .Type = 1
                                        End If
                                    End With
                                    Me.TBLDOCNo.Rows.Add(Me.rowDocNo)
                                End If

                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDocNo = e.RowIndex
                                        BgwDocNo.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New FrmTrans_Gaji_Override_Cari2
                            ij.txtCari.Text = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            If Me.CboReport.SelectedIndex = 0 Then
                                ij.vTipe = ClsTransaksiSalary_Override.EnumType.OtherPayment
                            Else
                                ij.vTipe = ClsTransaksiSalary_Override.EnumType.SalaryOverride
                            End If
                            ij.ShowDialog()


                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemDocNo(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjDocNo.GetDocumentNo(ij.mPrimary)
                                    Me.TBLDOCNo.DefaultView.RowFilter = Nothing
                                    Me.TBLDOCNo.DefaultView.RowFilter = "[Primary] = " & ij.mPrimary
                                    If Me.TBLDOCNo.DefaultView.Count > 0 Then
                                        Me.rowDocNo = Me.TBLDOCNo.DefaultView(0).Row
                                        Me.rowDocNo.No_Doc = Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value
                                    Else
                                        Me.rowDocNo = Me.TBLDOCNo.Newtbl_DocumentRow
                                        With Me.rowDocNo
                                            .Primary = ij.mPrimary
                                            .No_Doc = Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value
                                            If Me.CboReport.SelectedIndex = 0 Then
                                                .Type = 0
                                            Else
                                                .Type = 1
                                            End If
                                        End With
                                        Me.TBLDOCNo.Rows.Add(Me.rowDocNo)
                                    End If
                                Else
                                    If Me.Buffer = Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")

                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexDocNo = e.RowIndex
                                            BgwDocNo.RunWorkerAsync()
                                        Else
                                            Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value = Me.ObjDocNo.GetDocumentNo(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.DgvDocNo.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDocNo = e.RowIndex
                                        Me.BgwDocNo.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                                If Me.DgvDocNo.Item(Me.ColDocNo_Document.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDocNo = e.RowIndex
                                        Me.BgwDocNo.RunWorkerAsync()
                                    Else
                                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                    If e.ColumnIndex = Me.ColDocNo_Document.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.IndexDocNo = e.RowIndex
                            Me.BgwDocNo.RunWorkerAsync()
                        End If
                        Me.DgvDocNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvDocNo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvDocNo.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub BgwDocNo_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDocNo.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDocNo_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDocNo.RunWorkerCompleted
        Try
            If DgvDocNo.RowCount > 0 Then
                Me.DgvDocNo.Rows.RemoveAt(Me.IndexDocNo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboReport_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboReport.SelectedIndexChanged
        Try
            Dim rowDGV As DataGridViewRow
            Dim tipeDoc As String

            Me.DgvDocNo.Rows.Clear()
            tipeDoc = Me.CboReport.SelectedIndex
            Me.TBLDOCNo.DefaultView.RowFilter = Nothing
            Me.TBLDOCNo.DefaultView.RowFilter = "[Type] = '" & tipeDoc & "'"
            Me.DgvDocNo.AllowUserToAddRows = False
            If Me.TBLDOCNo.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.TBLDOCNo.DefaultView.Count - 1
                    rowDocNo = TBLDOCNo.DefaultView(i).Row
                    Me.DgvDocNo.Rows.Add()

                    rowDGV = Me.DgvDocNo.Rows(i)
                    With rowDocNo
                        rowDGV.Cells(Me.ColDocNo_Primary.Index).Value = .Primary
                        rowDGV.Cells(Me.ColDocNo_Document.Index).Value = .No_Doc
                    End With
                Next
            End If
            Me.DgvDocNo.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub DgvDocNo_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvDocNo.UserDeletingRow
        Try
            Me.TBLDOCNo.DefaultView.RowFilter = Nothing
            Me.TBLDOCNo.DefaultView.RowFilter = "[Primary] = " & Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, Me.DgvDocNo.CurrentRow.Index).Value
            If Me.TBLDOCNo.DefaultView.Count > 0 Then
                Me.rowDocNo = Me.TBLDOCNo.DefaultView(0).Row
                Me.TBLDOCNo.Rows.Remove(rowDocNo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            If ApakahDataBenar() = False Then
                Exit Sub
            Else
                If BuatLaporan(Me.crvLaporanKontrakKaryawan) = False Then
                    MsgBox("No Data.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function BuatLaporan(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim DS As New DseLaporanOtherPayment
        Dim TBLCetak As DseLaporanOtherPayment.TblCetakDataTable = DS.TblCetak
        Dim RowCetak As DseLaporanOtherPayment.TblCetakRow

        Dim DTADummy As New DseLaporanOtherPaymentTableAdapters.TblLaporanOtherSalaryOverrideTableAdapter
        Dim TBLDummy As New DseLaporanOtherPayment.TblLaporanOtherSalaryOverrideDataTable
        Dim RowDummy As DseLaporanOtherPayment.TblLaporanOtherSalaryOverrideRow
        Dim tipe As Integer = 0
        Dim c As Integer = 0

        Dim listPrimaryEmp As New List(Of String)
        Dim listPrimaryDept As New List(Of String)
        Dim listPrimaryJob As New List(Of String)
        Dim listPrimaryDoc As New List(Of String)
        Dim listIdentFormula As New List(Of String)


        Dim judul As String = ""
        Dim Report As New CrpLaporanOtherPaymentNoGroup
        Dim Report1 As New CrpLaporanOtherPaymentGroup
        Try
            Me.Cursor = Cursors.WaitCursor
            DTADummy.Connection.ConnectionString = Core_Module.GetConstring




            If Me.CboReport.SelectedIndex = 0 Then
                tipe = 1
            Else
                tipe = 0
            End If

            If DTADummy.FillBy_TglDariSampai(TBLDummy, tipe, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date) > 0 Then

                'filter Document No
                If Me.DgvDocNo.Rows.Count - 1 <> 0 Then
                    Me.DgvDocNo.AllowUserToAddRows = False
                    For i As Integer = 0 To Me.DgvDocNo.Rows.Count - 1
                        If listPrimaryDoc.Contains(Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, i).Value) = False Then
                            listPrimaryDoc.Add(Me.DgvDocNo.Item(Me.ColDocNo_Primary.Index, i).Value)
                        End If
                    Next
                    Me.DgvDocNo.AllowUserToAddRows = True
                    c = 0
                    While c < TBLDummy.Rows.Count
                        RowDummy = TBLDummy.Rows(c)
                        If listPrimaryDoc.Contains(RowDummy.Primary) = False Then
                            TBLDummy.Rows.Remove(RowDummy)
                        Else
                            c += 1
                        End If
                    End While
                End If

                'filter Employee
                If Me.dgvEmployee.Rows.Count - 1 <> 0 Then
                    Me.dgvEmployee.AllowUserToAddRows = False
                    For i As Integer = 0 To Me.dgvEmployee.Rows.Count - 1
                        If listPrimaryEmp.Contains(Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, i).Value) = False Then
                            listPrimaryEmp.Add(Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, i).Value)
                        End If
                    Next
                    Me.dgvEmployee.AllowUserToAddRows = True
                    c = 0
                    While c < TBLDummy.Rows.Count
                        RowDummy = TBLDummy.Rows(c)
                        If listPrimaryEmp.Contains(RowDummy.Primary_Employee) = False Then
                            TBLDummy.Rows.Remove(RowDummy)
                        Else
                            c += 1
                        End If
                    End While
                End If

                'filter Departemen
                If Me.dgvDepartment.Rows.Count - 1 <> 0 Then
                    Me.dgvDepartment.AllowUserToAddRows = False
                    For i As Integer = 0 To Me.dgvDepartment.Rows.Count - 1
                        If listPrimaryDept.Contains(Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, i).Value) = False Then
                            listPrimaryDept.Add(Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, i).Value)
                        End If
                    Next
                    Me.dgvDepartment.AllowUserToAddRows = True
                    c = 0
                    While c < TBLDummy.Rows.Count
                        RowDummy = TBLDummy.Rows(c)
                        If listPrimaryDept.Contains(RowDummy.primary_departemen) = False Then
                            TBLDummy.Rows.Remove(RowDummy)
                        Else
                            c += 1
                        End If
                    End While
                End If

                'filter Job Position
                If Me.dgvJobPosition.Rows.Count - 1 <> 0 Then
                    Me.dgvJobPosition.AllowUserToAddRows = False
                    For i As Integer = 0 To Me.dgvJobPosition.Rows.Count - 1
                        If listPrimaryJob.Contains(Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, i).Value) = False Then
                            listPrimaryJob.Add(Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, i).Value)
                        End If
                    Next
                    Me.dgvJobPosition.AllowUserToAddRows = True
                    c = 0
                    While c < TBLDummy.Rows.Count
                        RowDummy = TBLDummy.Rows(c)
                        If listPrimaryJob.Contains(RowDummy.primary_position) = False Then
                            TBLDummy.Rows.Remove(RowDummy)
                        Else
                            c += 1
                        End If
                    End While
                End If

                'filter Salary Component
                If Me.DgvSalaryFormula.Rows.Count - 1 <> 0 Then
                    Me.DgvSalaryFormula.AllowUserToAddRows = False
                    For i As Integer = 0 To Me.DgvSalaryFormula.Rows.Count - 1
                        If listIdentFormula.Contains(Me.DgvSalaryFormula.Item(Me.ColSC_IdentifierComponent.Index, i).Value) = False Then
                            listIdentFormula.Add(Me.DgvSalaryFormula.Item(Me.ColSC_IdentifierComponent.Index, i).Value)
                        End If
                    Next
                    Me.DgvSalaryFormula.AllowUserToAddRows = True
                    c = 0
                    While c < TBLDummy.Rows.Count
                        RowDummy = TBLDummy.Rows(c)
                        If listIdentFormula.Contains(RowDummy.Identifier_formula) = False Then
                            TBLDummy.Rows.Remove(RowDummy)
                        Else
                            c += 1
                        End If
                    End While
                End If



                'Salin ke Table Cetak
                For i As Integer = 0 To TBLDummy.Rows.Count - 1
                    RowDummy = TBLDummy.Rows(i)
                    RowCetak = TBLCetak.NewTblCetakRow
                    With RowDummy
                        RowCetak.Primary = .Primary
                        RowCetak.Document_No = .Document_No
                        RowCetak.Code = .Code
                        RowCetak.Full_Name = .Full_Name
                        RowCetak.NIK = .NIK
                        RowCetak.Description = .Description
                        RowCetak.tipe = .tipe
                        RowCetak._Date = ._Date
                        RowCetak.Override_Amount = .Override_Amount
                        If Me.CboGrouping.SelectedIndex = 1 Then
                            RowCetak.Group_ = .Code
                        ElseIf Me.CboGrouping.SelectedIndex = 2 Then
                            RowCetak.Group_ = .Description
                        Else
                            RowCetak.Group_ = .Document_No
                        End If
                    End With
                    TBLCetak.Rows.Add(RowCetak)
                Next

                If Me.CboReport.SelectedIndex = 0 Then
                    judul = "OTHER PAYMENT REPORT"
                Else
                    judul = "SALARY OVERRIDE REPORT"
                End If
                If Me.CboGrouping.Text = "No Grouping" Then
                    Report.SetDataSource(DS)
                    Report.SetParameterValue(Report.Parameter_judul.ParameterFieldName, judul)
                    Report.SetParameterValue(Report.Parameter_Tanggal_Dari.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Report.SetParameterValue(Report.Parameter_Tanggal_Sampai.ParameterFieldName, Me.dtpTo.Value.Date)
                    Report.SetParameterValue(Report.Parameter_DeffPrec.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency())
                    vCRV.ReportSource = Report
                    vCRV.Show()
                Else
                    Report1.SetDataSource(DS)
                    Report1.SetParameterValue(Report1.Parameter_judul.ParameterFieldName, judul)
                    Report1.SetParameterValue(Report1.Parameter_GroupBy.ParameterFieldName, Me.CboGrouping.SelectedIndex)
                    Report1.SetParameterValue(Report1.Parameter_Tanggal_Dari.ParameterFieldName, Me.dtpFrom.Value.Date)
                    Report1.SetParameterValue(Report1.Parameter_Tanggal_Sampai.ParameterFieldName, Me.dtpTo.Value.Date)
                    Report1.SetParameterValue(Report1.Parameter_DeffPrec.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency())
                    vCRV.ReportSource = Report1
                    vCRV.Show()
                End If
            Else
                Me.Cursor = Cursors.Default
                Return False
            End If
            Me.Cursor = Cursors.Default
            Return True

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Value From date must be < To Date", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalaryFormula_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSalaryFormula.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvSalaryFormula.Rows(Me.DgvSalaryFormula.Rows.Count - 1)
            dr.Cells(Me.ColSC_PrimarySalary.Index).Value = -1
            dr.Cells(Me.ColSC_ComponentName.Index).Value = ""
            dr.Cells(Me.ColSC_SalaryFormula.Index).Value = ""
            dr.Cells(Me.ColSC_IdentifierComponent.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
