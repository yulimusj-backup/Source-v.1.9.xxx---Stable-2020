Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPLApprovalScope
    Public PrimaryUser As Double = -1
    Public UserApproval As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTrans_SPLApprovalScope_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvScope.ReadOnly = True
            Me.colRemarks.Visible = False
            Me.Text &= " - " & Me.UserApproval
            Me.dgvScope.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.IsiDgv(Me.PrimaryUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgv(ByVal vPrimaryUser As Double)
        Dim Dta As New DseTrans_SPLTableAdapters.tbl_approval_scopeTableAdapter
        Dim Tbl As New DseTrans_SPL.tbl_approval_scopeDataTable
        Dim row As DseTrans_SPL.tbl_approval_scopeRow
        Dim rowD As DataGridViewRow
        Dim ApakahAllEmployee As Boolean = False

        Try
            Me.dgvScope.Rows.Clear()
            Me.dgvScope.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryUser)
            For Each row In Tbl
                With row
                    If .All_Employee = 1 Then
                        ApakahAllEmployee = True
                        Exit For
                    End If

                    Me.dgvScope.Rows.Add()
                    rowD = Me.dgvScope.Rows(Me.dgvScope.RowCount - 1)
                    rowD.Cells(Me.colNo.Index).Value = Me.dgvScope.RowCount
                    rowD.Cells(Me.ColDepartmentCode.Index).Value = .Dept_Code
                    rowD.Cells(Me.ColDepartmentDesc.Index).Value = .Dept_Desc
                    rowD.Cells(Me.ColJobPositionCode.Index).Value = .Post_Code
                    rowD.Cells(Me.ColJobPositionDesc.Index).Value = .Post_Desc
                    rowD.Cells(Me.ColEmployeeCode.Index).Value = .Emp_Code
                    rowD.Cells(Me.ColEmployeeName.Index).Value = .Emp_Name
                End With
            Next

            If ApakahAllEmployee = True Then
                Me.dgvScope.Rows.Clear()

                Me.ColDepartmentCode.HeaderText = "Scope"
                Me.ColDepartmentDesc.Visible = False
                Me.ColJobPositionCode.Visible = False
                Me.ColJobPositionDesc.Visible = False
                Me.ColEmployeeCode.Visible = False
                Me.ColEmployeeName.Visible = False

                Me.dgvScope.Rows.Add()
                rowD = Me.dgvScope.Rows(Me.dgvScope.RowCount - 1)
                rowD.Cells(Me.colNo.Index).Value = Me.dgvScope.RowCount
                rowD.Cells(Me.ColDepartmentCode.Index).Value = "All Employee"
            End If

            Me.dgvScope.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class