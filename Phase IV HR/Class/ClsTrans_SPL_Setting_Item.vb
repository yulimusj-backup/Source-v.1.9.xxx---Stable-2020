Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTrans_SPL_Setting_Item
    Dim ObjCd As New Core_Module.ClsCekData

    Public Function IsiDGVItem(ByVal vDGV_Scope As Windows.Forms.DataGridView, ByVal vDGV_User As Windows.Forms.DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal Frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA_Scope As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_scopeTableAdapter
        Dim DTA_User As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter
        Dim TBL_Scope As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeDataTable
        Dim TBL_User As New DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userDataTable
        Dim row_Scope As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_scopeRow
        Dim row_User As DseTrans_SPL_Setting.tbltrans_spl_approval_group_item_userRow
        Dim rowDGV As DataGridViewRow

        Try
            DTA_Scope.Connection.ConnectionString = Core_Module.GetConstring
            DTA_User.Connection.ConnectionString = Core_Module.GetConstring

            ''isikan dgvitem scope
            vDGV_Scope.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_Scope.Rows.Clear()
            If DTA_Scope.Fill(TBL_Scope, vPrimaryTransaksi) > 0 Then
                For Each row_Scope In TBL_Scope
                    With row_Scope
                        vDGV_Scope.Rows.Add("")
                        rowDGV = vDGV_Scope.Rows(vDGV_Scope.Rows.Count - 1)
                        rowDGV.Cells(Frm.colNo.Index).Value = vDGV_Scope.Rows.Count
                        rowDGV.Cells(Frm.ColDepartmentCode.Index).Value = .Dept_Code
                        rowDGV.Cells(Frm.ColDepartmentDesc.Index).Value = .Dept_Desc
                        rowDGV.Cells(Frm.ColJobPositionCode.Index).Value = .Job_Code
                        rowDGV.Cells(Frm.ColJobPositionDesc.Index).Value = .Job_Desc
                        rowDGV.Cells(Frm.ColEmployeeCode.Index).Value = .Employee_Code
                        rowDGV.Cells(Frm.ColEmployeeName.Index).Value = .Employee_Name
                        rowDGV.Cells(Frm.ColPrimaryDept.Index).Value = .Primary_Department
                        rowDGV.Cells(Frm.ColPrimaryPosition.Index).Value = .Primary_Position
                        rowDGV.Cells(Frm.ColPrimaryEmp.Index).Value = .Primary_Employee
                        rowDGV.Cells(Frm.colRemarks.Index).Value = .Remarks
                    End With
                Next
            End If
            vDGV_Scope.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            ''isikan dgvitem User
            vDGV_User.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_User.Rows.Clear()
            If DTA_User.Fill(TBL_User, vPrimaryTransaksi) > 0 Then
                For Each row_User In TBL_User
                    With row_User
                        vDGV_User.Rows.Add("")
                        rowDGV = vDGV_User.Rows(vDGV_User.Rows.Count - 1)
                        rowDGV.Cells(Frm.colNo.Index).Value = vDGV_User.Rows.Count
                        rowDGV.Cells(Frm.ColUserCode.Index).Value = .Kode_User
                        rowDGV.Cells(Frm.ColUserName.Index).Value = .Nama_User
                        rowDGV.Cells(Frm.ColUserEmail.Index).Value = .Email_Address
                        rowDGV.Cells(Frm.ColUserRemarks.Index).Value = .Remarks
                        If .Approve_Over_Budget_Overtime = 0 Then
                            rowDGV.Cells(Frm.ColUserAllowOverBudgetSPL.Index).Value = Boolean.FalseString
                        Else
                            rowDGV.Cells(Frm.ColUserAllowOverBudgetSPL.Index).Value = Boolean.TrueString
                        End If
                        rowDGV.Cells(Frm.ColUserApproved.Index).Value = .Approved_By
                        rowDGV.Cells(Frm.ColUserApprovedPosition.Index).Value = .Approved_By_Position
                        rowDGV.Cells(Frm.ColUserPrimaryUser.Index).Value = .Primary_User
                    End With
                Next
            End If
            vDGV_User.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Scope.Dispose()
            DTA_User.Dispose()
            TBL_Scope = Nothing
            TBL_User = Nothing
        End Try
    End Function

    Public Overridable Function SimpanItemScope(ByVal vPrimaryTransaksi As Double, ByVal vDGV_Scope As DataGridView, _
    ByVal Frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA_Scope As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_scopeTableAdapter
        Dim row As DataGridViewRow

        Try
            DTA_Scope.Connection.ConnectionString = Core_Module.GetConstring

            'simpan dgvitem scope
            For i As Integer = 0 To vDGV_Scope.RowCount - 2
                row = vDGV_Scope.Rows(i)
                If String.IsNullOrEmpty(row.Cells(Frm.colRemarks.Index).Value) = True Then
                    row.Cells(Frm.colRemarks.Index).Value = ""
                End If

                DTA_Scope.Insert(vPrimaryTransaksi, row.Cells(Frm.ColPrimaryDept.Index).Value, _
                row.Cells(Frm.ColPrimaryPosition.Index).Value, row.Cells(Frm.ColPrimaryEmp.Index).Value, _
                row.Cells(Frm.colRemarks.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Scope.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanItemUser(ByVal vPrimaryTransaksi As Double, ByVal vDGV_User As DataGridView, _
    ByVal Frm As FrmTrans_SPL_Setting) As Boolean
        Dim DTA_User As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter
        Dim row As DataGridViewRow
        Dim ApakahApprove As Integer = 0

        Try
            DTA_User.Connection.ConnectionString = Core_Module.GetConstring

            'simpan dgvitem user
            For i As Integer = 0 To vDGV_User.RowCount - 2
                row = vDGV_User.Rows(i)
                If String.IsNullOrEmpty(row.Cells(Frm.ColUserRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColUserRemarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColUserAllowOverBudgetSPL.Index).Value = Boolean.TrueString Then
                    ApakahApprove = 1
                Else : ApakahApprove = 0
                End If

                DTA_User.Insert(vPrimaryTransaksi, row.Cells(Frm.ColUserPrimaryUser.Index).Value, _
                row.Cells(Frm.ColUserEmail.Index).Value, row.Cells(Frm.ColUserRemarks.Index).Value, _
                ApakahApprove, row.Cells(Frm.ColUserApproved.Index).Value, row.Cells(Frm.ColUserApprovedPosition.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_User.Dispose()
        End Try
    End Function

    Public Function HapusItemScope(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA_Scope As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_scopeTableAdapter

        Try
            'hapus data item
            DTA_Scope.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Scope.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Scope.Dispose()
        End Try
    End Function

    Public Function HapusItemUser(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA_User As New DseTrans_SPL_SettingTableAdapters.tbltrans_spl_approval_group_item_userTableAdapter

        Try
            'hapus data item
            DTA_User.Connection.ConnectionString = Core_Module.GetConstring
            DTA_User.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_User.Dispose()
        End Try
    End Function
End Class
