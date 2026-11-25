Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Public Class ClsMasterPayrollAccess_Item

    Public Sub IsiListSalaryFormula(ByVal lbo As Windows.Forms.CheckedListBox, ByVal vTBL_SF As DseMasterPayrollAkses.tblmaster_salary_formula1DataTable)
        Dim DTA_SF As New DseMasterPayrollAksesTableAdapters.tblmaster_salary_formula1TableAdapter
        Try
            DTA_SF.Connection.ConnectionString = Core_Module.GetConstring
            vTBL_SF.Rows.Clear()
            DTA_SF.Fill(vTBL_SF)
            lbo.DataSource = Nothing
            lbo.DataSource = vTBL_SF.DefaultView

            'isikan No Event 
            lbo.DisplayMember = "CodeDesc"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_SF.Dispose()
        End Try
    End Sub

    Public Sub IsiListBoxUser(ByVal lbo As Windows.Forms.CheckedListBox, ByVal vTBL_User As DseMasterPayrollAkses.tbluserDataTable)
        Dim DTA_User As New DseMasterPayrollAksesTableAdapters.tbluserTableAdapter
        Try
            DTA_User.Connection.ConnectionString = Core_Module.GetConstring
            vTBL_User.Rows.Clear()
            DTA_User.Fill(vTBL_User)
            lbo.DataSource = Nothing
            lbo.DataSource = vTBL_User.DefaultView

            'isikan No Event 
            lbo.DisplayMember = "KodeNamaUser"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_User.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal TBL_SF As DseMasterPayrollAkses.tblmaster_salary_formula1DataTable, _
    ByVal TBL_User As DseMasterPayrollAkses.tbluserDataTable, ByVal i As FrmMasterPayrollAccess) As Boolean
        Dim Row_SF As DseMasterPayrollAkses.tblmaster_salary_formula1Row
        Dim Row_User As DseMasterPayrollAkses.tbluserRow
        Dim DTAItem_SF As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_salaryformula1TableAdapter
        Dim TBLItem_SF As New DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformula1DataTable
        'Dim RowItem_SF As DseMasterPayrollAkses.tblmaster_payroll_access_item_salaryformula1Row
        Dim DTAItem_user As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_userTableAdapter
        Dim TBLItem_user As New DseMasterPayrollAkses.tblmaster_payroll_access_item_userDataTable
        'Dim RowItem_user As DseMasterPayrollAkses.tblmaster_payroll_access_item_userRow
        Dim Primary_SF As Double = -1
        Dim Primary_User As Double = -1
        Dim Filter As String = ""
        Try
            'Connection
            DTAItem_SF.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem_user.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem_SF.FillBy_PrimaryMaster(TBLItem_SF, vPrimary)
            DTAItem_user.FillBy_PrimaryMaster(TBLItem_user, vPrimary)
            Me.IsiListSalaryFormula(i.clbSalary, TBL_SF)
            Me.IsiListBoxUser(i.clbUser, TBL_User)

            'CheckList CheckedListBox
            If TBL_SF.Rows.Count > 0 Then
                For ij As Integer = 0 To TBL_SF.Rows.Count - 1
                    Row_SF = TBL_SF.Rows(ij)
                    Filter = "[Primary Salary Formula] = " + Row_SF.primary.ToString + ""
                    TBLItem_SF.DefaultView.RowFilter = Filter
                    If TBLItem_SF.DefaultView.Count > 0 Then
                        i.clbSalary.SetItemChecked(ij, True)
                    End If
                Next
            End If
            If TBL_User.Rows.Count > 0 Then
                For ij As Integer = 0 To TBL_User.Rows.Count - 1
                    Row_User = TBL_User.Rows(ij)
                    Filter = "[Primary User] = " + Row_User.Primary.ToString + ""
                    TBLItem_user.DefaultView.RowFilter = Filter
                    If TBLItem_user.DefaultView.Count > 0 Then
                        i.clbUser.SetItemChecked(ij, True)
                    End If
                Next
            End If

            TBL_SF.DefaultView.RowFilter = ""
            TBL_User.DefaultView.RowFilter = ""

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem_SF.Dispose()
            DTAItem_user.Dispose()
            TBLItem_SF.Dispose()
            TBLItem_user.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem_SF As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_salaryformula1TableAdapter
        Dim DTAItem_user As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_userTableAdapter
        Try
            DTAItem_SF.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem_user.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem_SF.Delete(vPrimary)
            DTAItem_user.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem_SF.Dispose()
            DTAItem_user.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal i As FrmMasterPayrollAccess) As Boolean
        Dim DTAItem_SF As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_salaryformula1TableAdapter
        Dim DTAItem_user As New DseMasterPayrollAksesTableAdapters.tblmaster_payroll_access_item_userTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DV As Data.DataRowView
        Dim PrimarySalaryFormula As Double = -1
        Dim PrimaryUser As Double = -1
        Try
            DTAItem_SF.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem_user.Connection.ConnectionString = Core_Module.GetConstring

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each DV In i.clbSalary.CheckedItems
                PrimarySalaryFormula = DV.Row.Item("primary")
                DTAItem_SF.Insert(vPrimary, PrimarySalaryFormula)
            Next
            For Each DV In i.clbUser.CheckedItems
                PrimaryUser = DV.Row.Item("Primary")
                DTAItem_user.Insert(vPrimary, PrimaryUser)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem_SF.Dispose()
            DTAItem_user.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
