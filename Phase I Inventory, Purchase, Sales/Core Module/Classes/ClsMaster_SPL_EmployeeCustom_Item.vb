Option Explicit On
Option Strict Off

Public Class ClsMaster_SPL_EmployeeCustom_Item
    Public Function IsiDgv(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, ByVal i As FrmMaster_SPL_EmployeeCustom, ByVal vObjCD As ClsCekData) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_itemTableAdapter
        Dim TBLItem As New DseMaster_SPL_EmployeeCustom.tbltrans_spl_overtime_order_employee_custom_itemDataTable
        Dim rowItem As DseMaster_SPL_EmployeeCustom.tbltrans_spl_overtime_order_employee_custom_itemRow
        Dim DefaultPrecision As Integer = 0
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()

            vDGVItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColEmpCode.Index).Value = rowItem.Emp_Code
                    .Cells(i.ColEmpName.Index).Value = rowItem.Emp_Name
                    .Cells(i.ColRemarks.Index).Value = ""
                    .Cells(i.ColPrimaryEmp.Index).Value = rowItem.Primary_Employee
                End With
            Next
            vDGVItem.AllowUserToAddRows = False

            'set number
            vObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, ByVal i As FrmMaster_SPL_EmployeeCustom) As Boolean
        Dim DTAItem As New DseMaster_SPL_EmployeeCustomTableAdapters.tbltrans_spl_overtime_order_employee_custom_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColPrimaryEmp.Index).Value)
                End With
            Next
            vDGVItem.AllowUserToAddRows = True

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function
End Class
