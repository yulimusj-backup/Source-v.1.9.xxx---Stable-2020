Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class ClsTransaksiSalary_Override_Item

    Public Function IsiDGVItem(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_item_1TableAdapter
        Dim TBLItem As New DseTransaksiSalary_Override.tbltrans_salary_override_amount_item_1DataTable
        Dim RowItem As DseTransaksiSalary_Override.tbltrans_salary_override_amount_item_1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.AllowUserToAddRows = False

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColEmployeeCode.Index).Value = RowItem.Employee_Code
                    .Cells(i.ColEmployeeName.Index).Value = RowItem.Employee_Name
                    .Cells(i.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(RowItem.Override_Amount, vDef_Prec)
                    .Cells(i.ColOverrideReason.Index).Value = RowItem.Override_Reason
                    .Cells(i.ColPrimaryEmployee.Index).Value = RowItem.Primary_Employee
                    .Cells(i.ColDepartmentCode.Index).Value = RowItem.Dept_Code
                    .Cells(i.ColDepartmentDesc.Index).Value = RowItem.Dept_Desc
                    .Cells(i.ColJobPositionCode.Index).Value = RowItem.Job_Code
                    .Cells(i.ColJobPositionDescription.Index).Value = RowItem.Job_Desc
                End With
            Next

            'set number
            ObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTAItem As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AllowUserToAddRows = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColPrimaryEmployee.Index).Value, _
                    .Cells(i.ColOverrideReason.Index).Value, .Cells(i.ColOverrideAmount.Index).Value)
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
