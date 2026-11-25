Option Explicit On
Option Strict Off

Imports Core_module, System.Windows.Forms
Public Class ClsTransaksiSalary_Quick

    Public Sub IsiDGVQuick(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vObjCD As ClsCekData, _
    ByVal vFrm As FrmTrans_Gaji_Quick)
        Dim DTA As New DseTransaksiSalary_QuickTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSalary_Quick.DataTable1DataTable
        Dim Row As DseTransaksiSalary_Quick.DataTable1Row
        Dim RowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'autosize dgv
            vFrm.DGVSalaryCalculation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vFrm.DGVSalaryCalculation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If vFrm.CboTypeDateFilter.SelectedIndex = 0 Then
                'transaction date
                DTA.FillBy_TanggalDariSampai_Transaksi(TBL, vEndDate, vStartDate)
            Else
                'period
                DTA.FillBy_Period(TBL, vEndDate, vStartDate)
            End If

            'add aul
            Dim dv As DataRowView
            Dim ObjSort As New ClsTransaksiQuick_Sorting
            Dim ObjUser As New Core_Module.ClsUser
            Dim PrimaryUser As Double = ObjUser.GetPrimary(vFrm.UserCode)
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = ObjSort.GetSetting(PrimaryUser, ClsTransaksiQuick_Sorting.EnumTipeTransaksi.QuickSalary)

            vFrm.DGVSalaryCalculation.Rows.Clear()
            For Each dv In TBL.DefaultView
                Row = dv.Row
                vFrm.DGVSalaryCalculation.Rows.Add()
                RowDGV = vFrm.DGVSalaryCalculation.Rows(vFrm.DGVSalaryCalculation.Rows.Count - 1)
                RowDGV.Cells(vFrm.ColNo.Index).Value = ""
                My.Application.ChangeCulture("id-ID")
                My.Application.ChangeUICulture("id-ID")
                RowDGV.Cells(vFrm.ColStartDate.Index).Value = Row.Start_Date
                RowDGV.Cells(vFrm.ColEndDate.Index).Value = Row.End_Date
                RowDGV.Cells(vFrm.COlDate.Index).Value = Row.Tanggal_Transaksi
                My.Application.ChangeCulture("en-US")
                My.Application.ChangeUICulture("en-US")
                RowDGV.Cells(vFrm.ColEmployeeCode.Index).Value = Row.Code
                RowDGV.Cells(vFrm.ColEmployeeName.Index).Value = Row.Full_Name
                RowDGV.Cells(vFrm.ColDept.Index).Value = Row.Dept_Desc
                RowDGV.Cells(vFrm.ColJobPosition.Index).Value = Row.Position_Description
                RowDGV.Cells(vFrm.ColPrimaryTransaksi.Index).Value = Row.Primary_Salary
                RowDGV.Cells(vFrm.ColDocumentNo.Index).Value = Row.No_Transaksi

                RowDGV.Cells(vFrm.ColEmployeeNIK.Index).Value = Row.NIK
                RowDGV.Cells(vFrm.ColDeptCode.Index).Value = Row.Dept_Code
                RowDGV.Cells(vFrm.ColJobPositionCode.Index).Value = Row.Position_Code
            Next

            vFrm.COlDate.DefaultCellStyle.Format = "dd MMM yyy"
            vFrm.ColStartDate.DefaultCellStyle.Format = "dd MMM yyy"
            vFrm.ColEndDate.DefaultCellStyle.Format = "dd MMM yyy"

            'set number
            vObjCD.SetNumber(vFrm.DGVSalaryCalculation)

            'autosize dgv
            vFrm.DGVSalaryCalculation.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vFrm.DGVSalaryCalculation.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            vFrm.LblAllPaid.Text = "Paid: Check All"
            vFrm.LblAllPrint.Text = "Print: Check All"
            vFrm.LblAllDelete.Text = "Delete: Check All"
            vFrm.LblAllEmail.Text = "Email: Check All"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
