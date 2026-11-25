Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing

Public Class FrmLaporanOvertimeOrder
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim Buffer As String = ""
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjCD As New Core_Module.ClsCekData
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim IndexContractType As Integer = 0

    Enum EnumFormat As Integer
        Recap
        Detail
        Detail_Emp
        By_OT_Order
    End Enum

    Private Sub FrmLaporanOvertimeOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboFormat.Items.Add("Recap")
            Me.cboFormat.Items.Add("Detail")
            Me.cboFormat.Items.Add("Detail Employee")
            Me.cboFormat.Items.Add("By Overtime Order")

            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboFormat.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = EnumFormat.By_OT_Order Then
                Me.Label1.Text = "Period From"
                Me.Label2.Text = "Period To"
            Else
                Me.Label1.Text = "Actual From"
                Me.Label2.Text = "Actual To"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim dta As New DseLaporanOvertimeOrderTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseLaporanOvertimeOrder.DataTable1DataTable
        Dim row As DseLaporanOvertimeOrder.DataTable1Row
        Dim dta2 As New DseLaporanOvertimeOrderTableAdapters.tbl_employee_detailTableAdapter
        Dim tbl2 As New DseLaporanOvertimeOrder.tbl_employee_detailDataTable
        Dim dtaAtt As New DseLaporanOvertimeOrderTableAdapters.tbl_ot_attendanceTableAdapter
        Dim tblAtt As New DseLaporanOvertimeOrder.tbl_ot_attendanceDataTable

        Try
            Me.crv.ReportSource = Nothing
            Me.crv.Show()

            If Me.dtpFrom.Value > Me.dtpTo.Value Then
                MsgBox("From date must be < To date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta2.Connection.ConnectionString = Core_Module.GetConstring
            dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            If Me.cboFormat.SelectedIndex = EnumFormat.By_OT_Order Then 'format spl
                'fill by spl transaksi date 
                dta.Fill(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Approved)
            ElseIf Me.cboFormat.SelectedIndex = EnumFormat.Detail_Emp Then
                'fill by date actual item approval ot
                dta2.Fill(tbl2, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Approved)
            Else 'format recap/detail
                'fill by date actual item approval ot
                dta.Fill_DateItem(tbl, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, ClsTrans_SPL.EnumStatus.Approved)
            End If

            If tbl.Count > 0 Or tbl2.Count > 0 Then
                If tbl.Count > 0 Then
                    If Me.cboFormat.SelectedIndex = EnumFormat.By_OT_Order Then 'format spl
                        Dim FilterDateFrom, FilterDateTo As Date
                        tbl.DefaultView.Sort = "[Date Item] ASC"
                        row = tbl.DefaultView(0).Row
                        FilterDateFrom = row.Date_Item
                        row = tbl.DefaultView(tbl.DefaultView.Count - 1).Row
                        FilterDateTo = row.Date_Item
                        dtaAtt.FillBy_DateFromTo(tblAtt, FilterDateFrom.Date, FilterDateTo.Date)
                    Else
                        dtaAtt.FillBy_DateFromTo(tblAtt, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                    End If
                    'filter data berdasarkan employee, job position, dan department
                    Me.Filtering(tbl)
                Else
                    If tbl2.Count > 0 Then
                        Me.Filtering2(tbl2)
                    End If
                End If

                If tbl.Count > 0 Or tbl2.Count > 0 Then
                    If Me.cboFormat.SelectedIndex = EnumFormat.By_OT_Order Then 'group by spl
                        Me.CetakBySPL(tbl, tblAtt)
                    ElseIf Me.cboFormat.SelectedIndex = EnumFormat.Detail_Emp Then 'detail employee
                        Me.CetakDetail_Employee(tbl2)
                    ElseIf Me.cboFormat.SelectedIndex = EnumFormat.Detail Then 'detail, per spl - employee
                        Me.CetakDetail(tbl, tblAtt)
                    ElseIf Me.cboFormat.SelectedIndex = EnumFormat.Recap Then 'rekap, group by employee
                        Me.CetakRecap(tbl, tblAtt)
                    End If
                End If
            End If

            If tbl.Count = 0 And tbl2.Count = 0 Then
                MsgBox("Data not found", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CetakRecap(ByVal Tbl As DseLaporanOvertimeOrder.DataTable1DataTable, _
    ByVal tblAtt As DseLaporanOvertimeOrder.tbl_ot_attendanceDataTable)
        Dim report As New CrpLaporanOvertimeOrder_Recap
        Dim ds As New DseLaporanOvertimeOrder
        Dim row As DseLaporanOvertimeOrder.DataTable1Row
        Dim tbl2 As New DseLaporanOvertimeOrder.DataTable2DataTable
        Dim row2 As DseLaporanOvertimeOrder.DataTable2Row
        Dim tblC As DseLaporanOvertimeOrder.tbl_rekapDataTable = ds.tbl_rekap
        Dim rowC As DseLaporanOvertimeOrder.tbl_rekapRow
        Dim TotalApproved As Double = 0
        Dim TotalActual As Double = 0
        Dim PrimaryAcuan As Double = -1
        Dim SubTotalApprovedStr As String = ""
        Dim SubTotalActualStr As String = ""
        Dim HitungOTCtual As Double = 0

        Try
            For Each row In Tbl
                With row
                    HitungOTCtual = Me.GetOTAttendance(.Primary_Employee, .Date_Item, tblAtt)

                    If .Status = ClsTrans_SPL.EnumStatus.Approved Then
                        tblC.DefaultView.RowFilter = "[Primary Emp] = " & .Primary_Employee
                        If tblC.DefaultView.Count > 0 Then
                            rowC = tblC.DefaultView(0).Row
                            rowC.App_OT_Dbl += .Total_OT_Approved
                            rowC.Act_OT_Dbl += HitungOTCtual
                        Else
                            rowC = tblC.Newtbl_rekapRow
                            rowC.Primary_Dept = .Primary_Dept
                            rowC.Primary_Emp = .Primary_Employee
                            rowC.Department = .Dept_Desc
                            rowC.Emp_Code = .Code
                            rowC.Emp_Name = .Full_Name
                            rowC.App_OT_Dbl = .Total_OT_Approved
                            rowC.Act_OT_Dbl = HitungOTCtual
                            tblC.Rows.Add(rowC)
                        End If

                        'untuk hitungu subtotal per department
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary_Dept
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            row2.Hitung_OT_Approved += .Total_OT_Approved
                            row2.Hitung_OT_Actual += HitungOTCtual
                        Else
                            row2 = tbl2.NewDataTable2Row
                            row2.Primary_Acuan = .Primary_Dept
                            row2.Hitung_OT_Approved = .Total_OT_Approved
                            row2.Hitung_OT_Actual = HitungOTCtual
                            tbl2.Rows.Add(row2)
                        End If

                        'hitung untuk total ot
                        TotalApproved += .Total_OT_Approved
                        TotalActual += HitungOTCtual
                    End If
                End With
            Next

            tblC.DefaultView.RowFilter = Nothing
            tblC.DefaultView.Sort = "[Primary Dept] ASC"
            For Each dv As DataRowView In tblC.DefaultView
                rowC = dv.Row
                With rowC
                    If PrimaryAcuan <> .Primary_Dept Then
                        PrimaryAcuan = .Primary_Dept
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary_Dept
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            SubTotalApprovedStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Approved)
                            SubTotalActualStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Actual)
                        Else
                            MsgBox("Subtotal not found", MsgBoxStyle.Information, "Information")
                        End If
                    End If

                    'isi format time string
                    .App_OT = Me.ObjCD.Convert_ToString(.App_OT_Dbl)
                    .Act_OT = Me.ObjCD.Convert_ToString(.Act_OT_Dbl)
                    .App_OT_Subtotal = SubTotalApprovedStr
                    .Act_OT_Subtotal = SubTotalActualStr
                End With
            Next

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Total_Approved.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalApproved))
            report.SetParameterValue(report.Parameter_Total_Actual.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalActual))
            report.SetParameterValue(report.Parameter_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
            report.SetParameterValue(report.Parameter_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail(ByVal Tbl As DseLaporanOvertimeOrder.DataTable1DataTable, _
    ByVal tblAtt As DseLaporanOvertimeOrder.tbl_ot_attendanceDataTable)
        Dim report As New CrpLaporanOvertimeOrder_Detail
        Dim ds As New DseLaporanOvertimeOrder
        Dim row As DseLaporanOvertimeOrder.DataTable1Row
        Dim tbl2 As New DseLaporanOvertimeOrder.DataTable2DataTable
        Dim row2 As DseLaporanOvertimeOrder.DataTable2Row
        Dim tblC As DseLaporanOvertimeOrder.tbl_detailDataTable = ds.tbl_detail
        Dim rowC As DseLaporanOvertimeOrder.tbl_detailRow
        Dim TotalApproved As Double = 0
        Dim TotalActual As Double = 0
        Dim PrimaryAcuan As Double = -1
        Dim SubTotalApprovedStr As String = ""
        Dim SubTotalActualStr As String = ""
        Dim HitungOTCtual As Double = 0

        Try
            For Each row In Tbl
                With row
                    HitungOTCtual = Me.GetOTAttendance(.Primary_Employee, .Date_Item, tblAtt)

                    If .Status = ClsTrans_SPL.EnumStatus.Approved Then
                        tblC.DefaultView.RowFilter = "[Primary Emp] = " & .Primary_Employee & " AND [Primary Spl] = " & .Primary
                        If tblC.DefaultView.Count > 0 Then
                            rowC = tblC.DefaultView(0).Row
                            rowC.App_OT_Dbl += .Total_OT_Approved
                            rowC.Act_OT_Dbl += HitungOTCtual
                        Else
                            rowC = tblC.Newtbl_detailRow
                            rowC.Primary_Dept = .Primary_Dept
                            rowC.Primary_Emp = .Primary_Employee
                            rowC.Primary_Spl = .Primary
                            rowC.Department = .Dept_Desc
                            rowC.Spl_No = .No_Transaksi
                            rowC.Spl_Date = ._Date
                            rowC.Emp_Code = .Code
                            rowC.Emp_Name = .Full_Name
                            rowC.App_OT_Dbl = .Total_OT_Approved
                            rowC.Act_OT_Dbl = HitungOTCtual
                            tblC.Rows.Add(rowC)
                        End If

                        'untuk hitungu subtotal per department
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary_Dept
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            row2.Hitung_OT_Approved += .Total_OT_Approved
                            row2.Hitung_OT_Actual += HitungOTCtual
                        Else
                            row2 = tbl2.NewDataTable2Row
                            row2.Primary_Acuan = .Primary_Dept
                            row2.Hitung_OT_Approved = .Total_OT_Approved
                            row2.Hitung_OT_Actual = HitungOTCtual
                            tbl2.Rows.Add(row2)
                        End If

                        'hitung untuk total ot
                        TotalApproved += .Total_OT_Approved
                        TotalActual += HitungOTCtual
                    End If
                End With
            Next

            tblC.DefaultView.RowFilter = Nothing
            tblC.DefaultView.Sort = "[Primary Dept] ASC"
            For Each dv As DataRowView In tblC.DefaultView
                rowC = dv.Row
                With rowC
                    If PrimaryAcuan <> .Primary_Dept Then
                        PrimaryAcuan = .Primary_Dept
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary_Dept
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            SubTotalApprovedStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Approved)
                            SubTotalActualStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Actual)
                        Else
                            MsgBox("Subtotal not found", MsgBoxStyle.Information, "Information")
                        End If
                    End If

                    'isi format time string
                    .App_OT = Me.ObjCD.Convert_ToString(.App_OT_Dbl)
                    .Act_OT = Me.ObjCD.Convert_ToString(.Act_OT_Dbl)
                    .App_OT_Subtotal = SubTotalApprovedStr
                    .Act_OT_Subtotal = SubTotalActualStr
                End With
            Next

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Total_Approved.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalApproved))
            report.SetParameterValue(report.Parameter_Total_Actual.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalActual))
            report.SetParameterValue(report.Parameter_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
            report.SetParameterValue(report.Parameter_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail_Employee(ByVal Tbl As DseLaporanOvertimeOrder.tbl_employee_detailDataTable)
        Dim report As New CrpLaporanOvertimeOrder_DetailEmp
        Dim ds As New DseLaporanOvertimeOrder
        Dim row, row1 As DseLaporanOvertimeOrder.tbl_employee_detailRow
        Dim tblC As DseLaporanOvertimeOrder.tbl_employee_detail_cetakDataTable = ds.tbl_employee_detail_cetak
        Dim rowC As DseLaporanOvertimeOrder.tbl_employee_detail_cetakRow
        Dim TotalApproved As Double = 0
        Dim TotalActual As Double = 0
        Dim TotalApproved_Dept As Double = 0
        Dim TotalActual_Dept As Double = 0
        Dim GrandTotalApproved As Double = 0
        Dim GrandTotalActual As Double = 0
        Dim JmlLoop1, JmlLoop2 As Integer

        Try
            Tbl.DefaultView.RowFilter = ""
            Tbl.DefaultView.Sort = "[Dept Code] ASC, [Emp Code] ASC, [Date] ASC"
            For j As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(j).Row
                With row
                    JmlLoop1 = 2
                    JmlLoop2 = 2

                    If .OT_Early_Approved > 0 And .OT_Aprroved > 0 Then
                        JmlLoop1 = 1
                        JmlLoop2 = 2
                    Else
                        If .OT_Early_Approved > 0 Then
                            JmlLoop1 = 1
                            JmlLoop2 = 1
                        ElseIf .OT_Aprroved > 0 Then
                            JmlLoop1 = 2
                            JmlLoop2 = 2
                        End If
                    End If

                    If j <> 0 Then
                        row1 = Tbl.DefaultView(j - 1).Row
                        If .Primay_Dept <> row1.Primay_Dept Then
                            TotalApproved_Dept = 0
                            TotalActual_Dept = 0
                        End If
                    End If

                    For i As Integer = JmlLoop1 To JmlLoop2
                        rowC = tblC.Newtbl_employee_detail_cetakRow
                        rowC.Urutan = JmlLoop1
                        rowC.No_Transaksi = .Overime_Order_No
                        rowC._Date = ._Date
                        rowC.Emp_Code = .Emp_Code
                        rowC.Emp_NIK = .Emp_NIK
                        rowC.Emp_Name = .Emp_Name
                        rowC.Dept_Code = .Dept_Code
                        rowC.Dept_Desc = .Dept_Desc

                        rowC.Day_of_Week = Me.ObjCD.GetHari_English(._Date.DayOfWeek)
                        rowC.Date_Item = row._Date

                        If i = 1 Then
                            rowC.Start = Me.ObjCD.Convert_ToString(.Start_Time_Early)
                            If .Break_Out_Time_Early > 0 Then
                                rowC.Break_Out = Me.ObjCD.Convert_ToString(.Break_Out_Time_Early)
                            End If
                            If .Break_In_Time_Early > 0 Then
                                rowC.Break_In = Me.ObjCD.Convert_ToString(.Break_In_Time_Early)
                            End If
                            rowC._End = Me.ObjCD.Convert_ToString(.End_Time_Early)
                            rowC.Approved_Duration = Me.ObjCD.Convert_ToString(.OT_Early_Approved)
                            rowC.Actual_Duration = Me.ObjCD.Convert_ToString(.OT_Early_Attendance)

                            TotalApproved += .OT_Early_Approved
                            TotalActual += .OT_Early_Attendance
                            TotalApproved_Dept += .OT_Early_Approved
                            TotalActual_Dept += .OT_Early_Attendance
                        ElseIf i = 2 Then
                            rowC.Start = Me.ObjCD.Convert_ToString(.Start_Time)
                            If .Break_Out_Time > 0 Then
                                rowC.Break_Out = Me.ObjCD.Convert_ToString(.Break_Out_Time)
                            End If
                            If .Break_In_Time > 0 Then
                                rowC.Break_In = Me.ObjCD.Convert_ToString(.Break_In_Time)
                            End If
                            rowC._End = Me.ObjCD.Convert_ToString(.End_Time)
                            rowC.Approved_Duration = Me.ObjCD.Convert_ToString(.OT_Aprroved)
                            rowC.Actual_Duration = Me.ObjCD.Convert_ToString(.OT_Attendance)

                            TotalApproved += .OT_Aprroved
                            TotalActual += .OT_Attendance
                            TotalApproved_Dept += .OT_Aprroved
                            TotalActual_Dept += .OT_Attendance
                        End If

                        rowC.Total_Approved_Duration = Me.ObjCD.Convert_ToString(TotalApproved)
                        rowC.Total_Actual_Duration = Me.ObjCD.Convert_ToString(TotalActual)
                        rowC.Total_App_Duration_Dept = Me.ObjCD.Convert_ToString(TotalApproved_Dept)
                        rowC.Total_Act_Duration_Dept = Me.ObjCD.Convert_ToString(TotalActual_Dept)

                        tblC.Rows.Add(rowC)
                    Next
                End With

                If j <> Tbl.DefaultView.Count - 1 Then
                    row1 = Tbl.DefaultView(j + 1).Row
                    If row.Primary_Employee <> row1.Primary_Employee Then
                        GrandTotalApproved += TotalApproved
                        GrandTotalActual += TotalActual
                        TotalApproved = 0
                        TotalActual = 0
                    End If
                Else
                    GrandTotalApproved += TotalApproved
                    GrandTotalActual += TotalActual
                End If
            Next

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_GrandTotal_Approved.ParameterFieldName, Me.ObjCD.Convert_ToString(GrandTotalApproved))
            report.SetParameterValue(report.Parameter_GrandTotal_Actual.ParameterFieldName, Me.ObjCD.Convert_ToString(GrandTotalActual))
            report.SetParameterValue(report.Parameter_PeriodFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
            report.SetParameterValue(report.Parameter_PeriodTo.ParameterFieldName, Me.dtpTo.Value.Date)
            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakBySPL(ByVal Tbl As DseLaporanOvertimeOrder.DataTable1DataTable, _
    ByVal tblAtt As DseLaporanOvertimeOrder.tbl_ot_attendanceDataTable)
        Dim report As New CrpLaporanOvertimeOrder_BySPL
        Dim ds As New DseLaporanOvertimeOrder
        Dim row As DseLaporanOvertimeOrder.DataTable1Row
        Dim tbl2 As New DseLaporanOvertimeOrder.DataTable2DataTable
        Dim row2 As DseLaporanOvertimeOrder.DataTable2Row
        Dim tblC As DseLaporanOvertimeOrder.tbl_splDataTable = ds.tbl_spl
        Dim rowC As DseLaporanOvertimeOrder.tbl_splRow
        Dim TotalApproved As Double = 0
        Dim TotalActual As Double = 0
        Dim PrimaryAcuan As Double = -1
        Dim SubTotalApprovedStr As String = ""
        Dim SubTotalActualStr As String = ""
        Dim HitungOTCtual As Double = 0

        Try
            For Each row In Tbl
                With row
                    HitungOTCtual = Me.GetOTAttendance(.Primary_Employee, .Date_Item, tblAtt)

                    If .Status = ClsTrans_SPL.EnumStatus.Approved Then
                        tblC.DefaultView.RowFilter = "[Primary Emp] = " & .Primary_Employee & " AND [Primary Spl] = " & .Primary
                        If tblC.DefaultView.Count > 0 Then
                            rowC = tblC.DefaultView(0).Row
                            rowC.App_OT_Dbl += .Total_OT_Approved
                            rowC.Act_OT_Dbl += HitungOTCtual
                        Else
                            rowC = tblC.Newtbl_splRow
                            rowC.Primary_Dept = .Primary_Dept
                            rowC.Primary_Emp = .Primary_Employee
                            rowC.Primary_Spl = .Primary
                            rowC.Department = .Dept_Desc
                            rowC.Spl_No = .No_Transaksi
                            rowC.From_ = .From
                            rowC._To = ._To
                            rowC.Spl_Date = ._Date
                            rowC.Emp_Code = .Code
                            rowC.Emp_Name = .Full_Name
                            rowC.App_OT_Dbl = .Total_OT_Approved
                            rowC.Act_OT_Dbl = HitungOTCtual
                            tblC.Rows.Add(rowC)
                        End If

                        'untuk hitungu subtotal per spl
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            row2.Hitung_OT_Approved += .Total_OT_Approved
                            row2.Hitung_OT_Actual += HitungOTCtual
                        Else
                            row2 = tbl2.NewDataTable2Row
                            row2.Primary_Acuan = .Primary
                            row2.Hitung_OT_Approved = .Total_OT_Approved
                            row2.Hitung_OT_Actual = HitungOTCtual
                            tbl2.Rows.Add(row2)
                        End If

                        'hitung untuk total ot
                        TotalApproved += .Total_OT_Approved
                        TotalActual += HitungOTCtual
                    End If
                End With
            Next

            tblC.DefaultView.RowFilter = Nothing
            tblC.DefaultView.Sort = "[Primary Spl] ASC"
            For Each dv As DataRowView In tblC.DefaultView
                rowC = dv.Row
                With rowC
                    If PrimaryAcuan <> .Primary_Spl Then
                        PrimaryAcuan = .Primary_Spl
                        tbl2.DefaultView.RowFilter = "[Primary Acuan] = " & .Primary_Spl
                        If tbl2.DefaultView.Count > 0 Then
                            row2 = tbl2.DefaultView(0).Row
                            SubTotalApprovedStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Approved)
                            SubTotalActualStr = Me.ObjCD.Convert_ToString(row2.Hitung_OT_Actual)
                        Else
                            MsgBox("Subtotal not found", MsgBoxStyle.Information, "Information")
                        End If
                    End If

                    'isi format time string
                    .App_OT = Me.ObjCD.Convert_ToString(.App_OT_Dbl)
                    .Act_OT = Me.ObjCD.Convert_ToString(.Act_OT_Dbl)
                    .App_OT_Subtotal = SubTotalApprovedStr
                    .Act_OT_Subtotal = SubTotalActualStr
                End With
            Next

            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_Total_Approved.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalApproved))
            report.SetParameterValue(report.Parameter_Total_Actual.ParameterFieldName, Me.ObjCD.Convert_ToString(TotalActual))
            report.SetParameterValue(report.Parameter_DateFrom.ParameterFieldName, Me.dtpFrom.Value.Date)
            report.SetParameterValue(report.Parameter_DateTo.ParameterFieldName, Me.dtpTo.Value.Date)
            Me.crv.ReportSource = report
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetOTAttendance(ByVal vPrimaryEmp As Double, ByVal vDate As Date, _
    ByVal tblAtt As DseLaporanOvertimeOrder.tbl_ot_attendanceDataTable) As Double
        Dim rowAtt As DseLaporanOvertimeOrder.tbl_ot_attendanceRow
        Dim Retval As Double = 0

        Try
            tblAtt.DefaultView.RowFilter = "[primary karyawan] = " & vPrimaryEmp.ToString & " AND [date] = #" & vDate.Date & "#"
            If tblAtt.DefaultView.Count > 0 Then
                rowAtt = tblAtt.DefaultView(0).Row
                Retval = rowAtt.Total_OT_Attendance
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Private Sub Filtering2(ByRef vTbl As DseLaporanOvertimeOrder.tbl_employee_detailDataTable)
        Dim tblDummy As New DseLaporanOvertimeOrder.tbl_employee_detailDataTable
        Dim FilterAll As String = ""
        Dim FilterEmployee As String = ""
        Dim FilterDepartemen As String = ""
        Dim FilterPosisi As String = ""

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

            tblDummy.DefaultView.RowFilter = FilterAll
            vTbl.Rows.Clear()
            vTbl.Merge(tblDummy.DefaultView.ToTable)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Filtering(ByRef vTbl As DseLaporanOvertimeOrder.DataTable1DataTable)
        Dim tblDummy As New DseLaporanOvertimeOrder.DataTable1DataTable
        Dim FilterAll As String = ""
        Dim FilterEmployee As String = ""
        Dim FilterDepartemen As String = ""
        Dim FilterPosisi As String = ""

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

            tblDummy.DefaultView.RowFilter = FilterAll
            vTbl.Rows.Clear()
            vTbl.Merge(tblDummy.DefaultView.ToTable)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Function GetFilterByEmployee() As String
        Dim Filter As String = ""
        Try
            For Each dr As DataGridViewRow In Me.dgvEmployee.Rows
                With dr
                    If dr.Cells(Me.ColEmployee_primary.Index).Value <> -1 Then
                        If Filter = "" Then
                            Filter = "[Primary Employee] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
                        Else
                            Filter += " OR [Primary Employee] = " & .Cells(Me.ColEmployee_primary.Index).Value.ToString
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
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
End Class