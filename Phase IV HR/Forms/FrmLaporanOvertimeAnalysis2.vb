Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmLaporanOvertimeAnalysis2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim Buffer = ""
    Dim DateFrom As Date = Now
    Dim DateTo As Date = Now
    Dim FormatDecimal As Integer = 0

    Public Enum EnumTipeData As Integer
        Department
        Job_Position
        Employee
        Monthly_Period
    End Enum

    Public Enum EnumTipeGrouping As Integer
        NoGroup
        SingleGroup
        DoubleGroup
    End Enum

    Private Sub FrmLaporanOvertimeAnalysis2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboFilterDate.SelectedIndex = 1
            Me.cboFormat.SelectedIndex = 0
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ChkShowOther.Checked = True
            Me.ChkShowSubTotal.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim DS As New DseLaporanOvertimeAnalysis2
        Dim Dta As New DseLaporanOvertimeAnalysis2TableAdapters.TblSalaryOTTableAdapter
        Dim Tbl As New DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable

        Try
            Me.Cursor = Cursors.WaitCursor

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            '1. cek filter berdasarkan actual date atau salary calculation date
            If Me.CboFilterDate.SelectedIndex = 0 Then
                Me.DateFrom = Me.dtpFrom.Value
                Me.DateTo = Me.dtpTo.Value
                Dta.FillBy_DateFromTo(Tbl, Me.DateTo, Me.DateFrom)
            Else
                Me.DateFrom = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
                Me.DateTo = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value.Date)
                Dta.FillBy_SalaryCalculation(Tbl, Me.DateTo, Me.DateFrom)
            End If

            If Tbl.Count > 0 Then
                'Me.GetOvertimeAnalisisWithoutOT(Tbl)
                Me.GenerateOvertimeAnalisisReport_Filter(Tbl)
                If Tbl.Count > 0 Then
                    Me.FormatDecimal = Me.ObjCurr.GetDecimal_DefaultCurrency
                    'If Me.cboChart_DataType.SelectedIndex = 0 Then
                    '    Me.FormatDecimal = Me.ObjCurr.GetDecimal_DefaultCurrency
                    'Else
                    '    Me.FormatDecimal = 2
                    'End If

                    If Me.cboFormat.SelectedIndex = 0 Then
                        '2. cek apakah tampilkan sub total group 1 ataupun group 2
                        Me.GenerateOvertimeAnalisisReport_Data(Tbl, DS)
                    Else
                        '3. cek apakah show other data untuk tipe chart line ataupun pie
                        '4. cek urutan top to low untuk tipe chart line ataupun pie
                        Me.GenerateOvertimeAnalisisReport_Chart(Tbl, DS)
                    End If
                End If
            End If

            If Tbl.Count = 0 Then
                MsgBox("There is no data", MsgBoxStyle.Information, "Information")
                Me.crvLaporanAnalysisOvertime.ReportSource = Nothing
                Me.crvLaporanAnalysisOvertime.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            Dta.Dispose()
            Tbl.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub GetOvertimeAnalisisWithoutOT(ByVal Tbl As DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable)
        Dim DtaK As New DseLaporanOvertimeAnalysis2TableAdapters.tblmaster_karyawanTableAdapter
        Dim TblK As New DseLaporanOvertimeAnalysis2.tblmaster_karyawanDataTable
        Dim rowK As DseLaporanOvertimeAnalysis2.tblmaster_karyawanRow
        Dim row As DseLaporanOvertimeAnalysis2.TblSalaryOTRow

        Try
            DtaK.Connection.ConnectionString = Core_Module.GetConstring
            DtaK.Fill(TblK)
            For Each rowK In TblK
                Tbl.DefaultView.RowFilter = "[Primary Employee] = " & rowK.Primary_Employee
                If Tbl.DefaultView.Count = 0 Then
                    For i As Integer = 0 To DateDiff(DateInterval.Month, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                        row = Tbl.NewTblSalaryOTRow
                        With row
                            ._date = Me.dtpFrom.Value.Date.AddMonths(i)
                            .Primary_Employee = rowK.Primary_Employee
                            .primary_departemen = rowK.primary_departemen
                            .primary_position = rowK.primary_position
                            .Employee_Code = rowK.Employee_Code
                            .Employee_NIK = rowK.Employee_NIK
                            .Employee_Name = rowK.Employee_Name
                            .Dept_Code = row.Dept_Code
                            .Dept_Desc = rowK.Dept_Desc
                            .Job_Position_Code = rowK.Job_Position_Code
                            .Job_Position_Desc = rowK.Job_Position_Desc
                        End With
                        Tbl.Rows.Add(row)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dtak.Dispose()
            TblK.Dispose()
        End Try
    End Sub

    Private Sub GenerateOvertimeAnalisisReport_Filter(ByVal Tbl As DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable)
        Dim TblDummy As New DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable
        Dim rowD As DataGridViewRow
        Dim Filter As String = ""
        Dim FilterDept As String = ""
        Dim FilterJob As String = ""
        Dim FilterEmp As String = ""

        Try
            TblDummy = Tbl.Copy
            Tbl.Clear()

            For Each rowD In Me.dgvDepartment.Rows
                If rowD.Cells(Me.ColDeptPrimary.Index).Value <> -1 Then
                    If FilterDept = "" Then
                        FilterDept = "[primary departemen] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    Else
                        FilterDept &= " OR [primary departemen] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    End If
                End If
            Next

            For Each rowD In Me.dgvJobPosition.Rows
                If rowD.Cells(Me.ColJobPosisitonPrimary.Index).Value <> -1 Then
                    If FilterJob = "" Then
                        FilterJob = "[primary position] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    Else
                        FilterJob &= " OR [primary position] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    End If
                End If
            Next

            For Each rowD In Me.dgvEmployee.Rows
                If rowD.Cells(Me.ColEmployee_primary.Index).Value <> -1 Then
                    If FilterEmp = "" Then
                        FilterEmp = "[Primary Employee] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    Else
                        FilterEmp &= " OR [Primary Employee] = " & rowD.Cells(Me.ColDeptPrimary.Index).Value
                    End If
                End If
            Next

            If FilterDept <> "" Then
                Filter = "(" & FilterDept & ")"
            End If
            If FilterJob <> "" Then
                If Filter <> "" Then
                    Filter &= " AND "
                End If
                Filter &= "(" & FilterJob & ")"
            End If
            If FilterEmp <> "" Then
                If Filter <> "" Then
                    Filter &= " AND "
                End If
                Filter &= "(" & FilterEmp & ")"
            End If

            TblDummy.DefaultView.RowFilter = Filter
            If TblDummy.DefaultView.Count > 0 Then
                Tbl.Merge(TblDummy.DefaultView.ToTable)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateOvertimeAnalisisReport_Data(ByVal Tbl As DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable, _
    ByVal DS As DseLaporanOvertimeAnalysis2)
        Dim Report As New CrpLaporanOvertimeAnalysis_Data
        Dim TblC As New DseLaporanOvertimeAnalysis2.TblCetak_DataDataTable
        Dim TblCSort As DseLaporanOvertimeAnalysis2.TblCetak_DataDataTable = DS.TblCetak_Data
        Dim TblST As New DseLaporanOvertimeAnalysis2.TblCetak_Data_SubTotalDataTable
        Dim rowST As DseLaporanOvertimeAnalysis2.TblCetak_Data_SubTotalRow
        Dim rowC As DseLaporanOvertimeAnalysis2.TblCetak_DataRow
        Dim row As DseLaporanOvertimeAnalysis2.TblSalaryOTRow

        Dim EnumDataGrouping As EnumTipeData
        Dim EnumGroup_1 As Integer = -1
        Dim EnumGroup_2 As Integer = -1
        Dim EnumDataValues As Integer = -1
        Dim PrimaryAcuan1, PrimaryAcuan2, PrimaryAcuan3 As Double
        Dim Group_1, Group_2, DataValues As String
        Dim ParamGroup_1 As String = ""
        Dim ParamGroup_2 As String = ""
        Dim ParamHeader As String = ""
        Dim ParamTotalWeekDayOT As String = ""
        Dim ParamTotalWeekEndOT As String = ""
        Dim ParamTotalOT As String = ""
        Dim ParamShowGroup_I As Boolean = False
        Dim ParamShowGroup_II As Boolean = False
        Dim DateFilter As Date = Now
        Dim AddFilterYear As String = ""

        Try
            EnumDataGrouping = Me.GetTipeGrouping_Enum(Me.cboGrouping.Text)
            If EnumDataGrouping = EnumTipeGrouping.NoGroup Then
                EnumDataValues = Me.GetTipe_Enum(Me.cboData1.Text)
                ParamHeader = Me.cboData1.Text
            ElseIf EnumDataGrouping = EnumTipeGrouping.SingleGroup Then
                ParamGroup_1 = Me.cboData1.Text
                EnumGroup_1 = Me.GetTipe_Enum(Me.cboData1.Text)
                EnumDataValues = Me.GetTipe_Enum(Me.cboData2.Text)
                ParamHeader = Me.cboData2.Text
                If Me.ChkShowSubTotal.Checked = True Then
                    ParamShowGroup_I = True
                End If
            ElseIf EnumDataGrouping = EnumTipeGrouping.DoubleGroup Then
                ParamGroup_1 = Me.cboData1.Text
                ParamGroup_2 = Me.cboData2.Text
                EnumGroup_1 = Me.GetTipe_Enum(Me.cboData1.Text)
                EnumGroup_2 = Me.GetTipe_Enum(Me.cboData2.Text)
                EnumDataValues = Me.GetTipe_Enum(Me.cboData3.Text)
                ParamHeader = Me.cboData3.Text
                If Me.ChkShowSubTotal.Checked = True Then
                    ParamShowGroup_I = True
                    ParamShowGroup_II = True
                End If
            End If

            For Each row In Tbl
                AddFilterYear = ""
                PrimaryAcuan1 = -1
                Group_1 = ""
                If Me.CboFilterDate.SelectedIndex = 0 Then
                    DateFilter = row._date
                Else
                    DateFilter = row.End_Date_Salary
                End If

                If EnumGroup_1 = EnumTipeData.Department Then
                    PrimaryAcuan1 = row.primary_departemen
                    Group_1 = row.Dept_Desc
                ElseIf EnumGroup_1 = EnumTipeData.Employee Then
                    PrimaryAcuan1 = row.Primary_Employee
                    Group_1 = row.Employee_Name
                ElseIf EnumGroup_1 = EnumTipeData.Job_Position Then
                    PrimaryAcuan1 = row.primary_position
                    Group_1 = row.Job_Position_Desc
                ElseIf EnumGroup_1 = EnumTipeData.Monthly_Period Then
                    PrimaryAcuan1 = CDbl(DateFilter.Month) 'CDbl(Format(row._date, "MM"))
                    Group_1 = Me.ObjCD.GetMonthOfTheYear(CInt(DateFilter.Month), True) & " " & DateFilter.Year
                    AddFilterYear = " AND [Years_] = " & CDbl(DateFilter.Year)
                End If

                PrimaryAcuan2 = -1
                Group_2 = ""
                If EnumGroup_2 = EnumTipeData.Department Then
                    PrimaryAcuan2 = row.primary_departemen
                    Group_2 = row.Dept_Desc
                ElseIf EnumGroup_2 = EnumTipeData.Employee Then
                    PrimaryAcuan2 = row.Primary_Employee
                    Group_2 = row.Employee_Name
                ElseIf EnumGroup_2 = EnumTipeData.Job_Position Then
                    PrimaryAcuan2 = row.primary_position
                    Group_2 = row.Job_Position_Desc
                ElseIf EnumGroup_2 = EnumTipeData.Monthly_Period Then
                    PrimaryAcuan2 = CDbl(DateFilter.Month) 'CDbl(Format(row._date, "MM"))
                    Group_2 = Me.ObjCD.GetMonthOfTheYear(CInt(DateFilter.Month), True) & " " & DateFilter.Year
                    AddFilterYear = " AND [Years_] = " & CDbl(DateFilter.Year)
                End If

                PrimaryAcuan3 = -1
                DataValues = ""
                If EnumDataValues = EnumTipeData.Department Then
                    PrimaryAcuan3 = row.primary_departemen
                    DataValues = row.Dept_Desc
                ElseIf EnumDataValues = EnumTipeData.Employee Then
                    PrimaryAcuan3 = row.Primary_Employee
                    DataValues = row.Employee_Name
                ElseIf EnumDataValues = EnumTipeData.Job_Position Then
                    PrimaryAcuan3 = row.primary_position
                    DataValues = row.Job_Position_Desc
                ElseIf EnumDataValues = EnumTipeData.Monthly_Period Then
                    PrimaryAcuan3 = CDbl(DateFilter.Month) 'CDbl(Format(row._date, "MM"))
                    DataValues = Me.ObjCD.GetMonthOfTheYear(CInt(DateFilter.Month), True) & " " & DateFilter.Year
                    AddFilterYear = " AND [Years_] = " & CDbl(DateFilter.Year)
                End If

                TblC.DefaultView.RowFilter = "[Primary Acuan I] = " & PrimaryAcuan1 & _
                " AND [Primary Acuan II] = " & PrimaryAcuan2 & " AND [Primary Acuan III] = " & PrimaryAcuan3 & AddFilterYear
                If TblC.DefaultView.Count > 0 Then
                    rowC = TblC.DefaultView(0).Row
                    If row.tipe_ot = Core_Module.ClsMasterSalaryFormula.EnumTipe_OT.Weekday Then
                        rowC.Week_Day_OT += row.duration
                        rowC.Week_Day_Amount += row.amount
                    Else
                        rowC.Week_End_OT += row.duration
                        rowC.Week_End_Amount += row.amount
                    End If
                    rowC.Total_OT = rowC.Week_Day_OT + rowC.Week_End_OT
                    rowC.Total_Amount = rowC.Week_Day_Amount + rowC.Week_End_Amount
                Else
                    rowC = TblC.NewTblCetak_DataRow
                    rowC.Group_I = Group_1
                    rowC.Group_II = Group_2
                    rowC.Data_Value = DataValues
                    rowC.Years_ = DateFilter.Year
                    rowC.Primary_Acuan_I = PrimaryAcuan1
                    rowC.Primary_Acuan_II = PrimaryAcuan2
                    rowC.Primary_Acuan_III = PrimaryAcuan3
                    If row.tipe_ot = Core_Module.ClsMasterSalaryFormula.EnumTipe_OT.Weekday Then
                        rowC.Week_Day_OT = row.duration
                        rowC.Week_Day_Amount = row.amount
                    Else
                        rowC.Week_End_OT = row.duration
                        rowC.Week_End_Amount = row.amount
                    End If
                    rowC.Total_OT = rowC.Week_Day_OT + rowC.Week_End_OT
                    rowC.Total_Amount = rowC.Week_Day_Amount + rowC.Week_End_Amount
                    TblC.Rows.Add(rowC)
                End If
            Next

            If TblC.Count > 0 Then
                'hitung subtotal
                If Me.ChkShowSubTotal.Checked = True Then
                    For Each rowC In TblC
                        With rowC
                            TblST.DefaultView.RowFilter = "[Group I] = '" & .Group_I.ToString & "' AND [Group II] = '" & .Group_II.ToString & "'"
                            If TblST.DefaultView.Count > 0 Then
                                rowST = TblST.DefaultView(0).Row
                                rowST.Total_Week_Day_OT_GI += .Week_Day_OT
                                rowST.Total_Week_Day_OT_GII += .Week_Day_OT
                                rowST.Total_Week_End_OT_GI += .Week_End_OT
                                rowST.Total_Week_End_OT_GII += .Week_End_OT
                                rowST.Total_OT_GI_ += .Total_OT
                                rowST.Total_OT_GII_ += .Total_OT
                            Else
                                rowST = TblST.NewTblCetak_Data_SubTotalRow
                                rowST.Group_I = .Group_I
                                rowST.Group_II = .Group_II
                                rowST.Total_Week_Day_OT_GI = .Week_Day_OT
                                rowST.Total_Week_Day_OT_GII = .Week_Day_OT
                                rowST.Total_Week_End_OT_GI = .Week_End_OT
                                rowST.Total_Week_End_OT_GII = .Week_End_OT
                                rowST.Total_OT_GI_ = .Total_OT
                                rowST.Total_OT_GII_ = .Total_OT
                                TblST.Rows.Add(rowST)
                            End If
                        End With
                    Next
                End If

                'convert duration 
                Dim TotalWeekDayOT As Decimal = 0
                Dim TotalWeekEndOT As Decimal = 0
                Dim TotalOT As Decimal = 0
                For Each rowC In TblC
                    With rowC
                        If .Week_Day_OT > 0 Then
                            TotalWeekDayOT += .Week_Day_OT
                            .Week_Day_OT_Str = Me.ObjCD.Convert_ToString(.Week_Day_OT)
                        End If
                        If .Week_End_OT > 0 Then
                            TotalWeekEndOT += .Week_End_OT
                            .Week_End_OT_Str = Me.ObjCD.Convert_ToString(.Week_End_OT)
                        End If
                        If .Total_OT > 0 Then
                            TotalOT += .Total_OT
                            .Total_OT_Str = Me.ObjCD.Convert_ToString(.Total_OT)
                        End If

                        'jika show subtotal maka convert subtotal
                        If Me.ChkShowSubTotal.Checked = True And TblST.Count > 0 Then
                            If ParamShowGroup_I = True Then
                                Dim TotalWeekDayOT_Dumm As Decimal = 0
                                Dim TotalWeekEndOT_Dumm As Decimal = 0
                                Dim TotalOT_Dumm As Decimal = 0
                                TblST.DefaultView.RowFilter = "[Group I] = '" & .Group_I.ToString & "'"
                                For Each dv As DataRowView In TblST.DefaultView
                                    rowST = dv.Row
                                    TotalWeekDayOT_Dumm += rowST.Total_Week_Day_OT_GI
                                    TotalWeekEndOT_Dumm += rowST.Total_Week_End_OT_GI
                                    TotalOT_Dumm += rowST.Total_OT_GI_
                                Next
                                If TotalWeekDayOT_Dumm > 0 Then
                                    .Total_Week_Day_OT_GI_Str = Me.ObjCD.Convert_ToString(TotalWeekDayOT_Dumm)
                                End If
                                If TotalWeekEndOT_Dumm > 0 Then
                                    .Total_Week_End_OT_GI_Str = Me.ObjCD.Convert_ToString(TotalWeekEndOT_Dumm)
                                End If
                                If TotalOT_Dumm > 0 Then
                                    .Total_OT_GI_Str = Me.ObjCD.Convert_ToString(TotalOT_Dumm)
                                End If
                            End If

                            If ParamShowGroup_II = True Then
                                TblST.DefaultView.RowFilter = "[Group I] = '" & .Group_I.ToString & "' AND [Group II] = '" & .Group_II.ToString & "'"
                                If TblST.DefaultView.Count > 0 Then
                                    rowST = TblST.DefaultView(0).Row
                                    If rowST.Total_Week_Day_OT_GII > 0 Then
                                        .Total_Week_Day_OT_GII_Str = Me.ObjCD.Convert_ToString(rowST.Total_Week_Day_OT_GII)
                                    End If
                                    If rowST.Total_Week_End_OT_GII > 0 Then
                                        .Total_Week_End_OT_GII_Str = Me.ObjCD.Convert_ToString(rowST.Total_Week_End_OT_GII)
                                    End If
                                    If rowST.Total_OT_GII_ > 0 Then
                                        .Total_OT_GII_Str = Me.ObjCD.Convert_ToString(rowST.Total_OT_GII_)
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next

                'Isi parameter total overtime
                If TotalWeekDayOT > 0 Then
                    ParamTotalWeekDayOT = Me.ObjCD.DigitGrouping(Me.ObjCD.Convert_ToString(TotalWeekDayOT), 2)
                End If
                If TotalWeekEndOT > 0 Then
                    ParamTotalWeekEndOT = Me.ObjCD.DigitGrouping(Me.ObjCD.Convert_ToString(TotalWeekEndOT), 2)
                End If
                If TotalOT > 0 Then
                    ParamTotalOT = Me.ObjCD.DigitGrouping(Me.ObjCD.Convert_ToString(TotalOT), 2)
                End If

                Dim Sorting_1 As String = "[Group I] ASC, "
                Dim Sorting_2 As String = "[Group II] ASC, "
                Dim Sorting_3 As String = "[Data Value] ASC"
                If EnumGroup_1 = EnumTipeData.Monthly_Period Then
                    Sorting_1 = "[Years_] ASC, [Primary Acuan I] ASC, "
                End If
                If EnumGroup_2 = EnumTipeData.Monthly_Period Then
                    Sorting_2 = "[Years_] ASC, [Primary Acuan II] ASC, "
                End If
                If EnumDataValues = EnumTipeData.Monthly_Period Then
                    Sorting_3 = "[Years_] ASC, [Primary Acuan III] ASC"
                End If
                TblC.DefaultView.RowFilter = Nothing
                TblC.DefaultView.Sort = Sorting_1 & Sorting_2 & Sorting_3
                TblCSort.Merge(TblC.DefaultView.ToTable)
            End If

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_StartMonth.ParameterFieldName, Me.DateFrom.Date)
            Report.SetParameterValue(Report.Parameter_EndMonth.ParameterFieldName, Me.DateTo.Date)
            Report.SetParameterValue(Report.Parameter_DecPrec.ParameterFieldName, Me.FormatDecimal)
            Report.SetParameterValue(Report.Parameter_ParamGroup_1.ParameterFieldName, ParamGroup_1)
            Report.SetParameterValue(Report.Parameter_ParamGroup_2.ParameterFieldName, ParamGroup_2)
            Report.SetParameterValue(Report.Parameter_ParamHeader.ParameterFieldName, ParamHeader)
            Report.SetParameterValue(Report.Parameter_ParamShowSubTotalG_I.ParameterFieldName, ParamShowGroup_I)
            Report.SetParameterValue(Report.Parameter_ParamShowSubTotalG_II.ParameterFieldName, ParamShowGroup_II)
            Report.SetParameterValue(Report.Parameter_ParamTotalWeekdayOT.ParameterFieldName, ParamTotalWeekDayOT)
            Report.SetParameterValue(Report.Parameter_ParamTotalWeekendOT.ParameterFieldName, ParamTotalWeekEndOT)
            Report.SetParameterValue(Report.Parameter_ParamTotalOT.ParameterFieldName, ParamTotalOT)
            Me.crvLaporanAnalysisOvertime.ReportSource = Report
            Me.crvLaporanAnalysisOvertime.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblC.Dispose()
            TblCSort.Dispose()
            TblST.Dispose()
        End Try
    End Sub

    Private Sub GenerateOvertimeAnalisisReport_Chart(ByVal Tbl As DseLaporanOvertimeAnalysis2.TblSalaryOTDataTable, _
    ByVal DS As DseLaporanOvertimeAnalysis2)
        Dim TblC, TblDummy, TblDummy2 As New DseLaporanOvertimeAnalysis2.TblCetak_ChartDataTable
        Dim TblCSort As DseLaporanOvertimeAnalysis2.TblCetak_ChartDataTable = DS.TblCetak_Chart
        Dim rowC, rowDummy, rowDummy2 As DseLaporanOvertimeAnalysis2.TblCetak_ChartRow
        Dim row As DseLaporanOvertimeAnalysis2.TblSalaryOTRow

        Dim EnumDataGroup As EnumTipeData = Me.GetTipe_Enum(Me.cboData1.Text)
        Dim DataGroup As String = ""
        Dim AmountOrDuration As Decimal = 0
        Dim PrimaryAcuan1 As Double = -1
        Dim AddDate As Date = #1/1/1950#
        Dim AddFilterDate As String = ""
        Dim ParamDecimal As Integer = ObjCurr.GetDecimal_DefaultCurrency
        Dim ListTopPrimary As New List(Of Double)
        Dim ParamDescription As String = ""
        Dim TotalAmount As Decimal = 0

        Try
            For Each row In Tbl
                If Me.CboFilterDate.SelectedIndex = 0 Then
                    AddDate = Me.ObjCD.GetFirstDateofTheMonth(row._date)
                Else
                    AddDate = Me.ObjCD.GetFirstDateofTheMonth(row.End_Date_Salary)
                End If

                If Me.cboChartType.Text = "Line" Then
                    AddFilterDate = " AND [Date] = #" & AddDate & "#"
                Else
                    AddFilterDate = ""
                End If

                If Me.cboChart_DataType.Text = "OT Amount" Then
                    AmountOrDuration = row.amount
                    ParamDescription = "Amount"
                Else
                    AmountOrDuration = row.duration
                    ParamDecimal = 1
                    ParamDescription = "Duration"
                End If

                If EnumDataGroup = EnumTipeData.Department Then
                    PrimaryAcuan1 = row.primary_departemen
                    DataGroup = row.Dept_Desc
                ElseIf EnumDataGroup = EnumTipeData.Employee Then
                    PrimaryAcuan1 = row.Primary_Employee
                    DataGroup = row.Employee_Name
                ElseIf EnumDataGroup = EnumTipeData.Job_Position Then
                    PrimaryAcuan1 = row.primary_position
                    DataGroup = row.Job_Position_Desc
                ElseIf EnumDataGroup = EnumTipeData.Monthly_Period Then
                    PrimaryAcuan1 = CDbl(AddDate.Month & AddDate.Year)
                    DataGroup = Me.ObjCD.GetMonthOfTheYear(CInt(AddDate.Month), True) & " " & AddDate.Year
                End If

                TotalAmount += AmountOrDuration
                TblC.DefaultView.RowFilter = "[Primary Acuan I] = " & PrimaryAcuan1 & AddFilterDate
                If TblC.DefaultView.Count > 0 Then
                    rowC = TblC.DefaultView(0).Row
                    rowC.Amount += AmountOrDuration
                Else
                    rowC = TblC.NewTblCetak_ChartRow
                    rowC.Primary_Acuan_I = PrimaryAcuan1
                    rowC._Date = AddDate
                    rowC.Series_I = DataGroup
                    rowC.Amount = AmountOrDuration
                    TblC.Rows.Add(rowC)
                End If
            Next

            'fungsi filter untuk menampilkan top data
            If Me.cboChart_DataSelection.Text = "Top" Then
                For Each rowC In TblC
                    TblDummy.DefaultView.RowFilter = "[Primary Acuan I] = " & rowC.Primary_Acuan_I
                    If TblDummy.DefaultView.Count > 0 Then
                        rowDummy = TblDummy.DefaultView(0).Row
                        rowDummy.Amount += rowC.Amount
                    Else
                        rowDummy = TblDummy.NewTblCetak_ChartRow
                        For i As Integer = 0 To TblC.Columns.Count - 1
                            rowDummy.Item(i) = rowC.Item(i)
                        Next
                        TblDummy.Rows.Add(rowDummy)
                    End If
                Next

                TblDummy.DefaultView.RowFilter = Nothing
                TblDummy.DefaultView.Sort = "[Amount] DESC"
                Dim rowMax As Integer = Me.TxtChart_DataSelection.Text
                If rowMax > TblDummy.DefaultView.Count Then
                    rowMax = TblDummy.Count
                End If

                TotalAmount = 0
                For i As Integer = 0 To TblDummy.DefaultView.Count - 1
                    rowDummy = TblDummy.DefaultView(i).Row
                    TblC.DefaultView.RowFilter = "[Primary Acuan I] = " & rowDummy.Primary_Acuan_I
                    For ii As Integer = TblC.DefaultView.Count - 1 To 0 Step -1
                        rowC = TblC.DefaultView(ii).Row
                        If i < rowMax Then
                            rowC.Urutan_By_Amount = i
                            TotalAmount += rowC.Amount
                        Else
                            If Me.ChkShowOther.Checked = True Then
                                TotalAmount += rowC.Amount
                                If Me.cboChartType.Text = "Line" Then
                                    TblDummy2.DefaultView.RowFilter = "[Primary Acuan I] = -1 AND [Date] = #" & rowC._Date & "#"
                                Else
                                    TblDummy2.DefaultView.RowFilter = "[Primary Acuan I] = -1"
                                End If

                                If TblDummy2.DefaultView.Count > 0 Then
                                    rowDummy2 = TblDummy2.DefaultView(0).Row
                                    rowDummy2.Amount += rowC.Amount
                                Else
                                    rowDummy2 = TblDummy2.NewTblCetak_ChartRow
                                    rowDummy2.Primary_Acuan_I = -1
                                    rowDummy2._Date = rowC._Date
                                    rowDummy2.Series_I = "OTHER"
                                    rowDummy2.Amount = rowC.Amount
                                    rowDummy2.Urutan_By_Amount = rowMax
                                    TblDummy2.Rows.Add(rowDummy2)
                                End If
                            End If

                            TblC.Rows.Remove(rowC)
                        End If
                    Next
                Next

                If Me.ChkShowOther.Checked = True Then
                    TblC.Merge(TblDummy2)
                End If
            End If

            If Me.cboChartType.Text = "Pie" Or Me.cboChart_DataType.Text = "OT Duration" Then
                For Each rowC In TblC
                    If Me.cboChartType.Text = "Pie" Then
                        'jika chart tipe pie maka isi kolom percentage
                        rowC.Percentage = Math.Round(rowC.Amount * 100 / TotalAmount, 2, MidpointRounding.AwayFromZero)
                    End If
                    If Me.cboChart_DataType.Text = "OT Duration" Then
                        'convert duration 
                        If rowC.Amount > 0 Then
                            rowC.Amount = Me.ObjCD.DigitGrouping(rowC.Amount / 60, 1)
                        End If
                    End If
                Next
            End If

            Dim Sorting_1 As String = "[Urutan By Amount] ASC, [Series I] ASC"
            If EnumDataGroup = EnumTipeData.Monthly_Period Then
                Sorting_1 = "[Urutan By Amount] ASC, [Date] ASC"
            End If

            TblC.DefaultView.RowFilter = Nothing
            TblC.DefaultView.Sort = Sorting_1
            TblCSort.Merge(TblC.DefaultView.ToTable)

            Dim Count_ As Integer = 1
            Dim CountString_ As String = ""
            For i As Integer = 0 To TblCSort.Rows.Count - 1
                rowC = TblCSort.Rows(i)
                If rowC.Urutan_By_Amount <> -1 Then
                    Count_ = rowC.Urutan_By_Amount + 1
                Else
                    If i > 0 Then
                        rowDummy = TblCSort.Rows(i - 1)
                        If rowDummy.Primary_Acuan_I <> rowC.Primary_Acuan_I Then
                            Count_ += 1
                        End If
                    End If
                End If

                If Count_.ToString.Length = 1 Then
                    rowC.Series_I = "0" & Count_ & ". " & rowC.Series_I
                Else
                    rowC.Series_I = Count_ & ". " & rowC.Series_I
                End If
            Next

            If Me.cboChartType.Text = "Line" Then
                Count_ = 1
                CountString_ = ""
                TblCSort.DefaultView.RowFilter = Nothing
                TblCSort.DefaultView.Sort = "[Date] ASC"
                For i As Integer = 0 To TblCSort.DefaultView.Count - 1
                    rowC = TblCSort.DefaultView(i).Row
                    If i > 0 Then
                        rowDummy = TblCSort.DefaultView(i - 1).Row
                        If rowC._Date <> rowDummy._Date Then
                            Count_ += 1
                        End If
                    End If
                    CountString_ = Count_.ToString
                    If CountString_.Length = 1 Then
                        CountString_ = "0" & Count_.ToString
                    End If
                    rowC.Date_Str = CountString_ + ". " + Format(rowC._Date, "MMM yyyy")
                Next
            End If

            If Me.cboChartType.Text = "Line" Then
                Dim Report As New CrpLaporanOvertimeAnalysis_Chart_Line
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_StartMonth.ParameterFieldName, Me.DateFrom.Date)
                Report.SetParameterValue(Report.Parameter_EndMonth.ParameterFieldName, Me.DateTo.Date)
                Report.SetParameterValue(Report.Parameter_ParamDec.ParameterFieldName, ParamDecimal)
                Report.SetParameterValue(Report.Parameter_ParamSumDesc.ParameterFieldName, ParamDescription)
                Me.crvLaporanAnalysisOvertime.ReportSource = Report
                Me.crvLaporanAnalysisOvertime.Show()
            Else
                Dim Report As New CrpLaporanOvertimeAnalysis_Chart_Pie
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_StartMonth.ParameterFieldName, Me.DateFrom.Date)
                Report.SetParameterValue(Report.Parameter_EndMonth.ParameterFieldName, Me.DateTo.Date)
                Report.SetParameterValue(Report.Parameter_Report_Detail_ParamDecimal.ParameterFieldName, ParamDecimal)
                Report.SetParameterValue(Report.Parameter_Report_Detail_ParamDesc.ParameterFieldName, ParamDescription)
                Me.crvLaporanAnalysisOvertime.ReportSource = Report
                Me.crvLaporanAnalysisOvertime.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblC.Dispose()
            TblCSort.Dispose()
            TblDummy.Dispose()
            TblDummy2.Dispose()
        End Try
    End Sub

    Private Sub cboFormat_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedValueChanged
        Try
            Me.cboChartType.SelectedIndex = 0
            Me.ChkShowSubTotal.Enabled = False
            Me.ChkShowOther.Enabled = False

            If Me.cboFormat.Text = "Data" Then
                Me.cboGrouping.Text = Me.GetTipeGrouping_Str(EnumTipeGrouping.NoGroup)
                Me.cboGrouping.Enabled = True
                Me.LblDataGrouping.Visible = True
                Me.cboGrouping.Visible = True
                Me.cboGrouping.TabStop = True
                Me.LblChartType.Visible = False
                Me.cboChartType.Visible = False
                Me.cboChartType.TabStop = False

                Me.LblChart_DataSelection.Visible = False
                Me.cboChart_DataSelection.Visible = False
                Me.cboChart_DataSelection.TabStop = False
                Me.TxtChart_DataSelection.Visible = False
                Me.TxtChart_DataSelection.TabStop = False

                Me.LblChart_DataType.Visible = False
                Me.cboChart_DataType.Visible = False
                Me.cboChart_DataType.TabStop = False

                Me.LblData1.Text = "Data Values"
                Me.LblData2.Visible = True
                Me.cboData2.Visible = True
                Me.cboData2.TabStop = True
                Me.LblData3.Visible = True
                Me.cboData3.Visible = True
                Me.cboData3.TabStop = True

                'Me.btnShow.Top = Me.cboData3.Top + Me.cboData3.Height + 5
                'Me.btnClose.Top = Me.btnShow.Top
            Else
                Me.cboGrouping.Text = Me.GetTipeGrouping_Str(EnumTipeGrouping.NoGroup)
                Me.cboGrouping.Enabled = False
                Me.LblDataGrouping.Visible = False
                Me.cboGrouping.Visible = False
                Me.cboGrouping.TabStop = False
                Me.LblChartType.Visible = True
                Me.cboChartType.Visible = True
                Me.cboChartType.TabStop = True
                Me.cboChartType.SelectedIndex = 0

                Me.LblChart_DataSelection.Visible = True
                Me.cboChart_DataSelection.Visible = True
                Me.cboChart_DataSelection.TabStop = True
                Me.cboChart_DataSelection.SelectedIndex = 0
                Me.TxtChart_DataSelection.Visible = True
                Me.TxtChart_DataSelection.TabStop = True
                Me.TxtChart_DataSelection.Text = 0

                Me.LblChart_DataType.Visible = True
                Me.cboChart_DataType.Visible = True
                Me.cboChart_DataType.TabStop = True
                Me.cboChart_DataType.SelectedIndex = 0

                Me.LblData2.Visible = False
                Me.cboData2.Visible = False
                Me.cboData2.TabStop = False
                Me.cboData2.SelectedIndex = 0
                Me.LblData3.Visible = False
                Me.cboData3.Visible = False
                Me.cboData3.TabStop = False
                Me.cboData3.SelectedIndex = 0

                Me.SetCboDataChart()


                'Me.btnShow.Top = Me.cboData2.Top + Me.cboData2.Height + 5
                'Me.btnClose.Top = Me.btnShow.Top
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGrouping_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGrouping.SelectedValueChanged
        Try
            If Me.cboFormat.Text = "Chart" Then
                SetCboDataChart()
            Else
                If Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.NoGroup Then
                    Me.cboData1.Items.Clear()
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Department))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Employee))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Job_Position))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Monthly_Period))
                    Me.cboData1.SelectedIndex = 0
                    Me.cboData1.Enabled = True
                    Me.SetCBOUnactive(Me.cboData2)
                    Me.SetCBOUnactive(Me.cboData3)
                    Me.LblData1.Text = "Data Values"
                    Me.LblData2.Text = "n/a"
                    Me.LblData3.Text = "n/a"
                    Me.ChkShowSubTotal.Enabled = False
                ElseIf Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.SingleGroup Then
                    Me.cboData1.Items.Clear()
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Department))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Employee))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Job_Position))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Monthly_Period))
                    Me.cboData1.SelectedIndex = 0
                    Me.cboData1.Enabled = True
                    Me.cboData2.Enabled = True
                    Me.SetCBOUnactive(Me.cboData3)
                    Me.LblData1.Text = "Group By"
                    Me.LblData2.Text = "Data Values"
                    Me.LblData3.Text = "n/a"
                    Me.ChkShowSubTotal.Enabled = True
                ElseIf Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.DoubleGroup Then
                    Me.cboData1.Items.Clear()
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Department))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Employee))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Job_Position))
                    Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Monthly_Period))
                    Me.cboData1.SelectedIndex = 0
                    Me.cboData1.Enabled = True
                    Me.cboData2.Enabled = True
                    Me.cboData3.Enabled = True
                    Me.LblData1.Text = "1st Group"
                    Me.LblData2.Text = "2nd Group"
                    Me.LblData3.Text = "Data Values"
                    Me.ChkShowSubTotal.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboChartType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboChartType.SelectedValueChanged
        Try
            SetCboDataChart()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboData1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboData1.SelectedValueChanged
        Dim List1 As New List(Of EnumTipeData)

        Try
            If Me.cboFormat.Text = "Chart" Then
                'do nothing, sudah diatue semua pada cbo grouping selected value changed
            Else
                If Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.SingleGroup Or _
                Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.DoubleGroup Then
                    List1.Add(Me.GetTipe_Enum(Me.cboData1.Text))
                    Me.IsiCBOData(List1, Me.cboData2)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboData2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboData2.SelectedValueChanged
        Dim List1 As New List(Of EnumTipeData)

        Try
            If Me.cboFormat.Text = "Chart" Then
                'do nothing, sudah diatue semua pada cbo grouping selected value changed
            Else
                If Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.DoubleGroup Then
                    List1.Add(Me.GetTipe_Enum(Me.cboData1.Text))
                    List1.Add(Me.GetTipe_Enum(Me.cboData2.Text))
                    Me.IsiCBOData(List1, Me.cboData3)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboData3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboData3.SelectedIndexChanged

    End Sub

    Private Sub cboChart_DataSelection_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboChart_DataSelection.SelectedValueChanged
        Try
            Me.SetDataSelection()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboChart_DataType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboChart_DataType.SelectedIndexChanged

    End Sub

    Private Sub IsiCBOData(ByVal vListExcpetionItem As List(Of EnumTipeData), ByVal vCBO As Windows.Forms.ComboBox)
        Dim vCboName As String = vCBO.Name

        Try
            vCBO.Items.Clear()
            If Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.SingleGroup Then
                If Me.GetTipe_Enum(Me.cboData1.Text) = EnumTipeData.Employee And vCboName = "cboData2" Then
                    vListExcpetionItem.Add(EnumTipeData.Department)
                    vListExcpetionItem.Add(EnumTipeData.Job_Position)
                End If
            ElseIf Me.GetTipeGrouping_Enum(Me.cboGrouping.Text) = EnumTipeGrouping.DoubleGroup Then
                'kondisi jika cbodata1 = employee maka cbodata2 tidak bisa monthly
                If Me.GetTipe_Enum(Me.cboData1.Text) = EnumTipeData.Employee And vCboName = "cboData2" Then
                    vListExcpetionItem.Add(EnumTipeData.Monthly_Period)
                ElseIf Me.GetTipe_Enum(Me.cboData1.Text) = EnumTipeData.Monthly_Period And vCboName = "cboData2" Then
                    vListExcpetionItem.Add(EnumTipeData.Employee)
                End If
                If vCboName = "cboData3" Then
                    If Me.GetTipe_Enum(Me.cboData1.Text) = EnumTipeData.Employee Then
                        'kondisi jika cbodata1 = employee maka pastikan untuk cbodata3 hanya bisa monthly period
                        If Me.GetTipe_Enum(Me.cboData2.Text) = EnumTipeData.Department Then
                            vListExcpetionItem.Add(EnumTipeData.Job_Position)
                        ElseIf Me.GetTipe_Enum(Me.cboData2.Text) = EnumTipeData.Job_Position Then
                            vListExcpetionItem.Add(EnumTipeData.Department)
                        End If
                    Else
                        If Me.GetTipe_Enum(Me.cboData2.Text) = EnumTipeData.Employee Then
                            vListExcpetionItem.Add(EnumTipeData.Department)
                            vListExcpetionItem.Add(EnumTipeData.Job_Position)
                        End If
                    End If
                End If
            End If

            If vListExcpetionItem.Contains(EnumTipeData.Department) = False Then
                vCBO.Items.Add(Me.GetTipe_Str(EnumTipeData.Department))
            End If
            If vListExcpetionItem.Contains(EnumTipeData.Employee) = False Then
                vCBO.Items.Add(Me.GetTipe_Str(EnumTipeData.Employee))
            End If
            If vListExcpetionItem.Contains(EnumTipeData.Job_Position) = False Then
                vCBO.Items.Add(Me.GetTipe_Str(EnumTipeData.Job_Position))
            End If
            If vListExcpetionItem.Contains(EnumTipeData.Monthly_Period) = False Then
                vCBO.Items.Add(Me.GetTipe_Str(EnumTipeData.Monthly_Period))
            End If

            vCBO.SelectedIndex = 0
            vCBO.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDataSelection()
        Try
            If Me.cboChart_DataSelection.Text = "All Data" Then
                Me.TxtChart_DataSelection.Enabled = False
                Me.TxtChart_DataSelection.Text = 0
                Me.cboChart_DataSelection.Width = Me.TxtChart_DataSelection.Left + Me.TxtChart_DataSelection.Width - Me.cboChart_DataSelection.Left
                Me.cboChart_DataSelection.BringToFront()
                Me.ChkShowOther.Enabled = False
            Else
                Me.TxtChart_DataSelection.Enabled = True
                Me.TxtChart_DataSelection.Text = "5"
                Me.cboChart_DataSelection.Width = Me.TxtChart_DataSelection.Left - Me.cboChart_DataSelection.Left - 5
                Me.ChkShowOther.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCBOUnactive(ByVal vCbo As Windows.Forms.ComboBox)
        Try
            vCbo.Items.Clear()
            vCbo.Items.Add("n/a")
            vCbo.SelectedIndex = 0
            vCbo.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboDataChart()
        Try
            Me.cboData1.Items.Clear()
            Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Department))
            Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Employee))
            Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Job_Position))
            If Me.cboChartType.Text <> "Line" Then
                Me.cboData1.Items.Add(Me.GetTipe_Str(EnumTipeData.Monthly_Period))
            End If
            Me.cboData1.SelectedIndex = 0
            Me.cboData1.Enabled = True
            Me.SetCBOUnactive(Me.cboData2)
            Me.SetCBOUnactive(Me.cboData3)
            Me.LblData1.Text = "Data Group"
            Me.LblData2.Text = "n/a"
            Me.LblData3.Text = "n/a"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTipeGrouping_Str(ByVal vValue As EnumTipeGrouping) As String
        Try
            If vValue = EnumTipeGrouping.NoGroup Then
                Return "No Grouping"
            ElseIf vValue = EnumTipeGrouping.SingleGroup Then
                Return "Single Grouping"
            ElseIf vValue = EnumTipeGrouping.DoubleGroup Then
                Return "Double Grouping"
            End If

            Return "No Grouping"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "No Grouping"
        End Try
    End Function

    Private Function GetTipeGrouping_Enum(ByVal vValue As String) As EnumTipeGrouping
        Try
            If vValue = "No Grouping" Then
                Return EnumTipeGrouping.NoGroup
            ElseIf vValue = "Single Grouping" Then
                Return EnumTipeGrouping.SingleGroup
            ElseIf vValue = "Double Grouping" Then
                Return EnumTipeGrouping.DoubleGroup
            End If

            Return EnumTipeGrouping.NoGroup
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeGrouping.NoGroup
        End Try
    End Function

    Private Function GetTipe_Enum(ByVal vValue As String) As EnumTipeData
        Try
            If vValue = "Department" Then
                Return EnumTipeData.Department
            ElseIf vValue = "Employee" Then
                Return EnumTipeData.Employee
            ElseIf vValue = "Job Position" Then
                Return EnumTipeData.Job_Position
            ElseIf vValue = "Monthly Period" Then
                Return EnumTipeData.Monthly_Period
            End If

            Return EnumTipeData.Department
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeData.Department
        End Try
    End Function

    Private Function GetTipe_Str(ByVal vValue As EnumTipeData) As String
        Try
            If vValue = EnumTipeData.Department Then
                Return "Department"
            ElseIf vValue = EnumTipeData.Employee Then
                Return "Employee"
            ElseIf vValue = EnumTipeData.Job_Position Then
                Return "Job Position"
            ElseIf vValue = EnumTipeData.Monthly_Period Then
                Return "Monthly Period"
            End If

            Return "Department"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Department"
        End Try
    End Function

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

    Private Sub TxtChart_DataSelection_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtChart_DataSelection.Enter
        Me.ObjCD.Number_Enter(Me.TxtChart_DataSelection, sender, e)
    End Sub

    Private Sub TxtChart_DataSelection_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtChart_DataSelection.GotFocus
        TxtChart_DataSelection.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtChart_DataSelection_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtChart_DataSelection.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtChart_DataSelection, sender, e)
    End Sub

    Private Sub TxtChart_DataSelection_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtChart_DataSelection.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtChart_DataSelection, sender, e)
    End Sub

    Private Sub TxtChart_DataSelection_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtChart_DataSelection.Leave
        Try
            If Me.ObjCD.CekBilanganPositif(Me.TxtChart_DataSelection.Text) = False Then
                MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                Me.TxtChart_DataSelection.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtChart_DataSelection_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtChart_DataSelection.LostFocus
        TxtChart_DataSelection.BackColor = Color.White
    End Sub

    Private Sub ChkShowOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkShowOther.CheckedChanged

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub CboFilterDate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboFilterDate.SelectedIndexChanged
        Try
            If Me.CboFilterDate.SelectedIndex = 0 Then
                Me.dtpFrom.CustomFormat = "dd MMM yyyy"
                Me.dtpTo.CustomFormat = "dd MMM yyyy"
            Else
                Me.dtpFrom.CustomFormat = "MMM yyyy"
                Me.dtpTo.CustomFormat = "MMM yyyy"
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class