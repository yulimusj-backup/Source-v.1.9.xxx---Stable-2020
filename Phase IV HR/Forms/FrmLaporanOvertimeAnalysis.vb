''by    : rosadin
''date  : 2014-08-19

Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms
Public Class FrmLaporanOvertimeAnalysis

    Dim objCd As New ClsCekData
    Dim dtaDept As New DseLaporanOvertimeAnalysisTableAdapters.tblmaster_departemenTableAdapter
    Dim tblDept As New DseLaporanOvertimeAnalysis.tblmaster_departemenDataTable
    Dim rowDept As DseLaporanOvertimeAnalysis.tblmaster_departemenRow
    Dim filter As String

    Private Enum enumOT As Integer
        noOt
        weekdayOt
        holidayOt
    End Enum

    Private Sub FrmLaporanOvertimeAnalysis_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimumSize = New Drawing.Point(1000, 670)
            'Me.WindowState = FormWindowState.Maximized
            Me.dtpFrom.Value = Me.objCd.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.dtpTo.Value = Me.objCd.GetFirstDateofTheMonth(Me.dtpTo.Value)
            isiClbDept()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub isiClbDept()
        Dim Index As Integer = 0
        Try
            dtaDept.Fill(tblDept)
            For Each rowDept In tblDept
                Me.clbDept.Items.Add((Index + 1).ToString + ". " + rowDept.Dept_Code)
                Me.clbDept.SetItemChecked(Index, True)
                Index = Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            tblDept.Dispose()
            dtaDept.Dispose()
        End Try
    End Sub

    Private Sub LlbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.clbDept.Items.Count - 1
                Me.clbDept.SetItemChecked(i, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.clbDept.Items.Count - 1
                Me.clbDept.SetItemChecked(i, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim lListDept As New List(Of Double)
        Dim oItemChecked As Object
        Dim bFilterDept As Boolean = False
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.objCd.GetFirstDateofTheMonth(Me.dtpFrom.Value) > Me.objCd.GetLasDateOfTheMonth(Me.dtpTo.Value) Then
                MsgBox("Date From value can't greater than Date To value", MsgBoxStyle.Exclamation, "Exclamation")
                Exit Sub
            End If
            If clbDept.Items.Count > 0 Then
                For Each oItemChecked In clbDept.CheckedItems
                    If clbDept.GetItemCheckState(clbDept.Items.IndexOf(oItemChecked)) = CheckState.Checked Then
                        lListDept.Add(tblDept.Rows(clbDept.Items.IndexOf(oItemChecked)).Item("Primary"))
                    End If
                Next
            End If
            If lListDept.Count <> tblDept.Count Then
                bFilterDept = True
            Else
                bFilterDept = False
            End If
            If lListDept.Count = 0 Then
                MsgBox("There is no Department selected", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                Me.cetak(Me.objCd.GetFirstDateofTheMonth(Me.dtpFrom.Value), Me.objCd.GetLasDateOfTheMonth(Me.dtpTo.Value), bFilterDept, lListDept, Me.crvLaporanAnalysisOvertime)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cetak_old(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vFilterDept As Boolean, ByVal vListDept As List(Of Double), _
            ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim report As New CrpLaporanOvertimeAnalysis
        Dim dsCetak As New DseLaporanOvertimeAnalysis
        Dim dtaSalary As New DseLaporanOvertimeAnalysisTableAdapters.DataTable1TableAdapter
        Dim tblSalary As New DseLaporanOvertimeAnalysis.DataTable1DataTable
        Dim rowSalary As DseLaporanOvertimeAnalysis.DataTable1Row
        Dim dtaAtt As New DseLaporanOvertimeAnalysisTableAdapters.DataTable2TableAdapter
        Dim tblAtt As New DseLaporanOvertimeAnalysis.DataTable2DataTable
        Dim rowAtt As DseLaporanOvertimeAnalysis.DataTable2Row
        Dim tblCetak As DseLaporanOvertimeAnalysis.DataTable3DataTable = dsCetak.DataTable3
        Dim rowCetak As DseLaporanOvertimeAnalysis.DataTable3Row
        Dim filterDept As String = ""
        Dim sListDept As String = ""
        Dim arrayBaris As String()
        Dim arrayTanggal As String()
        Dim tanggal As Date
        Dim amount As Double
        Dim otweekday As Double = 0
        Dim otweekend As Double = 0
        Try

            ''fill salary
            dtaSalary.FillBy_Date(tblSalary, vEndDate, vStartDate)
            dtaAtt.FillBy_Date(tblAtt, vEndDate, vStartDate)

            ''filter department
            If vFilterDept = True Then
                For Each listitem As String In vListDept
                    If sListDept.Length = 0 Then
                        sListDept = listitem
                    Else
                        sListDept += "," & listitem
                    End If
                Next
                filterDept = "[primary departemen] in (" & sListDept & ")"
                tblSalary.DefaultView.RowFilter = filterDept
            End If

            If tblSalary.DefaultView.Count > 0 Then
                ''periksa tiap baris salary
                For i As Integer = 0 To tblSalary.DefaultView.Count - 1
                    rowSalary = tblSalary.DefaultView(i).Row

                    'bikin baris baru, nyimpen kombinasi durasi waktu per dept
                    Dim tgl As Date = Me.dtpFrom.Value
                    For ia As Integer = 0 To (DateDiff(DateInterval.Month, Me.dtpFrom.Value, Me.dtpTo.Value))
                        filter = "[Dept Code] = '" & rowSalary.Dept_Code & "' AND [periode] = #" & _
                                Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia)) & "#"
                        tblCetak.DefaultView.RowFilter = filter
                        If tblCetak.DefaultView.Count <= 0 Then
                            ''jika tidak ada, buat baris baru
                            rowCetak = tblCetak.NewDataTable3Row
                            rowCetak.Dept_Code = rowSalary.Dept_Code
                            rowCetak.Dept_Desc = rowSalary.Description
                            rowCetak.Periode = Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia))
                            rowCetak.OT_Weekday = "00.00"
                            rowCetak.OT_Weekday_Amount = 0
                            rowCetak.OT_Weekend = "00.00"
                            rowCetak.OT_Weekend_Amount = 0
                            rowCetak.OT_Total = "00.00"
                            rowCetak.OT_Total_Amount = 0
                            rowCetak.Bulan_Tahun = Format(rowCetak.Periode, "MMM yy")
                            tblCetak.AddDataTable3Row(rowCetak)
                        End If
                    Next

                    ''pisahkan tiap baris
                    arrayBaris = rowSalary.Explaination.Split(vbNewLine)

                    ''periksa hanya yang OT
                    If arrayBaris(0).Contains("OVERTIME") = True Then

                        ''temukan dan periksa tiap tanggal
                        For i2 As Integer = 0 To arrayBaris.Length - 1
                            otweekday = 0
                            otweekend = 0
                            filter = ""
                            amount = 0
                            arrayTanggal = arrayBaris(i2).Split(vbTab)
                            If IsDate(arrayTanggal(0)) = True Then
                                tanggal = CDate(arrayTanggal(0))
                                amount = CDbl(arrayTanggal(1))

                                ''hanya periksa tanggal yang nyangkut pada form
                                If tanggal >= vStartDate And tanggal <= vEndDate Then

                                    ''get ot duration 
                                    ''hanya periksa yang OT
                                    filter = "[primary karyawan] = " & rowSalary.Primary_Employee.ToString & _
                                            " and [date] = #" & tanggal & "#"
                                    tblAtt.DefaultView.RowFilter = filter
                                    rowAtt = tblAtt.DefaultView(0).Row
                                    If rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.weekdayOt Then
                                        otweekday = rowAtt.ot_duration + rowAtt.early_ot_duration
                                    ElseIf rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.holidayOt Then
                                        otweekend = rowAtt.ot_duration + rowAtt.early_ot_duration
                                    End If

                                    ''periksa dan perbarui tabel cetakan
                                    filter = "[Dept Code] = '" & rowSalary.Dept_Code & "' AND [periode] = #" & Me.objCd.GetFirstDateofTheMonth(tanggal) & "#"
                                    tblCetak.DefaultView.RowFilter = filter

                                    If tblCetak.DefaultView.Count > 0 Then
                                        ''jika ada, update baris yang ada
                                        rowCetak = tblCetak.DefaultView(0).Row
                                        If otweekday > 0 Then
                                            If Not rowCetak.Periode_Unik.Contains(tanggal.ToShortDateString & "-" & rowSalary.Primary_Employee) Then
                                                rowCetak.OT_Weekday = Me.objCd.Convert_ToString(Me.objCd.Convert_ToMinutes(rowCetak.OT_Weekday) + otweekday)
                                                rowCetak.Periode_Unik += tanggal.ToShortDateString & "-" & rowSalary.Primary_Employee & " "
                                            End If
                                            rowCetak.OT_Weekday_Amount += amount
                                        ElseIf otweekend > 0 Then
                                            If Not rowCetak.Periode_Unik.Contains(tanggal.ToShortDateString & "-" & rowSalary.Primary_Employee) Then
                                                rowCetak.OT_Weekend = Me.objCd.Convert_ToString(Me.objCd.Convert_ToMinutes(rowCetak.OT_Weekend) + otweekend)
                                                rowCetak.Periode_Unik += tanggal.ToShortDateString & "-" & rowSalary.Primary_Employee & " "
                                            End If
                                            rowCetak.OT_Weekend_Amount += amount
                                        End If
                                        rowCetak.OT_Total = Me.objCd.Convert_ToString(Me.objCd.Convert_ToMinutes(rowCetak.OT_Weekday) + Me.objCd.Convert_ToMinutes(rowCetak.OT_Weekend))
                                        rowCetak.OT_Total_Amount = rowCetak.OT_Weekday_Amount + rowCetak.OT_Weekend_Amount
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            End If


            ''update laporan
            If tblCetak.Rows.Count > 0 Then
                For Each rowCetak In tblCetak
                    With rowCetak
                        If .OT_Weekday = "00.00" Or .OT_Weekday = "0.0" Or .OT_Weekday = "0" Then
                            .OT_Weekday = "-"
                        End If
                        If .OT_Weekend = "00.00" Or .OT_Weekend = "0.0" Or .OT_Weekend = "0" Then
                            .OT_Weekend = "-"
                        End If
                        If .OT_Total = "00.00" Or .OT_Total = "0.0" Or .OT_Total = "0" Then
                            .OT_Total = "-"
                        End If
                    End With
                Next
                report.SetDataSource(dsCetak)
                report.SetParameterValue("StartMonth", vStartDate)
                report.SetParameterValue("EndMonth", vEndDate)
                vCrv.ReportSource = report
                vCrv.ShowRefreshButton = False
                vCrv.Show()
                vCrv.Refresh()
            Else
                MsgBox("No data for this period.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub cetak_old2(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vFilterDept As Boolean, ByVal vListDept As List(Of Double), _
            ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim report As New CrpLaporanOvertimeAnalysis
        Dim dsCetak As New DseLaporanOvertimeAnalysis
        Dim DtaParam As New DseLaporanOvertimeAnalysisTableAdapters.DataTable3NewTableAdapter
        Dim TblParam As New DseLaporanOvertimeAnalysis.DataTable3NewDataTable
        Dim rowParam As DseLaporanOvertimeAnalysis.DataTable3NewRow
        Dim dtaSalary As New DseLaporanOvertimeAnalysisTableAdapters.DataTable1NewTableAdapter
        Dim tblSalary As New DseLaporanOvertimeAnalysis.DataTable1NewDataTable
        Dim rowSalary As DseLaporanOvertimeAnalysis.DataTable1NewRow
        Dim dtaAtt As New DseLaporanOvertimeAnalysisTableAdapters.DataTable2NewTableAdapter
        Dim tblAtt As New DseLaporanOvertimeAnalysis.DataTable2NewDataTable
        Dim rowAtt As DseLaporanOvertimeAnalysis.DataTable2NewRow
        Dim tblCetak As DseLaporanOvertimeAnalysis.DataTable3DataTable = dsCetak.DataTable3
        Dim rowCetak As DseLaporanOvertimeAnalysis.DataTable3Row
        Dim filterDept As String = ""
        Dim sListDept As String = ""
        Dim arrayBaris As String()
        Dim otweekday, otweekend As Double
        Dim otweekday_amount, otweekend_amount As Decimal
        Dim tgl As Date = Me.dtpFrom.Value
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DecPrecIDR As Integer = ObjCur.GetDecimal_DefaultCurrency()

        Try
            dtaSalary.Connection.ConnectionString = Core_Module.GetConstring
            dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            dtaDept.Connection.ConnectionString = Core_Module.GetConstring
            DtaParam.Connection.ConnectionString = Core_Module.GetConstring

            ''fill salary
            dtaSalary.FillBy_Date(tblSalary, vEndDate, vStartDate)
            dtaAtt.FillBy_Date(tblAtt, vEndDate, vStartDate)
            DtaParam.Fill(TblParam)

            For Each listitem As String In vListDept
                If sListDept.Length = 0 Then
                    sListDept = listitem
                Else
                    sListDept += "," & listitem
                End If

                'bikin baris baru, nyimpen kombinasi durasi waktu per dept
                For ia As Integer = 0 To (DateDiff(DateInterval.Month, Me.dtpFrom.Value, Me.dtpTo.Value))
                    filter = "[primary departmen] = '" & listitem & "' AND [periode] = #" & _
                            Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia)) & "#"
                    tblCetak.DefaultView.RowFilter = filter
                    If tblCetak.DefaultView.Count <= 0 Then
                        ''jika tidak ada, buat baris baru
                        rowCetak = tblCetak.NewDataTable3Row
                        rowCetak.primary_departmen = listitem
                        tblDept.DefaultView.RowFilter = Nothing
                        tblDept.DefaultView.RowFilter = "[Primary] = " & rowCetak.primary_departmen
                        If tblDept.DefaultView.Count > 0 Then
                            rowDept = tblDept.DefaultView(0).Row
                            rowCetak.Dept_Code = rowDept.Dept_Code
                            rowCetak.Dept_Desc = rowDept.Dept_Desc
                        End If
                        rowCetak.Periode = Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia))
                        rowCetak.OT_Weekday = "00.00"
                        rowCetak.OT_Weekday_Amount = 0
                        rowCetak.OT_Weekend = "00.00"
                        rowCetak.OT_Weekend_Amount = 0
                        rowCetak.OT_Total = "00.00"
                        rowCetak.OT_Total_Amount = 0
                        rowCetak.Bulan_Tahun = Format(rowCetak.Periode, "MMM yy")
                        tblCetak.AddDataTable3Row(rowCetak)
                    End If
                Next
            Next

            ''filter department
            If vFilterDept = True Then
                filterDept = "[primary departemen] in (" & sListDept & ")"
                tblSalary.DefaultView.RowFilter = filterDept
                tblAtt.DefaultView.RowFilter = filterDept
            End If

            For Each rowCetak In tblCetak.Rows
                otweekday = 0
                otweekend = 0
                otweekday_amount = 0
                otweekend_amount = 0

                tblSalary.DefaultView.RowFilter = Nothing
                tblSalary.DefaultView.RowFilter = "[primary departemen] = " & rowCetak.primary_departmen.ToString & _
                " AND [Start Date] <= #" & rowCetak.Periode & "# AND [End Date] >= #" & rowCetak.Periode & "#"
                ''periksa tiap baris salary
                For i As Integer = 0 To tblSalary.DefaultView.Count - 1
                    rowSalary = tblSalary.DefaultView(i).Row
                    tblAtt.DefaultView.RowFilter = "[primary departemen] = " & rowSalary.primary_departemen.ToString & _
                    " AND [Date] <= #" & rowSalary.End_Date & "# AND [Date] >= #" & rowSalary.Start_Date & "#"

                    ''pisahkan tiap baris
                    arrayBaris = rowSalary.Explaination.Split(vbNewLine)
                    ''periksa hanya yang OT
                    If arrayBaris(0).Contains("OVERTIME") = True Then
                        If i = 0 Then
                            For Each dv As DataRowView In tblAtt.DefaultView
                                rowAtt = dv.Row
                                If rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.weekdayOt Then
                                    otweekday += rowAtt.ot_duration + rowAtt.early_ot_duration
                                ElseIf rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.holidayOt Then
                                    otweekend += rowAtt.ot_duration + rowAtt.early_ot_duration
                                End If
                            Next
                        End If

                        TblParam.DefaultView.RowFilter = Nothing
                        TblParam.DefaultView.RowFilter = "[Identifier Salary Component] = '" & rowSalary.Identifier_Salary_Component.ToString & "'"
                        If TblParam.DefaultView.Count > 0 Then
                            rowParam = TblParam.DefaultView(0).Row
                            'cari di parameter bahwa salary component overtime otweekend atau otholiday
                            rowParam.Code = rowParam.Code
                            If rowParam.Primary_Parameter = 32 Then
                                otweekday_amount += rowSalary.Amount
                            ElseIf rowParam.Primary_Parameter = 34 Then
                                otweekend_amount += rowSalary.Amount
                            End If
                        End If
                    End If
                Next

                rowCetak.OT_Weekday = Me.objCd.Convert_ToString(otweekday)
                rowCetak.OT_Weekend = Me.objCd.Convert_ToString(otweekend)
                rowCetak.OT_Total = Me.objCd.Convert_ToString(otweekday + otweekend)
                rowCetak.OT_Weekday_Amount = otweekday_amount
                rowCetak.OT_Weekend_Amount = otweekend_amount
                rowCetak.OT_Total_Amount = otweekday_amount + otweekend_amount
            Next

            ''update laporan
            If tblCetak.Rows.Count > 0 Then
                For Each rowCetak In tblCetak
                    With rowCetak
                        If .OT_Weekday = "00.00" Or .OT_Weekday = "0.0" Or .OT_Weekday = "0" Then
                            .OT_Weekday = "-"
                        End If
                        If .OT_Weekend = "00.00" Or .OT_Weekend = "0.0" Or .OT_Weekend = "0" Then
                            .OT_Weekend = "-"
                        End If
                        If .OT_Total = "00.00" Or .OT_Total = "0.0" Or .OT_Total = "0" Then
                            .OT_Total = "-"
                        End If
                    End With
                Next
                report.SetDataSource(dsCetak)
                report.SetParameterValue("StartMonth", vStartDate)
                report.SetParameterValue("EndMonth", vEndDate)
                report.SetParameterValue("DecPrec", DecPrecIDR)
                vCrv.ReportSource = report
                vCrv.ShowRefreshButton = False
                vCrv.Show()
                vCrv.Refresh()
            Else
                MsgBox("No data for this period.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dtaSalary.Dispose()
            dtaAtt.Dispose()
            DtaParam.Dispose()
            dtaDept.Dispose()
            tblSalary.Dispose()
            tblAtt.Dispose()
            TblParam.Dispose()
            tblDept.Dispose()
            tblCetak.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub dtpFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpFrom.KeyDown
        Try
            If e.KeyValue = Keys.Enter Then
                Me.btnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub dtpTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpTo.KeyDown
        Try
            If e.KeyValue = Keys.Enter Then
                Me.btnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub cetak(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vFilterDept As Boolean, ByVal vListDept As List(Of Double), _
            ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim report As New CrpLaporanOvertimeAnalysis
        Dim dsCetak As New DseLaporanOvertimeAnalysis
        Dim DtaParam As New DseLaporanOvertimeAnalysisTableAdapters.DataTable3NewTableAdapter
        Dim TblParam As New DseLaporanOvertimeAnalysis.DataTable3NewDataTable
        Dim dtaSalary As New DseLaporanOvertimeAnalysisTableAdapters.DataTable1NewTableAdapter
        Dim tblSalary As New DseLaporanOvertimeAnalysis.DataTable1NewDataTable
        Dim rowSalary As DseLaporanOvertimeAnalysis.DataTable1NewRow
        Dim dtaAtt As New DseLaporanOvertimeAnalysisTableAdapters.DataTable2NewTableAdapter
        Dim tblAtt As New DseLaporanOvertimeAnalysis.DataTable2NewDataTable
        Dim rowAtt As DseLaporanOvertimeAnalysis.DataTable2NewRow
        Dim tblCetak As DseLaporanOvertimeAnalysis.DataTable3DataTable = dsCetak.DataTable3
        Dim rowCetak As DseLaporanOvertimeAnalysis.DataTable3Row
        Dim filterDept As String = ""
        Dim sListDept As String = ""
        Dim arrayBaris As String()
        Dim otweekday, otweekend As Double
        Dim otweekday_amount, otweekend_amount As Decimal
        Dim tgl As Date = Me.dtpFrom.Value
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DecPrecIDR As Integer = ObjCur.GetDecimal_DefaultCurrency()

        Try
            dtaSalary.Connection.ConnectionString = Core_Module.GetConstring
            dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            dtaDept.Connection.ConnectionString = Core_Module.GetConstring
            DtaParam.Connection.ConnectionString = Core_Module.GetConstring

            ''fill salary
            dtaSalary.FillBy_Date(tblSalary, vEndDate, vStartDate)
            dtaAtt.FillBy_Date(tblAtt, vEndDate, vStartDate)
            DtaParam.Fill(TblParam)

            For Each listitem As String In vListDept
                If sListDept.Length = 0 Then
                    sListDept = listitem
                Else
                    sListDept += "," & listitem
                End If

                'bikin baris baru, nyimpen kombinasi durasi waktu per dept
                For ia As Integer = 0 To (DateDiff(DateInterval.Month, Me.dtpFrom.Value, Me.dtpTo.Value))
                    filter = "[primary departmen] = '" & listitem & "' AND [periode] = #" & _
                            Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia)) & "#"
                    tblCetak.DefaultView.RowFilter = filter
                    If tblCetak.DefaultView.Count <= 0 Then
                        ''jika tidak ada, buat baris baru
                        rowCetak = tblCetak.NewDataTable3Row
                        rowCetak.primary_departmen = listitem
                        tblDept.DefaultView.RowFilter = Nothing
                        tblDept.DefaultView.RowFilter = "[Primary] = " & rowCetak.primary_departmen
                        If tblDept.DefaultView.Count > 0 Then
                            rowDept = tblDept.DefaultView(0).Row
                            rowCetak.Dept_Code = rowDept.Dept_Code
                            rowCetak.Dept_Desc = rowDept.Dept_Desc
                        End If
                        rowCetak.Periode = Me.objCd.GetFirstDateofTheMonth(tgl.AddMonths(ia))
                        rowCetak.OT_Weekday = "00.00"
                        rowCetak.OT_Weekday_Amount = 0
                        rowCetak.OT_Weekend = "00.00"
                        rowCetak.OT_Weekend_Amount = 0
                        rowCetak.OT_Total = "00.00"
                        rowCetak.OT_Total_Amount = 0
                        rowCetak.Bulan_Tahun = Format(rowCetak.Periode, "MMM yy")
                        tblCetak.AddDataTable3Row(rowCetak)
                    End If
                Next
            Next

            For Each rowCetak In tblCetak.Rows
                otweekday = 0
                otweekend = 0
                otweekday_amount = 0
                otweekend_amount = 0

                tblAtt.DefaultView.RowFilter = Nothing
                tblAtt.DefaultView.RowFilter = "[primary departemen] = " & rowCetak.primary_departmen.ToString & _
                " AND [Date] >= #" & rowCetak.Periode & "# AND [Date] <= #" & _
                Me.objCd.GetLasDateOfTheMonth(rowCetak.Periode) & "#"
                If tblAtt.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In tblAtt.DefaultView
                        rowAtt = dv.Row
                        If rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.weekdayOt Then
                            otweekday += rowAtt.ot_duration + rowAtt.early_ot_duration
                        ElseIf rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.holidayOt Then
                            otweekend += rowAtt.ot_duration + rowAtt.early_ot_duration
                        End If

                        tblSalary.DefaultView.RowFilter = "[primary departemen] = " & rowCetak.primary_departmen.ToString & _
                        " AND [Primary Employee] = " & rowAtt.primary_karyawan.ToString & " AND [Start Date] <= #" & _
                        rowAtt._date & "# AND [End Date] >= #" & rowAtt._date & "#"
                        If tblSalary.DefaultView.Count > 0 Then
                            For Each dv1 As DataRowView In tblSalary.DefaultView
                                rowSalary = dv1.Row
                                arrayBaris = rowSalary.Explaination.Split(vbNewLine)
                                If arrayBaris(0).Contains("OVERTIME") = True Then
                                    Dim indexStart As Integer = -1
                                    Dim strBetul As String = ""
                                    Dim arrBetul As String()
                                    Dim amountBetul As Decimal = 0

                                    indexStart = rowSalary.Explaination.IndexOf(Format(rowAtt._date, "dd-MMM-yyyy")) + 12
                                    strBetul = rowSalary.Explaination.Substring(indexStart, 10)
                                    arrBetul = strBetul.Split(vbNewLine)
                                    amountBetul = arrBetul(0)
                                    If amountBetul > 0 Then
                                        If rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.weekdayOt Then
                                            otweekday_amount += amountBetul
                                        ElseIf rowAtt.tipe_overtime = FrmLaporanOvertimeAnalysis.enumOT.holidayOt Then
                                            otweekend_amount += amountBetul
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If

                rowCetak.OT_Weekday = Me.objCd.Convert_ToString(otweekday)
                rowCetak.OT_Weekend = Me.objCd.Convert_ToString(otweekend)
                rowCetak.OT_Total = Me.objCd.Convert_ToString(otweekday + otweekend)
                rowCetak.OT_Weekday_Amount = otweekday_amount
                rowCetak.OT_Weekend_Amount = otweekend_amount
                rowCetak.OT_Total_Amount = otweekday_amount + otweekend_amount
            Next

            ''update laporan
            If tblCetak.Rows.Count > 0 Then
                For Each rowCetak In tblCetak
                    With rowCetak
                        If .OT_Weekday = "00.00" Or .OT_Weekday = "0.0" Or .OT_Weekday = "0" Then
                            .OT_Weekday = "-"
                        End If
                        If .OT_Weekend = "00.00" Or .OT_Weekend = "0.0" Or .OT_Weekend = "0" Then
                            .OT_Weekend = "-"
                        End If
                        If .OT_Total = "00.00" Or .OT_Total = "0.0" Or .OT_Total = "0" Then
                            .OT_Total = "-"
                        End If
                    End With
                Next
                report.SetDataSource(dsCetak)
                report.SetParameterValue("StartMonth", vStartDate)
                report.SetParameterValue("EndMonth", vEndDate)
                report.SetParameterValue("DecPrec", DecPrecIDR)
                vCrv.ReportSource = report
                vCrv.ShowRefreshButton = False
                vCrv.Show()
                vCrv.Refresh()
            Else
                MsgBox("No data for this period.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dtaSalary.Dispose()
            dtaAtt.Dispose()
            DtaParam.Dispose()
            dtaDept.Dispose()
            tblSalary.Dispose()
            tblAtt.Dispose()
            TblParam.Dispose()
            tblDept.Dispose()
            tblCetak.Dispose()
            ObjCur = Nothing
        End Try
    End Sub
End Class