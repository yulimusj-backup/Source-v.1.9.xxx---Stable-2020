Option Strict Off
Option Explicit On

Public Class ClsMasterParameterGaji
    Dim ObjCD As New Core_Module.ClsCekData
    Dim dtaKompetensi As New DseMasterParameterGajiTableAdapters.tblmaster_kompetensiTableAdapter
    Dim tblKompetensi As New DseMasterParameterGaji.tblmaster_kompetensiDataTable
    Dim rowKompetensi As DseMasterParameterGaji.tblmaster_kompetensiRow

    Public Sub New()
        Try
            dtaKompetensi.Connection.ConnectionString = Core_Module.GetConstring
            dtaKompetensi.Fill(tblKompetensi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumComparer As Integer
        SamaDengan
        LebihBesar
        LebihKecil
        LebihBesar_SamaDengan
        LebihKecil_SamaDengan
        TidakSamaDengan
    End Enum

    Public Enum EnumTipeOvertime As Integer
        NoOvertime
        WorkdayOvertime
        HolidayOvertime
    End Enum

    Public Enum EnumTipeWorkdays As Integer
        Workday
        Holiday
    End Enum

    Public Sub FillDataOnly(ByVal vTBLParam1 As DseMasterParameter.DataTable1DataTable, _
    ByVal vTBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, _
    ByVal vTglDari As Date, ByVal vTglSampai As Date, ByVal vPrimaryKaryawan As Double)
        Dim DTAParam1 As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim DTAParam2 As New DseMasterParameterTableAdapters.tbltrans_attendanceTableAdapter
        Dim DTAParam As New DseMasterParameterTableAdapters.tbltrans_attendance1TableAdapter
        Dim DTAKaryawan As New DseMasterParameterTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTAParam1.Connection.ConnectionString = Core_Module.GetConstring
            DTAParam2.Connection.ConnectionString = Core_Module.GetConstring
            DTAParam.Connection.ConnectionString = Core_Module.GetConstring
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring

            DTAParam1.FillBy_PrimaryKaryawan1(vTBLParam1, vTglSampai, vTglDari, vPrimaryKaryawan)
            DTAParam2.FillBy_PrimaryKaryawan(vTBLParam2, vTglSampai, vTglDari, vPrimaryKaryawan)
            DTAParam.FillBy_PrimaryKaryawan(TBLWeekDay, vTglSampai, vTglDari, EnumTipeOvertime.WorkdayOvertime, vPrimaryKaryawan)
            DTAParam.FillBy_PrimaryKaryawan(TBLHoliday, vTglSampai, vTglDari, EnumTipeOvertime.HolidayOvertime, vPrimaryKaryawan)
            DTAKaryawan.FillBy_PrimaryKary(TblKaryawan, vPrimaryKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAParam1.Dispose()
            DTAParam2.Dispose()
            DTAParam.Dispose()
            DTAKaryawan.Dispose()
        End Try
    End Sub

    Public Sub FillDataOnly(ByVal vTBLParam1 As DseMasterParameter.DataTable1DataTable, _
  ByVal vTBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, _
  ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, _
  ByVal vTglDari As Date, ByVal vTglSampai As Date)
        Dim DTAParam1 As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim DTAParam2 As New DseMasterParameterTableAdapters.tbltrans_attendanceTableAdapter
        Dim DTAParam As New DseMasterParameterTableAdapters.tbltrans_attendance1TableAdapter
        Dim DTAKaryawan As New DseMasterParameterTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTAParam1.Connection.ConnectionString = Core_Module.GetConstring
            DTAParam2.Connection.ConnectionString = Core_Module.GetConstring
            DTAParam.Connection.ConnectionString = Core_Module.GetConstring
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring

            DTAParam1.FillBy_AllData(vTBLParam1, vTglSampai, vTglDari)
            DTAParam2.Fill(vTBLParam2, vTglSampai, vTglDari)
            DTAParam.Fill(TBLWeekDay, vTglSampai, vTglDari, EnumTipeOvertime.WorkdayOvertime)
            DTAParam.Fill(TBLHoliday, vTglSampai, vTglDari, EnumTipeOvertime.HolidayOvertime)
            DTAKaryawan.Fill(TblKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAParam1.Dispose()
            DTAParam2.Dispose()
            DTAParam.Dispose()
            DTAKaryawan.Dispose()
        End Try
    End Sub

    Public Function GetValue_Daily(ByVal TBL1 As DseMasterParameter.DataTable1DataTable, _
    ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, _
    ByVal vPrimaryEmployee As Double, ByVal vPrimaryParameter As Double, ByVal vTglDari As Date, _
    ByVal vTipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable, ByVal vObjKaryawan As ClsMasterKaryawan, _
    Optional ByVal vPrimaryLeave As Double = -1) As String
        Dim Row1 As DseMasterParameter.DataTable1Row
        Dim RowKaryawan As DseMasterParameter.tblmaster_karyawanRow
        Dim OTType As Integer
        Dim WorkdayType As Integer
        Dim Filter As String = ""
        Dim Dummy As String = ""
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            If vTipeParameter = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                Dummy = vObjKaryawan.GetKompetensi(vPrimaryEmployee, vPrimaryParameter)
                If Dummy = "" Then
                    Me.tblKompetensi.DefaultView.RowFilter = "[Primary] = " & vPrimaryParameter.ToString
                    If Me.tblKompetensi.DefaultView.Count > 0 Then
                        Me.rowKompetensi = Me.tblKompetensi.DefaultView(0).Row
                        Return Me.rowKompetensi.Value
                    Else
                        Dim ObjKomp As New Core_Module.ClsMasterKompetensi
                        Dim TipeKomp As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
                        TipeKomp = ObjKomp.GetDataType(vPrimaryParameter)
                        If TipeKomp = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            Return CInt(ClsMasterKompetensi.Aphabetic.F)
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Number Then
                            Return 0
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            Return #1/1/1950#
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Text Then
                            Return ""
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.YN Then
                            Return 0
                        End If
                    End If
                Else
                    Return Dummy
                End If
            ElseIf vTipeParameter = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                If vPrimaryParameter = 1 Then
                    Return vTglDari.DayOfWeek
                ElseIf vPrimaryParameter = 2 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 3 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 4 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 5 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 6 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.late_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 7 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_in_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 8 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_out_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 9 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.ot_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 10 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_ot_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 11 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_late
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 12 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 13 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 14 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_overtime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 15 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_overtime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 16 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_work_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 17 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_work_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 18 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_present
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 19 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_sick
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 20 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_cuti
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 21 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.Is_alfa
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 22 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_permit
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 23 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 24 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 25 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 26 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 27 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 28 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.late_break_in_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 31 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_break_out_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 32 Then
                    OTType = ClsMasterParameterGaji.EnumTipeOvertime.WorkdayOvertime
                    Filter = "[tipe overtime] = " + OTType.ToString + " AND [primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.tipe_overtime = OTType Then
                            Return 1
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 34 Then
                    OTType = ClsMasterParameterGaji.EnumTipeOvertime.HolidayOvertime
                    Filter = "[tipe overtime] = " + OTType.ToString + " AND [primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.tipe_overtime = OTType Then
                            Return 1
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 35 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    WorkdayType = ClsMasterParameterGaji.EnumTipeWorkdays.Workday

                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.Tipe_Workdays = WorkdayType Then
                            Return 1
                        Else
                            Return 0
                        End If
                        'Return Row1.is_present
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 37 Then
                    WorkdayType = ClsMasterParameterGaji.EnumTipeWorkdays.Holiday
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.Tipe_Workdays = WorkdayType Then
                            Return 1
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 109 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 112 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.Permit_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 73 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.sex
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 74 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.Basic_Salary
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 75 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.marital_status
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 115 Then
                    Return Me.GetJmlAnak(vPrimaryEmployee)
                ElseIf vPrimaryParameter = 116 Then
                    Return Me.GetTotalOTDuration(vTglDari, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 117 Then
                    Return Me.GetIsOvertime(vTglDari, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 130 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.Employee_Status
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 131 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.primary_position
                    Else
                        Return -1
                    End If
                ElseIf vPrimaryParameter = 133 Then
                    Return Me.GetLeaveCode(vPrimaryEmployee, vPrimaryLeave)
                ElseIf vPrimaryParameter = 134 Then
                    'total masa jabatan - hari
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        If RowKaryawan.admision_date.Date <= vTglDari.Date Then
                            Return (vTglDari.Date - RowKaryawan.admision_date.Date).Days
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 135 Then
                    'total masa jabatan - bulan
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.admision_date.Date
                        If DummyRefDate.Date <= vTglDari.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) < vTglDari.Date
                                DummyRefDate = DummyRefDate.AddMonths(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                            If DummyRefDate.Date <= vTglDari.Date Then
                                DummyRetval += Math.Round((vTglDari.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 136 Then
                    'total masa jabatan - tahun
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.admision_date.Date
                        If DummyRefDate.Date <= vTglDari.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) < vTglDari.Date
                                DummyRefDate = DummyRefDate.AddYears(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                            'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                            If DummyRefDate.Date <= vTglDari.Date Then
                                DummyRetval += Math.Round((vTglDari.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 137 Then
                    'total masa jabatan - hari
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        If RowKaryawan.appointment_date.Date <= vTglDari.Date Then
                            Return (vTglDari.Date - RowKaryawan.appointment_date.Date).Days
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 138 Then
                    'total masa jabatan - bulan
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.appointment_date.Date
                        If DummyRefDate.Date <= vTglDari.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) < vTglDari.Date
                                DummyRefDate = DummyRefDate.AddMonths(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                            If DummyRefDate.Date <= vTglDari.Date Then
                                DummyRetval += Math.Round((vTglDari.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 139 Then
                    'total masa jabatan - tahun
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.appointment_date.Date
                        If DummyRefDate.Date <= vTglDari.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) < vTglDari.Date
                                DummyRefDate = DummyRefDate.AddYears(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                            'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                            If DummyRefDate.Date <= vTglDari.Date Then
                                DummyRetval += Math.Round((vTglDari.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 140 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_on_holiday_list
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 141 Then
                    ' aul -penambahan outstation-
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_outstation
                    Else
                        Return 0
                    End If
                Else
                    Return ""
                End If
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetValue(ByVal TBL1 As DseMasterParameter.DataTable1DataTable, _
    ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, _
    ByVal vPrimaryEmployee As Double, ByVal vPrimaryParameter As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal vTipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable, ByVal vObjKaryawan As ClsMasterKaryawan, _
    Optional ByVal vPrimaryLeave As Double = -1) As String
        Dim Row1 As DseMasterParameter.DataTable1Row
        Dim RowParam2 As DseMasterParameter.tbltrans_attendanceRow
        Dim RowWeekday As DseMasterParameter.tbltrans_attendance1Row
        Dim RowHoliday As DseMasterParameter.tbltrans_attendance1Row
        Dim RowKaryawan As DseMasterParameter.tblmaster_karyawanRow
        Dim ObjCuti As New ClsMasterCuti
        Dim OTType As Integer
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView
        Dim Dummy As String = ""
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            If vTipeParameter = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                Dummy = vObjKaryawan.GetKompetensi(vPrimaryEmployee, vPrimaryParameter)
                If Dummy = "" Then
                    Me.tblKompetensi.DefaultView.RowFilter = "[Primary] = " & vPrimaryParameter.ToString
                    If Me.tblKompetensi.DefaultView.Count > 0 Then
                        Me.rowKompetensi = Me.tblKompetensi.DefaultView(0).Row
                        Return Me.rowKompetensi.Value
                    Else
                        Dim ObjKomp As New Core_Module.ClsMasterKompetensi
                        Dim TipeKomp As Core_Module.ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
                        TipeKomp = ObjKomp.GetDataType(vPrimaryParameter)
                        If TipeKomp = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                            Return CInt(ClsMasterKompetensi.Aphabetic.F)
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Number Then
                            Return 0
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Tanggal Then
                            Return #1/1/1950#
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.Text Then
                            Return ""
                        ElseIf TipeKomp = ClsMasterKompetensi.EnumDataType.YN Then
                            Return 0
                        End If
                    End If
                Else
                    Return Dummy
                End If
            Else
                If vPrimaryParameter = 1 Then
                    Return vTglDari.DayOfWeek
                ElseIf vPrimaryParameter = 2 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 3 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 4 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 5 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 6 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.late_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 7 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_in_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 8 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_out_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 9 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.ot_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 10 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_ot_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 11 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_late
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 12 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 13 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 14 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_overtime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 15 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_early_overtime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 16 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_work_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 17 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_work_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 18 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_present
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 19 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_sick
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 20 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_cuti
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 21 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.Is_alfa
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 22 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_permit
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 23 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 24 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 25 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_out
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 26 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_in
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 27 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.actual_break_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 28 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.late_break_in_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 31 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.early_break_out_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 32 Then
                    OTType = ClsMasterParameterGaji.EnumTipeOvertime.WorkdayOvertime
                    Filter = "[tipe overtime] = " + OTType.ToString + " AND [primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        For Each RowView In TBL1.DefaultView
                            Row1 = RowView.Row
                            If Row1.tipe_overtime = OTType Then
                                Return 1
                            End If
                        Next
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 34 Then
                    OTType = ClsMasterParameterGaji.EnumTipeOvertime.HolidayOvertime
                    Filter = "[tipe overtime] = " + OTType.ToString + " AND [primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        For Each RowView In TBL1.DefaultView
                            Row1 = RowView.Row
                            If Row1.tipe_overtime = OTType Then
                                Return 1
                            End If
                        Next
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 35 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    Dim WorkdayType As ClsMasterParameterGaji.EnumTipeWorkdays = ClsMasterParameterGaji.EnumTipeWorkdays.Workday

                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.Tipe_Workdays = workdaytype Then
                            Return 1
                        Else
                            Return 0
                        End If
                        'Return Row1.is_present
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 37 Then
                    Dim WorkdayType As ClsMasterParameterGaji.EnumTipeWorkdays = ClsMasterParameterGaji.EnumTipeWorkdays.Holiday
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        If Row1.Tipe_Workdays = WorkdayType Then
                            Return 1
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 40 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 41 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyInDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 42 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOutDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 43 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalOvertimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 44 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOvertimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 45 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 46 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyInDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 47 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOutDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 48 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalOvertimeDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 49 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOvertimeDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 50 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalActualWorkDaysDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 51 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPresent
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 52 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalSick
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 53 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLeave
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 54 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalAlpa
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 55 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPermit
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 56 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalActualBreakTime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 57 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateBreakInDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 60 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyBreakOutDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 61 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateBreakInDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 64 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyBreakOutDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 65 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLWeekDay.DefaultView.RowFilter = Filter
                    If TBLWeekDay.DefaultView.Count > 0 Then
                        RowWeekday = TBLWeekDay.DefaultView(0).Row
                        Return RowWeekday.TotalOvertimeWeekHoliday
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 67 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLHoliday.DefaultView.RowFilter = Filter
                    If TBLHoliday.DefaultView.Count > 0 Then
                        RowHoliday = TBLHoliday.DefaultView(0).Row
                        Return RowHoliday.TotalOvertimeWeekHoliday
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 68 Then
                    'TotalWeekdaysMonthly
                    Return Me.GetTotalWeekdays(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 70 Then
                    'TotalHolidaysMonthly
                    Return Me.GetTotalHoliday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 73 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.sex
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 74 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.Basic_Salary
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 75 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.marital_status
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 76 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 77 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyInDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 78 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOutDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 79 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalOvertimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 80 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOvertimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 81 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 82 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyInDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 83 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOutDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 84 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalOvertimeDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 85 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyOvertimeDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 86 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalActualWorkDaysDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 87 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPresent
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 88 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalSick
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 89 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLeave
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 90 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalAlpa
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 91 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPermit
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 92 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalActualBreakTime
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 93 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateBreakInDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 96 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyBreakOutDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 97 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalLateBreakInDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 100 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalEarlyBreakOutDays
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 101 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLWeekDay.DefaultView.RowFilter = Filter
                    If TBLWeekDay.DefaultView.Count > 0 Then
                        RowWeekday = TBLWeekDay.DefaultView(0).Row
                        Return RowWeekday.TotalOvertimeWeekHoliday
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 103 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLHoliday.DefaultView.RowFilter = Filter
                    If TBLHoliday.DefaultView.Count > 0 Then
                        RowHoliday = TBLHoliday.DefaultView(0).Row
                        Return RowHoliday.TotalOvertimeWeekHoliday
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 104 Then
                    'TotalWeekdaysYearly
                    Return Me.GetTotalWeekdays(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 106 Then
                    'TotalHolidaysYearly
                    Return Me.GetTotalHoliday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 109 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.sch_break_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 110 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalScheduleBreakTimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 111 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalScheduleBreakTimeDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 112 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.Permit_duration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 113 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPermitDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 114 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TBLParam2.DefaultView.RowFilter = Filter
                    If TBLParam2.DefaultView.Count > 0 Then
                        RowParam2 = TBLParam2.DefaultView(0).Row
                        Return RowParam2.TotalPermitDuration
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 115 Then
                    Return Me.GetJmlAnak(vPrimaryEmployee)
                ElseIf vPrimaryParameter = 119 Then
                    Return Me.GetJmlTotalHariOvertime(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 121 Then
                    Return Me.GetJmlTotalHariOvertime(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 118 Then
                    Return Me.GetTotalOTDuration(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 120 Then
                    Return Me.GetTotalOTDuration(vTglSampai, vTglDari, vPrimaryEmployee)
                    'Total overtime day weekday, dan holiday (monthly, yearly)
                ElseIf vPrimaryParameter = 126 Then
                    Return Me.GetIsOvertime_Weekday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 127 Then
                    Return Me.GetIsOvertime_Holiday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 128 Then
                    Return Me.GetIsOvertime_Weekday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 129 Then
                    Return Me.GetIsOvertime_Holiday(vTglSampai, vTglDari, vPrimaryEmployee)
                    'Total OT duration weekday, dan holiday (monthly, yearly)
                ElseIf vPrimaryParameter = 122 Then
                    Return Me.GetTotalOTDuration_Weekdays(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 123 Then
                    Return Me.GetTotalOTDuration_Holiday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 124 Then
                    Return Me.GetTotalOTDuration_Weekdays(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 125 Then
                    Return Me.GetTotalOTDuration_Holiday(vTglSampai, vTglDari, vPrimaryEmployee)
                ElseIf vPrimaryParameter = 130 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.Employee_Status
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 131 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        Return RowKaryawan.primary_position
                    Else
                        Return -1
                    End If
                ElseIf vPrimaryParameter = 132 Then
                    Return ObjCuti.GetBalance(vPrimaryEmployee, vPrimaryLeave).ToString
                ElseIf vPrimaryParameter = 133 Then
                    Return Me.GetLeaveCode(vPrimaryEmployee, vPrimaryLeave)
                ElseIf vPrimaryParameter = 134 Then
                    'total masa jabatan - hari
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        If RowKaryawan.admision_date.Date <= vTglSampai.Date Then
                            Return (vTglSampai.Date - RowKaryawan.admision_date.Date).Days
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 135 Then
                    'total masa jabatan - bulan
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.admision_date.Date
                        If DummyRefDate.Date <= vTglSampai.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) < vTglSampai.Date
                                DummyRefDate = DummyRefDate.AddMonths(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                            If DummyRefDate.Date < vTglSampai.Date Then
                                DummyRetval += Math.Round((vTglSampai.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 136 Then
                    'total masa jabatan - tahun
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.admision_date.Date
                        If DummyRefDate.Date <= vTglSampai.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) < vTglSampai.Date
                                DummyRefDate = DummyRefDate.AddYears(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                            'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                            If DummyRefDate.Date < vTglSampai.Date Then
                                DummyRetval += Math.Round((vTglSampai.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 137 Then
                    'total masa jabatan pegawai tetap- hari
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        If RowKaryawan.appointment_date.Date <= vTglSampai.Date Then
                            Return (vTglSampai.Date - RowKaryawan.appointment_date.Date).Days
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 138 Then
                    'total masa jabatan pegawai tetap- bulan
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.appointment_date.Date
                        If DummyRefDate.Date <= vTglSampai.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) < vTglSampai.Date
                                DummyRefDate = DummyRefDate.AddMonths(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                            If DummyRefDate.Date < vTglSampai.Date Then
                                DummyRetval += Math.Round((vTglSampai.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 139 Then
                    'total masa jabatan pegawai tetap - tahun
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + ""
                    TblKaryawan.DefaultView.RowFilter = Filter
                    If TblKaryawan.DefaultView.Count > 0 Then
                        RowKaryawan = TblKaryawan.DefaultView(0).Row
                        DummyRefDate = RowKaryawan.appointment_date.Date
                        If DummyRefDate.Date <= vTglSampai.Date Then
                            DummyRetval = 0
                            While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) < vTglSampai.Date
                                DummyRefDate = DummyRefDate.AddYears(1)
                                DummyRetval += 1
                            End While
                            'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                            'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                            If DummyRefDate.Date < vTglSampai.Date Then
                                DummyRetval += Math.Round((vTglSampai.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                            End If

                            Return DummyRetval
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                    'aul -penambahan outstation-
                ElseIf vPrimaryParameter = 141 Then
                    Filter = "[primary karyawan] = " + vPrimaryEmployee.ToString + " AND [date] = #" + vTglDari.ToString + "#"
                    TBL1.DefaultView.RowFilter = Filter
                    If TBL1.DefaultView.Count > 0 Then
                        Row1 = TBL1.DefaultView(0).Row
                        Return Row1.is_outstation
                    Else
                        Return 0
                    End If
                ElseIf vPrimaryParameter = 142 Then 'Outstation monthly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Holiday, EnumTipeWorkdays.Workday, EnumType.Monthly)
                ElseIf vPrimaryParameter = 143 Then 'Outstation yearly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Holiday, EnumTipeWorkdays.Workday, EnumType.Yearly)
                ElseIf vPrimaryParameter = 144 Then 'Outstation weekday monthly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Workday, EnumTipeWorkdays.Workday, EnumType.Monthly)
                ElseIf vPrimaryParameter = 145 Then 'Outstation holiday monthly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Holiday, EnumTipeWorkdays.Holiday, EnumType.Monthly)
                ElseIf vPrimaryParameter = 146 Then 'Outstation weekday yearly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Workday, EnumTipeWorkdays.Workday, EnumType.Yearly)
                ElseIf vPrimaryParameter = 147 Then 'Outstation holiday yearly
                    Return Me.GetIsOutstation(vTglSampai, vTglDari, vPrimaryEmployee, EnumTipeWorkdays.Holiday, EnumTipeWorkdays.Holiday, EnumType.Yearly)
                ElseIf vPrimaryParameter = 148 Then 'Is Resign
                    Return Me.GetIsResign(vTglSampai, vTglDari, vPrimaryEmployee)
                Else
                    Return ""
                End If
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCuti = Nothing
        End Try
    End Function

    Public Function GetLeaveCode(ByVal vPrimaryEmployee As Double, ByVal vPrimaryLeave As Double) As String
        Dim DTA As New DseMasterParameterTableAdapters.TblMaster_CutiTableAdapter
        Dim TBL As New DseMasterParameter.TblMaster_CutiDataTable
        Dim Row As DseMasterParameter.TblMaster_CutiRow
        Dim Filter As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryKaryawan(TBL, vPrimaryEmployee)
            Filter = "[primary leave] = " + vPrimaryLeave.ToString + ""
            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count > 0 Then
                Row = TBL.DefaultView(0).Row
                Return Row.primary_leave.ToString
            Else
                Return "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1"
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetTotalOTDuration_Weekdays(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTotalOTDuration_Weekday(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalOTDuration_Holiday(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTotalOTDuration_Holiday(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsOvertime_Weekday(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim Retval As Decimal = 0
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml_Is_OvertimeByDate_Weekday(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
            If Retval > 0 Then
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsOvertime_Holiday(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim Retval As Decimal = 0
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml_Is_OvertimeByDate_Holiday(vPrimaryEmployee, vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
            If Retval > 0 Then
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalOTDuration(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTotalOTDuration(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlTotalHariOvertime(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml_Is_OvertimeByDate(vPrimaryEmployee, vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsOvertime(ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim Retval As Decimal = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml_Is_OvertimeByDate(vPrimaryEmployee, vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
            If Retval > 0 Then
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAnak(ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlAnak(vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalWeekdays(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTotalWeekdays(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTotalHoliday(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTotalHoliday(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ComparerNumeric(ByVal Var1 As String, ByVal Var2 As String, ByVal Comparer As EnumComparer) As Boolean
        Try
            If Comparer = EnumComparer.SamaDengan Then
                If CDec(Var1) = CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihBesar Then
                If CDec(Var1) > CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihKecil Then
                If CDec(Var1) < CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                If CDec(Var1) >= CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                If CDec(Var1) <= CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.TidakSamaDengan Then
                If CDec(Var1) <> CDec(Var2) Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ComparerDate(ByVal Var1 As Date, ByVal Var2 As Date, ByVal Comparer As EnumComparer) As Boolean
        Try
            If Comparer = EnumComparer.SamaDengan Then
                If Var1 = Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihBesar Then
                If Var1 > Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihKecil Then
                If Var1 < Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                If Var1 >= Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                If Var1 <= Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If Comparer = EnumComparer.TidakSamaDengan Then
                If Var1 <> Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function Comparer_Text_Alphabetic_YesNo(ByVal Var1 As String, ByVal Var2 As String, ByVal Comparer As EnumComparer) As Boolean
        Try
            If Comparer = EnumComparer.SamaDengan Then
                If Var1 = Var2 Then
                    Return True
                Else
                    Return False
                End If
            ElseIf Comparer = EnumComparer.TidakSamaDengan Then
                If Var1 <> Var2 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsConstraintBenar(ByVal Var1 As String, ByVal Var2 As String, ByVal Comparer As EnumComparer, _
    ByVal PrimaryParameter As Double, ByVal vTipeParameter2 As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable, _
    ByVal vObjKompetensi As Core_Module.ClsMasterKompetensi) As Boolean
        Dim JenisValue As String = ""
        Dim ObjCD As New ClsCekData

        Try
            If vTipeParameter2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi Then
                Dim TipeKompetensi As ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
                TipeKompetensi = vObjKompetensi.GetDataType(PrimaryParameter)
                If TipeKompetensi = ClsMasterKompetensi.EnumDataType.Alfabetic Or _
                TipeKompetensi = ClsMasterKompetensi.EnumDataType.Text Or _
                TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Then
                    If TipeKompetensi = ClsMasterKompetensi.EnumDataType.YN Then
                        If Var1.ToLower = "yes" Then
                            Var1 = "1"
                        ElseIf Var1.ToLower = "no" Then
                            Var1 = "0"
                        End If
                    ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                        If Var1.ToLower = "a" Then
                            Var1 = "0"
                        ElseIf Var1.ToLower = "b" Then
                            Var1 = "1"
                        ElseIf Var1.ToLower = "c" Then
                            Var1 = "2"
                        ElseIf Var1.ToLower = "d" Then
                            Var1 = "3"
                        ElseIf Var1.ToLower = "e" Then
                            Var1 = "4"
                        ElseIf Var1.ToLower = "f" Then
                            Var1 = "5"
                        End If
                    End If

                    Return Me.Comparer_Text_Alphabetic_YesNo(Var1, Var2, Comparer)
                ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Number Then
                    Return Me.ComparerNumeric(Var1, Var2, Comparer)
                ElseIf TipeKompetensi = ClsMasterKompetensi.EnumDataType.Tanggal Then
                    Return Me.ComparerDate(CDate(Var1), CDate(Var2), Comparer)
                End If
            ElseIf vTipeParameter2 = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode Then
                'Keterangan : Var1 dari database, Var2 dari Master Salary Formula
                JenisValue = Me.GetUnit(PrimaryParameter)
                If JenisValue = "day of week" Then
                    If Var2 = "Sunday" Then
                        Var2 = "0"
                    ElseIf Var2 = "Monday" Then
                        Var2 = "1"
                    ElseIf Var2 = "Tuesday" Then
                        Var2 = "2"
                    ElseIf Var2 = "Wednesday" Then
                        Var2 = "3"
                    ElseIf Var2 = "Thursday" Then
                        Var2 = "4"
                    ElseIf Var2 = "Friday" Then
                        Var2 = "5"
                    ElseIf Var2 = "Saturday" Then
                        Var2 = "6"
                    End If
                    If Comparer = EnumComparer.SamaDengan Then
                        If CInt(Var1) = CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CInt(Var1) > CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CInt(Var1) < CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CInt(Var1) >= CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CInt(Var1) <= CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CInt(Var1) <> CInt(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "Time" Then
                    Var2 = ObjCD.Convert_ToMinutes(Var2).ToString
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CDbl(Var1) > CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CDbl(Var1) < CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CDbl(Var1) >= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CDbl(Var1) <= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "Duration" Then
                    Var2 = ObjCD.Convert_ToMinutes(Var2).ToString
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CDbl(Var1) > CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CDbl(Var1) < CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CDbl(Var1) >= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CDbl(Var1) <= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "0/1/0.5" Then
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CDbl(Var1) > CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CDbl(Var1) < CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CDbl(Var1) >= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CDbl(Var1) <= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "yes/no" Then
                    If Var2.ToLower = "yes" Then
                        Var2 = "1"
                    ElseIf Var2.ToLower = "no" Then
                        Var2 = "0"
                    End If
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "Days" Then
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CDbl(Var1) > CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CDbl(Var1) < CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CDbl(Var1) >= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CDbl(Var1) <= CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue.ToLower = "numeric" Then
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDec(Var1) = CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar Then
                        If CDec(Var1) > CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil Then
                        If CDec(Var1) < CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihBesar_SamaDengan Then
                        If CDec(Var1) >= CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.LebihKecil_SamaDengan Then
                        If CDec(Var1) <= CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDec(Var1) <> CDec(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "male/female" Then
                    If Var2 = "male" Then
                        Var2 = ClsMasterKaryawan.EnumSex.Male
                    ElseIf Var2 = "female" Then
                        Var2 = ClsMasterKaryawan.EnumSex.Female
                    End If
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "Single/Married" Then
                    If Var2.ToLower = "single" Then
                        Var2 = ClsMasterKaryawan.EnumMartialStatus._Single
                    ElseIf Var2.ToLower = "married" Then
                        Var2 = ClsMasterKaryawan.EnumMartialStatus.Married
                    End If
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "probation/contract/jobholder" Then
                    If Trim(Var2).ToLower = "probation" Then
                        Var2 = ClsMasterKaryawan.EnumEmployStatus.Probation
                    ElseIf Trim(Var2).ToLower = "contract" Then
                        Var2 = ClsMasterKaryawan.EnumEmployStatus.Contract
                    ElseIf Trim(Var2).ToLower = "jobholder" Then
                        Var2 = ClsMasterKaryawan.EnumEmployStatus.Jobholder
                    End If
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "From master job position" Then
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                ElseIf JenisValue = "From master leave" Then
                    If Comparer = EnumComparer.SamaDengan Then
                        If CDbl(Var1) = CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    If Comparer = EnumComparer.TidakSamaDengan Then
                        If CDbl(Var1) <> CDbl(Var2) Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function

    'Private Function GetNilaiComparer(ByVal Var1 As String, ByVal Var2 As String, ByVal Comparer As EnumComparer) As Boolean
    '    Try
    '        If Comparer = EnumComparer.SamaDengan Then
    '            If Var1 = Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '        If Comparer = EnumComparer.LebihBesar Then
    '            If Var1 > Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '        If Comparer = EnumComparer.LebihKecil Then
    '            If Var1 < Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '        If Comparer = EnumComparer.LebihBesar_SamaDengan Then
    '            If Var1 >= Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '        If Comparer = EnumComparer.LebihKecil_SamaDengan Then
    '            If Var1 <= Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '        If Comparer = EnumComparer.TidakSamaDengan Then
    '            If Var1 <> Var2 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function GetUnit(ByVal PrimaryParamater As Double) As String
        Dim DTA As New DseMasterParameterTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetUnit(PrimaryParamater)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Enum EnumSearch
        Code
        Desc
        Remarks
    End Enum

    Enum EnumType
        Daily
        Monthly
        Yearly
        PersonBase
    End Enum

    Enum EnumIsVariable
        False_
        true_
    End Enum

    Enum EnumIsConstrain
        False_
        true_
    End Enum

    Public Function IsiDGV2(ByVal Criteria As String, ByVal DGV As DataGridView, _
    ByVal SearchCriteria As EnumSearch, ByVal tsl As ToolStripLabel, _
    ByVal vDta As Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_parameterTableAdapter, _
    ByVal vtbl As DseMasterSalaryFormula.tblmaster_salary_formula_parameterDataTable, _
    ByVal vType As Integer, ByVal vVariable As Integer, ByVal vConstrain As Integer) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date = Now
        Try
            vDta.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.Code Then
                QueryResult = vDta.FillBy_LikeCode(vtbl, Criteria + "%", vType, vVariable, vConstrain)
            ElseIf SearchCriteria = EnumSearch.Desc Then
                QueryResult = vDta.FillBy_LikeDesc(vtbl, vType, vVariable, Criteria + "%", vConstrain)
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = vDta.FillBy_LikeRemark(vtbl, vType, vVariable, Criteria + "%", vConstrain)
            End If
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml2(ByVal vKriteria As String, ByVal vSearchBy As EnumSearch, _
    ByVal vDta As Core_Module.DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_parameterTableAdapter, _
    ByVal vType As Integer, ByVal vVariable As Integer, ByVal vConstrain As Integer) As Double
        Dim Dummy As Nullable(Of Long)
        Try
            vDta.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearch.Code Then
                Dummy = vDta.ScalarQuery_GetJmlLikeCode(vType, vKriteria + "%", vVariable, vConstrain)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Desc Then
                Dummy = vDta.ScalarQuery_GetJmlLikeDesc(vType, vKriteria + "%", vVariable, vConstrain)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Remarks Then
                Dummy = vDta.ScalarQuery_GetJlmByRemark(vType, vKriteria + "%", vVariable, vConstrain)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vKode As String, ByVal vType As Integer, _
    ByVal vVariable As Integer, ByVal vConstrain As Integer) As Boolean
        Dim dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim Jml As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Jml = dta.ScalarQuery_GetJmlByCode(vType, vKode, vVariable, vConstrain)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryByCode(ByVal vCode As String) As Double
        Dim dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim primary As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            primary = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If primary.HasValue = True Then
                Return primary.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.DefaultButton1, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim code As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            code = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
            If String.IsNullOrEmpty(code) = False Then
                Return code
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetDescByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim desc As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            desc = dta.ScalarQuery_GetDescByPrimary(vPrimary)
            If String.IsNullOrEmpty(desc) = False Then
                Return desc
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTypeByPrimary(ByVal vPrimary As Double) As Integer
        Dim dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim type As Nullable(Of Integer)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            type = dta.ScalarQuery_GetTypeByPrimary(vPrimary)
            If type.HasValue = True Then
                Return type
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetUnitByPrimary(ByVal vPrimary As Double) As String
        Dim Dta As New DseMasterParameterGajiTableAdapters.tblmaster_salary_formula_parameterTableAdapter
        Dim unit As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            unit = Dta.ScalarQuery_GetUnitByPrimary(vPrimary)
            If String.IsNullOrEmpty(unit) = False Then
                Return unit
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetIsOutstation(ByVal vEndDate As Date, ByVal vStartDate As Date, _
   ByVal vPrimaryEmployee As Double, ByVal vTypeWorkDay_1 As EnumTipeWorkdays, _
   ByVal vTypeWorkDay_2 As EnumTipeWorkdays, ByVal vType As EnumType) As Double
        Dim DTA As New DseMasterParameterTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim StartDate, EndDate As Date

        Try
            If vType = EnumType.Yearly Then
                StartDate = vEndDate.AddYears(-1)
                StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate)
                EndDate = vEndDate.AddMonths(-1)
                EndDate = Me.ObjCD.GetLasDateOfTheMonth(EndDate)

                If vStartDate <= EndDate And EndDate <= vEndDate Then
                Else
                    Return 0
                End If
            Else
                StartDate = vStartDate
                EndDate = vEndDate
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml_Is_OutstationByDate_TypeWorkDay(vPrimaryEmployee, EndDate, StartDate, vTypeWorkDay_1, vTypeWorkDay_2)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsResign(ByVal vEndDate As Date, ByVal vStartDate As Date, _
   ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseMasterParameterTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIsResign(vPrimaryEmployee, vStartDate, vEndDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
