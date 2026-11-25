Option Explicit On
Option Strict Off

''' <summary>
''' kelas ini mengatur semua fungsi yang terkait jadwal produksi
''' </summary>
''' <remarks></remarks>
Public Class ClsTrans_Schedule
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjMySQL As New Core_Module.ClsMySQLReference

    Public Enum EnumTipeSchBar As Integer
        free_time 'waktu workshift yang sudah dialokasikan tetapi belum dapat jadwal kerja
        off_time 'break time, holiday, no work shift
        scheduled_dt
        free_ot
        work_time 'work order sesuai workshift allocate, digunakan untuk kerja
        work_time_ot 'work order sesuai sch overtime, digunakan untuk kerja
    End Enum

    Public Enum enumDurationScale As Integer
        By_Hour
        By_Day
        By_Weeks
    End Enum

    Public Function GetDefaultSchHor_StartDate() As Integer
        Return -7
    End Function

    Public Function GetDefaultSchHor_EndDate() As Integer
        Return 30
    End Function

    Public Function GetDefault_ShowFreeTime() As Boolean
        Return False
    End Function

    Public Function GetDefault_ShowOffTime() As Boolean
        Return False
    End Function

    Public Function GetDefaultSchHor_FirmPlanned1() As Integer
        Try
            Return 7
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDefaultSchHor_FirmPlanned(ByVal vCurrentTime As Date) As Date
        Try
            Return (vCurrentTime + TimeSpan.FromDays(7)).Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDefaultColor_AllocatedWS() As Color
        Return Core_Module.GetSMESSBackColorAll
    End Function
    Public Function GetDefaultColor_OffTime() As Color
        Return Color.LightGray
    End Function
    Public Function GetDefaultColor_SchDT() As Color
        Return Color.Beige
    End Function
    Public Function GetDefaultColor_SchOT() As Color
        Return Color.LightCyan
    End Function

    Public Function GetTipeSchBar_String(ByVal vTipe As EnumTipeSchBar) As String
        Try
            If vTipe = EnumTipeSchBar.free_time Then
                Return "Free Time"
            ElseIf vTipe = EnumTipeSchBar.off_time Then
                Return "Off Time"
            ElseIf vTipe = EnumTipeSchBar.scheduled_dt Then
                Return "Scheduled Downtime"
            ElseIf vTipe = EnumTipeSchBar.free_ot Then
                Return "Free Overtime"
            ElseIf vTipe = EnumTipeSchBar.work_time Then
                Return "Work Time"
            ElseIf vTipe = EnumTipeSchBar.work_time_ot Then
                Return "Work Overtime"
            End If

            Return "Off Time"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "Off Time"
        End Try
    End Function

    Public Function GetColor_AllocatedWS() As Color
        Dim a, r, g, b As String

        Try
            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_a)
            If a = "" Then
                Return Me.GetDefaultColor_AllocatedWS
            Else
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_b)

                Return Color.FromArgb(a, r, g, b)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetColor_Offtime() As Color
        Dim a, r, g, b As String

        Try
            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_a)
            If a = "" Then
                Return Me.GetDefaultColor_OffTime
            Else
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_Offtime_b)

                Return Color.FromArgb(a, r, g, b)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetColor_SchDT() As Color
        Dim a, r, g, b As String

        Try
            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_a)
            If a = "" Then
                Return Me.GetDefaultColor_SchDT
            Else
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchDT_b)

                Return Color.FromArgb(a, r, g, b)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetColor_SchOT() As Color
        Dim a, r, g, b As String

        Try
            a = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_a)
            If a = "" Then
                Return Me.GetDefaultColor_SchOT
            Else
                r = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_r)
                g = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_g)
                b = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchOT_b)

                Return Color.FromArgb(a, r, g, b)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetFirmPlan(ByVal vRefDate As Date, ByVal vCurrentDate As Date)
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, (vRefDate.Date - vCurrentDate.Date).Days)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetFirmPlan() As Date
        Dim Dummy As String
        Dim CurrentTime As Date = Today 
        Dim Dummy1 As Double
        Try
            CurrentTime = Me.ObjMySQL.GetServerTime.Date
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1)
            If Dummy = "" Then
                Dummy = Format(Me.GetDefaultSchHor_FirmPlanned(CurrentTime), "dd MMM yyyy HH:mm")
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, Dummy)
            End If

            If Double.TryParse(Dummy, Dummy1) = False Then
                Dummy = 0
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, Dummy)
            End If

            Return CDate(CurrentTime.AddDays(Dummy))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function GetStartDateSch() As Date
        Dim Dummy As String = ""
        Dim CurrentTime As Date = Today

        Try
            CurrentTime = Me.ObjMySQL.GetServerTime
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate)
            If Dummy = "" Then
                Dummy = Me.GetDefaultSchHor_StartDate
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate, Dummy)
            End If
            Return CDate(Format(CurrentTime, "dd MMM yyyy")) - TimeSpan.FromDays(Math.Abs(CInt(Dummy))) + TimeSpan.FromHours(Now.Hour) + TimeSpan.FromMinutes(Now.Minute)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetEndDateSch() As Date
        Dim Dummy As String = ""
        Dim CurrentTime As Date = Today

        Try
            CurrentTime = Me.ObjMySQL.GetServerTime
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate)
            If Dummy = "" Then
                Dummy = Me.GetDefaultSchHor_EndDate
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate, Dummy)
            End If
            Return CDate(Format(CurrentTime, "dd MMM yyyy")) + TimeSpan.FromDays(Math.Abs(CInt(Dummy)))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetLast_WS_AllocatedDate() As Date
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetMaxDate
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLast_WS_OvertimeDate() As Date
        Dim DTA As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetMaxDate
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub HitungJadwalModifikasi3(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vDummyIdentifierWO As String, ByVal vTBLWUAcuan As DseTrans_Schedule.TblWU1DataTable, ByVal vPrimaryPredefinedWU As Double)
        '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
        '2. cancel alokasi sch item dan sch bar
        '3. hitung ulang jadwal, khusus untuk identifier acuan, jumlah work unit yang dialokasikan sudah tertentu
        '4. hitung ulang jadwal sisanya
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        Dim TBL_JmlWU As New DseTrans_Schedule.TblWUDataTable
        Dim row_JmlWU As DseTrans_Schedule.TblWURow
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim a, r, g, b As Integer
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim DefColor As Color = Color.AliceBlue
        Dim ObjCombi As New ClsTrans_Schedule_Combi
        Dim TBLSchItem_Dummy As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim FirmPlanDate As Date
        Dim TBLCombination As New DseTrans_Schedule.TBlCombinationDataTable
        Dim ListPrimaryWU As New List(Of String)
        Dim vTBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable

        Dim DTAProses As New DseTrans_ScheduleTableAdapters.tbltrans_mo_master_predefine_workunit_itemprocessTableAdapter
        Dim TBLProses As New DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessDataTable
        Dim DTAWUConfig As New DseTrans_ScheduleTableAdapters.tbltrans_mo_master_predefine_workunit_itemworkunitTableAdapter
        Dim TBLWUConfig As New DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable

        Try
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring
            DTAProses.Connection.ConnectionString = DTAWU.Connection.ConnectionString
            DTAWUConfig.Connection.ConnectionString = DTAWU.Connection.ConnectionString

            If vPrimaryPredefinedWU <> -1 Then
                DTAProses.Fill(TBLProses, vPrimaryPredefinedWU)
                DTAWUConfig.Fill(TBLWUConfig, vPrimaryPredefinedWU)
            Else
                TBLProses.Rows.Clear()
                TBLWUConfig.Rows.Clear()
            End If

            '0 other preparation
            FirmPlanDate = Me.GetFirmPlan
            ObjMO.Get_argb(vIdentifierJO, a, r, g, b)
            If a = 0 And r = 0 And g = 0 And b = 0 Then
                DefColor = Me.GetDefaultColor_AllocatedWS
                a = DefColor.A
                r = DefColor.R
                g = DefColor.G
                b = DefColor.B
            End If

            '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
            TBL_JmlWU.Rows.Clear()
            Me.GetTBLJmlWU(TBL_JmlWU, vDummyIdentifierWO, vTBLItemWO, False)

            '2. cancel alokasi sch item dan sch bar
            For Each row_JmlWU In TBL_JmlWU.Rows
                Me.UnAllocateSchedule(row_JmlWU.Identifier_WO, vTBLSchBar)
                Me.UnAllocateScheduleItem(row_JmlWU.Identifier_WO, vTBLItemWO_ItemSch)
            Next

            '3. hitung ulang jadwal, khusus untuk identifier acuan, work unit yang dialokasikan sudah tertentu
            Me.LoadScheduleBar_TanpaScheduledItem(FirmPlanDate, #1/1/2200#, vTBLSchBarDummy)
            For Each rowItemWO In vTBLItemWO.Rows
                TBL_JmlWU.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier + "'"
                If TBL_JmlWU.DefaultView.Count > 0 Then
                    row_JmlWU = TBL_JmlWU.DefaultView(0).Row
                    If rowItemWO.Identifier = vDummyIdentifierWO Then
                        Me.GetProductionFinishDate2(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, rowItemWO, a, r, g, b, vTBLWUAcuan, vTBLSchBarDummy, vPrimaryPredefinedWU)
                    Else
                        Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, ObjCombi, ListPrimaryWU, TBLCombination, rowItemWO, row_JmlWU.WU_Allocated, a, r, g, b, vPrimaryPredefinedWU, TBLProses, TBLWUConfig, vTBLSchBarDummy)
                        'Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, vTBLSchItem_Dummy, DTAWU, TBLWU, _
                        'FirmPlanDate, ObjCombi, ListPrimaryWU, vTBLCombination, rowItemWO, 1, a, r, g, b, vPrimaryPredefinedWU, TBLProses, TBLWUConfig, TBLSchBar_Dummy)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWU.Dispose()
            TBLWU.Dispose()
            TBL_JmlWU.Dispose()
            ObjMO = Nothing
            ObjCombi = Nothing
            TBLSchItem_Dummy.Dispose()
            TBLCombination.Dispose()
            ListPrimaryWU = Nothing
            vTBLSchBarDummy.Dispose()
            DTAProses.Dispose()
            TBLProses.Dispose()
            DTAWUConfig.Dispose()
            TBLWUConfig.Dispose()
        End Try
    End Sub

    Public Sub HitungJadwalModifikasi2(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vDummyIdentifierWO As String, ByVal vJmlWUAcuan As Integer)
        '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
        '2. cancel alokasi sch item dan sch bar
        '3. hitung ulang jadwal, khusus untuk identifier acuan, jumlah work unit yang dialokasikan sudah tertentu
        '4. hitung ulang jadwal sisanya
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        Dim TBL_JmlWU As New DseTrans_Schedule.TblWUDataTable
        Dim row_JmlWU As DseTrans_Schedule.TblWURow
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim a, r, g, b As Integer
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim DefColor As Color = Color.AliceBlue
        Dim ObjCombi As New ClsTrans_Schedule_Combi
        Dim TBLSchItem_Dummy As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim FirmPlanDate As Date
        Dim TBLCombination As New DseTrans_Schedule.TBlCombinationDataTable
        Dim ListPrimaryWU As New List(Of String)
        Dim vTBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable

        Try
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring

            '0 other preparation
            FirmPlanDate = Me.GetFirmPlan
            ObjMO.Get_argb(vIdentifierJO, a, r, g, b)
            If a = 0 And r = 0 And g = 0 And b = 0 Then
                DefColor = Me.GetDefaultColor_AllocatedWS
                a = DefColor.A
                r = DefColor.R
                g = DefColor.G
                b = DefColor.B
            End If

            '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
            TBL_JmlWU.Rows.Clear()
            Me.GetTBLJmlWU(TBL_JmlWU, vDummyIdentifierWO, vTBLItemWO, False)

            '2. cancel alokasi sch item dan sch bar
            For Each row_JmlWU In TBL_JmlWU.Rows
                Me.UnAllocateSchedule(row_JmlWU.Identifier_WO, vTBLSchBar)
                Me.UnAllocateScheduleItem(row_JmlWU.Identifier_WO, vTBLItemWO_ItemSch)
            Next

            '3. hitung ulang jadwal, khusus untuk identifier acuan, work unit yang dialokasikan sudah tertentu
            Me.LoadScheduleBar_TanpaScheduledItem(FirmPlanDate, #1/1/2200#, vTBLSchBarDummy)
            For Each rowItemWO In vTBLItemWO.Rows
                TBL_JmlWU.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier + "'"
                If TBL_JmlWU.DefaultView.Count > 0 Then
                    row_JmlWU = TBL_JmlWU.DefaultView(0).Row
                    If rowItemWO.Identifier = vDummyIdentifierWO Then
                        Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, ObjCombi, ListPrimaryWU, TBLCombination, rowItemWO, vJmlWUAcuan, a, r, g, b, -1, Nothing, Nothing, vTBLSchBarDummy)
                    Else
                        Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, ObjCombi, ListPrimaryWU, TBLCombination, rowItemWO, row_JmlWU.WU_Allocated, a, r, g, b, -1, Nothing, Nothing, vTBLSchBarDummy)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWU.Dispose()
            TBLWU.Dispose()
            TBL_JmlWU.Dispose()
            ObjMO = Nothing
            ObjCombi = Nothing
            TBLSchItem_Dummy.Dispose()
            TBLCombination.Dispose()
            ListPrimaryWU = Nothing
            vTBLSchBarDummy.Dispose()
        End Try
    End Sub

    Public Sub HitungJadwalModifikasi(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vDummyIdentifierWO As String, ByVal vListPrimaryWUAcuan As List(Of String), ByVal vPrimaryWUConfig As Double)
        '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
        '2. cancel alokasi sch item dan sch bar
        '3. hitung ulang jadwal, khusus untuk identifier acuan, work unit yang dialokasikan sudah tertentu
        '4. hitung ulang jadwal sisanya
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        Dim TBL_JmlWU As New DseTrans_Schedule.TblWUDataTable
        Dim row_JmlWU As DseTrans_Schedule.TblWURow
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim a, r, g, b As Integer
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim DefColor As Color = Color.AliceBlue
        Dim ObjCombi As New ClsTrans_Schedule_Combi
        Dim TBLSchItem_Dummy As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim FirmPlanDate As Date
        Dim TBLCombination As New DseTrans_Schedule.TBlCombinationDataTable
        Dim ListPrimaryWU As New List(Of String)
        Dim vTBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable

        Try
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring

            '0 other preparation
            FirmPlanDate = Me.GetFirmPlan
            ObjMO.Get_argb(vIdentifierJO, a, r, g, b)
            If a = 0 And r = 0 And g = 0 And b = 0 Then
                DefColor = Me.GetDefaultColor_AllocatedWS
                a = DefColor.A
                r = DefColor.R
                g = DefColor.G
                b = DefColor.B
            End If

            '1. identifikasi, identifier wo mana saja yang perlu di-alokasi ulang. mulai dari identifier wo acuan, turun ke bawah
            TBL_JmlWU.Rows.Clear()
            Me.GetTBLJmlWU(TBL_JmlWU, vDummyIdentifierWO, vTBLItemWO, False)

            '2. cancel alokasi sch item dan sch bar
            For Each row_JmlWU In TBL_JmlWU.Rows
                Me.UnAllocateSchedule(row_JmlWU.Identifier_WO, vTBLSchBar)
                Me.UnAllocateScheduleItem(row_JmlWU.Identifier_WO, vTBLItemWO_ItemSch)
            Next

            '3. hitung ulang jadwal, khusus untuk identifier acuan, work unit yang dialokasikan sudah tertentu
            For Each rowItemWO In vTBLItemWO.Rows
                TBL_JmlWU.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier + "'"
                If TBL_JmlWU.DefaultView.Count > 0 Then
                    row_JmlWU = TBL_JmlWU.DefaultView(0).Row
                    If rowItemWO.Identifier = vDummyIdentifierWO Then
                        Me.GetProductionFinishDate1(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, ObjCombi, TBLCombination, rowItemWO, a, r, g, b, vListPrimaryWUAcuan, vTBLSchBarDummy, vPrimaryWUConfig)
                    Else
                        Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, TBLSchItem_Dummy, DTAWU, TBLWU, _
                        FirmPlanDate, ObjCombi, ListPrimaryWU, TBLCombination, rowItemWO, row_JmlWU.WU_Allocated, a, r, g, b, -1, Nothing, Nothing, vTBLSchBarDummy)
                    End If
                End If
            Next

            '4. hitung ulang jadwal sisanya
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWU.Dispose()
            TBLWU.Dispose()
            TBL_JmlWU.Dispose()
            ObjMO = Nothing
            ObjCombi = Nothing
            TBLSchItem_Dummy.Dispose()
            TBLCombination.Dispose()
            ListPrimaryWU = Nothing
            vTBLSchBarDummy.Dispose()
        End Try
    End Sub

    Private Sub CopyTblSchBar(ByVal vTBlSumber As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vTBLTarget As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim rowAsal, rowTarget As DseTrans_Schedule.tbl_sch_barRow

        Try
            vTBLTarget.Rows.Clear()
            For Each rowAsal In vTBlSumber.Rows
                rowTarget = vTBLTarget.Newtbl_sch_barRow
                For ij As Integer = 0 To vTBLTarget.Columns.Count - 1
                    rowTarget.Item(ij) = rowAsal.Item(ij)
                Next
                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungJadwalBaru(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
   ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
   ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vTargetFinishDate As Date, ByVal vPrimaryPredefinedWU As Double, _
   ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable, ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vFrmProgress As FrmTrans_ProductionSchedule_Progress2)
        Dim TotalStep As Integer = 2
        Dim CurStep As Integer = 0
        Dim DummyStr As String = ""
        Dim AvgTime As Decimal = 0
        Dim TotalTime As Decimal = 0
        Dim ElapsTime As Decimal = 0
        Dim RemainingTime As Decimal = 0
        Dim RemainingMin, RemainingSec As Double
        Dim StartTime As Date = Now

        Dim StartDate, EndDate, FirmPlanDate As Date
        Dim DTAItemWO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCombi As New ClsTrans_Schedule_Combi
        Dim ListPrimaryWU As New List(Of String)
        Dim vTBLCombination As New DseTrans_Schedule.TBlCombinationDataTable
        Dim vTBLSchItem_Dummy As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim TBLSchBar_Dummy As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim a, r, g, b As Integer
        Dim DefColor As Color = Color.AliceBlue
        Dim RateProdGabungan As Decimal = 0
        Dim TotalWH As Decimal = 0
        Dim StartDate1 As Date = Now

        Dim DTAProses As New DseTrans_ScheduleTableAdapters.tbltrans_mo_master_predefine_workunit_itemprocessTableAdapter
        Dim TBLProses As New DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessDataTable
        Dim DTAWUConfig As New DseTrans_ScheduleTableAdapters.tbltrans_mo_master_predefine_workunit_itemworkunitTableAdapter
        Dim TBLWUConfig As New DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable

        '1. load schedule bar
        '2. load detail item wo, sesuai level
        '3. untuk setiap item wo, alocate pada 1 work unit, 
        '4. jika sampai yang terakhir, tapi kemudian masih lewat target date, maka optimasi ulang sampai habis

        Try
            If vPrimaryPredefinedWU = -1 Then
                TotalStep += 1
            End If

            DTAItemWO.Connection.ConnectionString = Core_Module.GetConstring
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring
            DTAProses.Connection.ConnectionString = DTAItemWO.Connection.ConnectionString
            DTAWUConfig.Connection.ConnectionString = DTAItemWO.Connection.ConnectionString

            'load schedule bar sesuai schedule horizon, mulai dari firm planned+1 s.d 3 hari setelah item allocate workshift terakhir, 
            'schedule downtime terakhir, scheduled overtime terakhir (mana yang terakhir)
            CurStep += 1
            vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": loading gantt chart"
            vFrmProgress.LblProgress.Text = ""
            vFrmProgress.prg1.Value = 0
            FirmPlanDate = Me.GetFirmPlan
            StartDate = FirmPlanDate + TimeSpan.FromDays(1)
            EndDate = Me.GetLast_WS_AllocatedDate
            If Me.GetLast_WS_OvertimeDate.Date > EndDate Then
                EndDate = Me.GetLast_WS_OvertimeDate.Date
            End If

            '1. load schedule bar
            Me.LoadScheduleBar(StartDate, EndDate, vTBLSchBar, vTBLWU)
            Me.LoadScheduleBar_TanpaScheduledItem(StartDate, EndDate, TBLSchBar_Dummy)
            Me.CopyTblSchBar(vTBLSchBar, vTBLSchBarKosong)
            Me.CopyTblSchBar(TBLSchBar_Dummy, vTBLSchBarTanpaWO)

            '2. load detail item wo, sesuai level
            DTAItemWO.Fill1(vTBLItemWO, vIdentifierJO)

            'jika a,r,g,b masih belum di set, maka set ke warna default
            ObjMO.Get_argb(vIdentifierJO, a, r, g, b)
            If a = 0 And r = 0 And g = 0 And b = 0 Then
                DefColor = Me.GetDefaultColor_AllocatedWS
                a = DefColor.A
                r = DefColor.R
                g = DefColor.G
                b = DefColor.B
            End If

            '3. untuk setiap item wo, alocate pada 1 work unit, 
            'jika 1 work unit tidak sampe target, maka tambah work unit sampai salah satu kondisi tercapai:
            'a. semua work unit sudah di alokasi
            'b. target tercapai
            If vPrimaryPredefinedWU <> -1 Then
                DTAProses.Fill(TBLProses, vPrimaryPredefinedWU)
                DTAWUConfig.Fill(TBLWUConfig, vPrimaryPredefinedWU)
            Else
                TBLProses.Rows.Clear()
                TBLWUConfig.Rows.Clear()
            End If

            CurStep += 1
            vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": calculating schedule"
            DummyStr = vFrmProgress.LblDescription.Text
            vFrmProgress.LblProgress.Text = ""
            vFrmProgress.prg1.Value = 0
            vFrmProgress.prg1.Maximum = vTBLItemWO.Rows.Count
            StartTime = Now

            For Each rowItemWO In vTBLItemWO.Rows
                vFrmProgress.prg1.Value += 1
                vFrmProgress.LblProgress.Text = "Data " + vFrmProgress.prg1.Value.ToString + " of " + vFrmProgress.prg1.Maximum.ToString + " - Process: " + rowItemWO.Process_Code + " - WO Number: " + rowItemWO.WO_Number
                If vFrmProgress.prg1.Value > 1 Then
                    ElapsTime = (Now - StartTime).TotalSeconds
                    AvgTime = ElapsTime / (vFrmProgress.prg1.Value - 1)
                    TotalTime = AvgTime * vTBLItemWO.Rows.Count
                    RemainingTime = TotalTime - ElapsTime
                    If RemainingTime < 0 Then
                        RemainingTime = 0
                    End If
                    RemainingMin = Math.Floor(RemainingTime / 60)
                    RemainingSec = Math.Floor(RemainingTime - (RemainingMin * 60))

                    If RemainingMin > 0 Then
                        vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": calculating schedule -  est. remaining time: " + ObjCD.DigitGrouping(RemainingMin, 0) + " min " + ObjCD.DigitGrouping(RemainingSec, 0) + " sec"
                    Else
                        vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": calculating schedule -  est. remaining time: " + ObjCD.DigitGrouping(RemainingSec, 0) + " sec"
                    End If
                End If
                Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, vTBLSchItem_Dummy, DTAWU, TBLWU, _
                FirmPlanDate, ObjCombi, ListPrimaryWU, vTBLCombination, rowItemWO, 1, a, r, g, b, vPrimaryPredefinedWU, TBLProses, TBLWUConfig, TBLSchBar_Dummy)
            Next

            '4. jika sampai yang terakhir, tapi kemudian masih lewat target date, maka optimasi ulang sampai habis
            'hanya dilakukan jika di optimasi oleh software
            If vPrimaryPredefinedWU = -1 Then
                CurStep += 1
                vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": optimizing schedule"
                vFrmProgress.LblProgress.Text = ""
                vFrmProgress.prg1.Value = 0

                If Me.GetPlanedFinishDate(vTBLItemWO_ItemSch) > vTargetFinishDate Then
                    Me.OptimasiJadwalBaru(vTBLSchBar, vIdentifierJO, vTBLItemWO, vTBLItemWO_ItemSch, vTargetFinishDate, _
                    DTAWU, TBLWU, FirmPlanDate, ObjCombi, ListPrimaryWU, vTBLCombination, vTBLSchItem_Dummy, a, r, g, b, TBLSchBar_Dummy)
                End If
            End If

            'MsgBox("Total time:" + (Now - StartDate1).TotalSeconds.ToString + " s")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemWO.Dispose()
            DTAWU.Dispose()
            TBLWU.Dispose()
            ObjCD = Nothing
            ObjCombi = Nothing
            ListPrimaryWU = Nothing
            vTBLCombination = Nothing
            vTBLSchItem_Dummy = Nothing
            ObjMO = Nothing
            DTAProses.Dispose()
            TBLProses.Dispose()
            DTAWUConfig.Dispose()
            TBLWUConfig.Dispose()
            TBLSchBar_Dummy.Dispose()
        End Try
    End Sub

    Public Sub HitungJadwalBaru(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vTargetFinishDate As Date, ByVal vPrimaryPredefinedWU As Double, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable, ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim ij As New FrmTrans_ProductionSchedule_Progress2

        Try
            ij.ObjSch = Me
            ij.vTBLSchBar = vTBLSchBar
            ij.vIdentifierJO = vIdentifierJO
            ij.vTBLItemWO = vTBLItemWO
            ij.vTBLItemWO_ItemSch = vTBLItemWO_ItemSch
            ij.vTargetFinishDate = vTargetFinishDate
            ij.vPrimaryPredefinedWU = vPrimaryPredefinedWU
            ij.vTBLSchBarKosong = vTBLSchBarKosong
            ij.vTBLSchBarTanpaWO = vTBLSchBarTanpaWO
            ij.vTBLWU = vTBLWU
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub OptimasiJadwalBaru(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vTargetFinishDate As Date, _
    ByVal DTAWU As DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter, _
    ByVal TBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, ByVal vFirmPlanDate As Date, _
    ByVal vObjCombi As ClsTrans_Schedule_Combi, ByVal vListPrimaryWU As List(Of String), ByVal vTBLCombination As DseTrans_Schedule.TBlCombinationDataTable, _
    ByVal vTBLSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
    ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable)
        'Dim RateMin As Decimal = 0
        Dim DummyIdentifier As String = ""
        Dim TBL_JmlWU As New DseTrans_Schedule.TblWUDataTable
        Dim row_JmlWU As DseTrans_Schedule.TblWURow
        Dim Recalculate As Boolean = False
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        'Dim Ketemu As Boolean = False
        Dim JmlLoop As Integer = 0
        Dim DummyIdentifierB4 As String = ""
        Dim ListIdentifierBanned As New List(Of String)
        Dim FinishDateRef As Date
        Dim TBlAddAcuan As New DseTrans_Schedule.DataTable1DataTable
        Dim rowAddAcuan As DseTrans_Schedule.DataTable1Row
        Dim JmlAddAcuan As Integer = 0

        Try
            'Exit Sub
            '4.1. untuk setiap item wo, mulai dari yang rate gabungan paling kecil, selama masih ada item wo yang bisa tambah alokasi work unit, dengan syarat
            '   4.1.a. jumlah mesin yang dialokasi < total mesin
            '   4.1.b. jika tidak ada item wo yang memenuhi syarat, maka keluar dari fungsi
            While Me.GetIdentifierWO_RateMin(DummyIdentifier, vTBLItemWO, FinishDateRef, ListIdentifierBanned) = True And Me.GetPlanedFinishDate(vTBLItemWO_ItemSch).Date > vTargetFinishDate.Date
                'Exit Sub
                JmlLoop += 1

                'If JmlLoop = 17 Then
                '    Exit Sub
                'End If

                '4.2. alokasikan tambahan work unit
                '   4.2.a. catat item wo yang harus di jadwalkan ulang, mulai dari dummy identifier, dan jumlah mesinnya
                '       khusus untuk item wo yang identifiernya = dummy identifier dummy, naikan jumlah wu +1
                TBL_JmlWU.Rows.Clear()
                JmlAddAcuan = 0
                TBlAddAcuan.DefaultView.RowFilter = "[Identifier WO] = '" + DummyIdentifier + "'"
                If TBlAddAcuan.DefaultView.Count > 0 Then
                    rowAddAcuan = TBlAddAcuan.DefaultView(0).Row
                    JmlAddAcuan = rowAddAcuan.Jml_Add
                End If
                Me.GetTBLJmlWU(TBL_JmlWU, DummyIdentifier, vTBLItemWO, True, JmlAddAcuan)

                '   4.2.b. untuk setiap item wo yang akan dijadwal ulang 
                '       4.2.b.i.  konversi balik item sch bar
                '       4.2.b.ii. buang baris di tbl item sch wo
                For Each row_JmlWU In TBL_JmlWU.Rows
                    Me.UnAllocateSchedule(row_JmlWU.Identifier_WO, vTBLSchBar)
                    Me.UnAllocateScheduleItem(row_JmlWU.Identifier_WO, vTBLItemWO_ItemSch)
                Next

                '   4.3. untuk setiap item wo yang perlu dijadwalkan ulang, lakukan jadwal ulang dengan jumlah wu yang dialokasikan
                For Each rowItemWO In vTBLItemWO.Rows
                    TBL_JmlWU.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier + "'"
                    If TBL_JmlWU.DefaultView.Count > 0 Then
                        row_JmlWU = TBL_JmlWU.DefaultView(0).Row
                        Me.GetProductionFinishDate(vTBLSchBar, vTBLItemWO, vTBLItemWO_ItemSch, vTBLSchItem_Dummy, DTAWU, TBLWU, _
                        vFirmPlanDate, vObjCombi, vListPrimaryWU, vTBLCombination, rowItemWO, row_JmlWU.WU_Allocated, a, r, g, b, -1, Nothing, Nothing, vTBLSchBarDummy)

                        'cek untuk menghindari siklik bundlik karena penambahn jumlah work unit yang diallocate tidak ngaruh pada 
                        'percepatan jadwal produksi
                        If DummyIdentifierB4 = DummyIdentifier And rowItemWO.Identifier = DummyIdentifierB4 Then
                            If rowItemWO.Plan_End_Time <= FinishDateRef And rowItemWO.WU_Allocated < row_JmlWU.WU_Allocated Then
                                If row_JmlWU.WU_Allocated + 1 <= TBLWU.Rows.Count Then
                                    TBlAddAcuan.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier + "'"
                                    If TBlAddAcuan.DefaultView.Count = 0 Then
                                        rowAddAcuan = TBlAddAcuan.NewDataTable1Row
                                        rowAddAcuan.Identifier_WO = rowItemWO.Identifier
                                        rowAddAcuan.Jml_Add = 1
                                        TBlAddAcuan.Rows.Add(rowAddAcuan)
                                    Else
                                        rowAddAcuan = TBlAddAcuan.DefaultView(0).Row
                                        rowAddAcuan.Jml_Add += 1
                                    End If
                                Else
                                    'work unit sudah terpakai semua, kalo mau optimasi yang lain
                                    'selain proses ini
                                    ListIdentifierBanned.Add(rowItemWO.Identifier)
                                End If
                            ElseIf row_JmlWU.WU_Allocated = TBLWU.Rows.Count Then
                                ListIdentifierBanned.Add(rowItemWO.Identifier)
                            End If
                        End If
                    End If
                Next

                DummyIdentifierB4 = DummyIdentifier
                'If JmlLoop = 16 Then
                '    Exit Sub
                'End If
            End While

            'MsgBox(JmlLoop.ToString)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL_JmlWU.Dispose()
            ListIdentifierBanned = Nothing
            TBlAddAcuan.Dispose()
        End Try
    End Sub

    Private Sub GetProductionFinishDate2(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
  ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
  ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
  ByVal vTBLSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
  ByVal vDTAWU As DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter, _
  ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, ByVal vFirmPlanDate As Date, _
  ByVal rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
  ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal vTBLWUAcuan As DseTrans_Schedule.TblWU1DataTable, _
  ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vPrimaryWUConfig As Double)
        Dim rowSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim TotalWH As Decimal = 0
        Dim TotalProduce As Decimal = 0
        Dim JmlWUActive As Integer = 0
        Dim DummyRate As Decimal = 0

        Try
            vTBLWU.Rows.Clear()
            vDTAWU.FillBy_PrimaryProduct(vTBLWU, rowItemWO.Primary_Product)
            For Each rowWU In vTBLWU.Rows
                rowWU.effective_rate = Math.Round((rowWU.efficiency * rowWU.rate_per_hour) / 100, rowWU._Decimal, MidpointRounding.AwayFromZero)
            Next
            'tentukan start time untuk setiap work unit
            Me.SetStartTime(vTBLWU, rowItemWO.Identifier, vTBLSchBar, vFirmPlanDate, vTBLItemWO, vTBLSchBarDummy)

            'hitung tanggal selesai produksi
            Me.GenerateRowSch2(vTBLWUAcuan, vTBLWU, vTBLSchItem_Dummy, rowItemWO.Primary_MO, rowItemWO.Identifier_JO, rowItemWO.Identifier, rowItemWO.Primary_Process, _
            a, r, g, b, rowItemWO.Primary_Unit, vPrimaryWUConfig)
            Me.GetProductionFinishDate3(vTBLSchBar, rowItemWO, vTBLSchItem_Dummy, ObjCD, vTBLWUAcuan)

            'update rate work unit gabungan + end time
            'rate ini adalah penjumlahan rate dari semua work unit yang aktif bekerja 
            TotalProduce = 0
            TotalWH = 0
            For Each rowSchItem_Dummy In vTBLSchItem_Dummy.Rows
                If rowSchItem_Dummy.qty_produce + rowSchItem_Dummy.qty_produce_add > 0 Then
                    TotalProduce += rowSchItem_Dummy.qty_produce_add + rowSchItem_Dummy.qty_produce
                    TotalWH += rowSchItem_Dummy.plan_work_duration
                    DummyRate += rowSchItem_Dummy.plan_prod_rate
                    JmlWUActive += 1
                End If
            Next
            If JmlWUActive = 1 Then
                rowItemWO.Plan_Prod_Rate = DummyRate
            Else
                If TotalWH <> 0 Then
                    rowItemWO.Plan_Prod_Rate = Math.Round((TotalProduce * 60) / (TotalWH / JmlWUActive), rowItemWO._Decimal, MidpointRounding.AwayFromZero)
                Else
                    rowItemWO.Plan_Prod_Rate = Math.Round(0, MidpointRounding.AwayFromZero)
                End If
            End If
            rowItemWO.Plan_Work_Hour = TotalWH
            rowItemWO.WU_Allocated = JmlWUActive
            rowItemWO.WU_Available = vTBLWU.Rows.Count

            'masukan ke item jadwal
            Me.CopyItemJadwal(vTBLItemWO_ItemSch, vTBLSchItem_Dummy, rowItemWO.Identifier)
            rowItemWO.Plan_End_Time = Me.GetPlanedFinishDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Start_Time = Me.GetPlanedStartDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Total_Time = rowItemWO.Plan_Setup_Time + rowItemWO.Plan_Work_Hour + rowItemWO.Plan_Delay_Time

            'masukkan ke sch bar, supaya item selanjutnya memperhitungkan jadwal yang baru saja di alokasikan
            Me.AllocateSchedule(vTBLSchBar, vTBLSchItem_Dummy, a, r, g, b, rowItemWO, vPrimaryWUConfig)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetProductionFinishDate1(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vTBLSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vDTAWU As DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, ByVal vFirmPlanDate As Date, _
    ByVal vObjCombi As ClsTrans_Schedule_Combi, ByVal vTBLCombination As DseTrans_Schedule.TBlCombinationDataTable, _
    ByVal rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
    ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal vListPrimaryWUAcuan As List(Of String), _
    ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vPrimaryWUConfig As Double)
        Dim rowSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim rowCombination As DseTrans_Schedule.TBlCombinationRow
        Dim TotalWH As Decimal = 0
        Dim TotalProduce As Decimal = 0
        Dim JmlWUActive As Integer = 0
        Dim DummyRate As Decimal = 0

        Try
            vTBLWU.Rows.Clear()
            vDTAWU.FillBy_PrimaryProduct(vTBLWU, rowItemWO.Primary_Product)
            For Each rowWU In vTBLWU.Rows
                rowWU.effective_rate = Math.Round((rowWU.efficiency * rowWU.rate_per_hour) / 100, rowWU._Decimal, MidpointRounding.AwayFromZero)
            Next
            'tentukan start time untuk setiap work unit
            Me.SetStartTime(vTBLWU, rowItemWO.Identifier, vTBLSchBar, vFirmPlanDate, vTBLItemWO, vTBLSchBarDummy)

            For Each Combi As String In vListPrimaryWUAcuan
                rowCombination = vTBLCombination.NewTBlCombinationRow
                rowCombination.Combination = Combi
                rowCombination.End_Date = #1/1/1950#
                vTBLCombination.Rows.Add(rowCombination)
            Next

            'hitung tanggal selesai produksi
            Me.GetProductionFinishDate(vTBLSchBar, vTBLCombination, vTBLWU, rowItemWO, vTBLSchItem_Dummy, ObjCD, a, r, g, b)

            'update rate work unit gabungan + end time
            'rate ini adalah penjumlahan rate dari semua work unit yang aktif bekerja 
            TotalProduce = 0
            TotalWH = 0
            For Each rowSchItem_Dummy In vTBLSchItem_Dummy.Rows
                If rowSchItem_Dummy.qty_produce + rowSchItem_Dummy.qty_produce_add > 0 Then
                    TotalProduce += rowSchItem_Dummy.qty_produce_add + rowSchItem_Dummy.qty_produce
                    TotalWH += rowSchItem_Dummy.plan_work_duration
                    DummyRate += rowSchItem_Dummy.plan_prod_rate
                    JmlWUActive += 1
                End If
            Next
            If JmlWUActive = 1 Then
                rowItemWO.Plan_Prod_Rate = DummyRate
            Else
                If TotalWH <> 0 Then
                    rowItemWO.Plan_Prod_Rate = Math.Round((TotalProduce * 60) / (TotalWH / JmlWUActive), rowItemWO._Decimal, MidpointRounding.AwayFromZero)
                Else
                    rowItemWO.Plan_Prod_Rate = Math.Round(0, MidpointRounding.AwayFromZero)
                End If
            End If
            rowItemWO.Plan_Work_Hour = TotalWH
            rowItemWO.WU_Allocated = JmlWUActive
            rowItemWO.WU_Available = vTBLWU.Rows.Count

            'masukan ke item jadwal
            Me.CopyItemJadwal(vTBLItemWO_ItemSch, vTBLSchItem_Dummy, rowItemWO.Identifier)
            rowItemWO.Plan_End_Time = Me.GetPlanedFinishDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Start_Time = Me.GetPlanedStartDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Total_Time = rowItemWO.Plan_Setup_Time + rowItemWO.Plan_Work_Hour + rowItemWO.Plan_Delay_Time

            'masukkan ke sch bar, supaya item selanjutnya memperhitungkan jadwal yang baru saja di alokasikan
            Me.AllocateSchedule(vTBLSchBar, vTBLSchItem_Dummy, a, r, g, b, rowItemWO, vPrimaryWUConfig)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetProductionFinishDate(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vTBLSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vDTAWU As DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, ByVal vFirmPlanDate As Date, _
    ByVal vObjCombi As ClsTrans_Schedule_Combi, ByVal vListPrimaryWU As List(Of String), ByVal vTBLCombination As DseTrans_Schedule.TBlCombinationDataTable, _
    ByVal rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal AllocatedWU As Integer, _
    ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal vPrimaryPredefinedWU As Double, _
    ByVal vTBLProses As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessDataTable, _
    ByVal vTBLWUConfig As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitDataTable, _
    ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim rowSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim rowCombination As DseTrans_Schedule.TBlCombinationRow
        Dim TotalWH As Decimal = 0
        Dim TotalProduce As Decimal = 0
        Dim JmlWUActive As Integer = 0
        Dim DummyRate As Decimal = 0
        Dim rowWUConfig As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunitRow
        Dim rowProses As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemprocessRow

        Try
            vTBLWU.Rows.Clear()
            vDTAWU.FillBy_PrimaryProduct(vTBLWU, rowItemWO.Primary_Product)
            For Each rowWU In vTBLWU.Rows
                rowWU.effective_rate = Math.Round((rowWU.efficiency * rowWU.rate_per_hour) / 100, rowWU._Decimal, MidpointRounding.AwayFromZero)
            Next
            If vPrimaryPredefinedWU <> -1 Then
                For Each rowProses In vTBLProses.Rows
                    If rowProses.primary_product = rowItemWO.Primary_Product And rowProses.primary_process = rowItemWO.Primary_Process Then
                        For Each rowWUConfig In vTBLWUConfig.Rows
                            If rowWUConfig.identifier_process = rowProses.identifier Then
                                For Each rowWU In vTBLWU.Rows
                                    If rowWU.primary_work_unit = rowWUConfig.primary_work_unit Then
                                        rowWU.effective_rate = rowWUConfig.plan_rate
                                        rowWU.efficiency = rowWUConfig.effeciency
                                        rowWU.rate_per_hour = rowWUConfig.rate
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                        Exit For
                    End If
                Next
            End If
            'tentukan start time untuk setiap work unit
            Me.SetStartTime(vTBLWU, rowItemWO.Identifier, vTBLSchBar, vFirmPlanDate, vTBLItemWO, vTBLSchBarDummy)

            vObjCombi.GetListKombinasi(vTBLWU, AllocatedWU, vListPrimaryWU, vPrimaryPredefinedWU, vTBLProses, vTBLWUConfig, rowItemWO.Primary_Product, rowItemWO.Primary_Process)
            vTBLCombination.Clear()
            For Each Combi As String In vListPrimaryWU
                rowCombination = vTBLCombination.NewTBlCombinationRow
                rowCombination.Combination = Combi
                rowCombination.End_Date = #1/1/1950#
                vTBLCombination.Rows.Add(rowCombination)
            Next

            'hitung tanggal selesai produksi
            Me.GetProductionFinishDate(vTBLSchBar, vTBLCombination, vTBLWU, rowItemWO, vTBLSchItem_Dummy, ObjCD, a, r, g, b)

            'update rate work unit gabungan + end time
            'rate ini adalah penjumlahan rate dari semua work unit yang aktif bekerja 
            TotalProduce = 0
            TotalWH = 0
            For Each rowSchItem_Dummy In vTBLSchItem_Dummy.Rows
                If rowSchItem_Dummy.qty_produce + rowSchItem_Dummy.qty_produce_add > 0 Then
                    TotalProduce += rowSchItem_Dummy.qty_produce_add + rowSchItem_Dummy.qty_produce
                    TotalWH += rowSchItem_Dummy.plan_work_duration
                    DummyRate += rowSchItem_Dummy.plan_prod_rate

                    JmlWUActive += 1
                End If
                rowSchItem_Dummy.tipe_dependency = rowItemWO.tipe_dependency
                rowSchItem_Dummy.delay_dependency = rowItemWO.delay_dependency
            Next
            If JmlWUActive = 1 Then
                rowItemWO.Plan_Prod_Rate = DummyRate
            Else
                If TotalWH <> 0 Then
                    rowItemWO.Plan_Prod_Rate = Math.Round((TotalProduce * 60) / (TotalWH / JmlWUActive), rowItemWO._Decimal, MidpointRounding.AwayFromZero)
                Else
                    rowItemWO.Plan_Prod_Rate = Math.Round(0, MidpointRounding.AwayFromZero)
                End If
            End If
            rowItemWO.Plan_Work_Hour = TotalWH
            rowItemWO.WU_Allocated = JmlWUActive
            rowItemWO.WU_Available = vTBLWU.Rows.Count

            'masukan ke item jadwal
            Me.CopyItemJadwal(vTBLItemWO_ItemSch, vTBLSchItem_Dummy, rowItemWO.Identifier)

            rowItemWO.Plan_End_Time = Me.GetPlanedFinishDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Start_Time = Me.GetPlanedStartDate(vTBLItemWO_ItemSch, rowItemWO.Identifier)
            rowItemWO.Plan_Total_Time = rowItemWO.Plan_Setup_Time + rowItemWO.Plan_Work_Hour + rowItemWO.Plan_Delay_Time

            'masukkan ke sch bar, supaya item selanjutnya memperhitungkan jadwal yang baru saja di alokasikan
            If rowItemWO.Plan_End_Time <> #1/1/2200# Then
                Me.AllocateSchedule(vTBLSchBar, vTBLSchItem_Dummy, a, r, g, b, rowItemWO, vPrimaryPredefinedWU)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'mengembalikan semua item sch bar, dari work item ke free time atau free ot, sesuai kondisi awalnya
    Private Sub UnAllocateSchedule(ByVal vIdentifierWO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim ListPrimaryWU As New List(Of Double)
        Dim ij As Double = 0
        Dim rowSchBar, rowSchBar1, rowSchBar2 As DseTrans_Schedule.tbl_sch_barRow
        Dim MaxDate, MinDate As Date
        Dim FreeBarColor As Color = GetColor_AllocatedWS()
        Dim jk As Double = 0
        Dim Lanjut As Boolean = True

        Try
            '1. cari semua primary work unit yang digunakan oleh item wo ybs
            vTBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + vIdentifierWO + "' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            For ij = 0 To vTBLSchBar.DefaultView.Count - 1
                rowSchBar = vTBLSchBar.DefaultView(ij).Row
                If ListPrimaryWU.Contains(rowSchBar.primary_work_unit) = False Then
                    ListPrimaryWU.Add(rowSchBar.primary_work_unit)
                End If
            Next

            '2. untuk semua primary work unit
            '   2.a. catat nilai max dan min date untuk kombinasi work unit dan identifier wo
            '   2.b. buang sch bar yang merepresentasikan proses pengerjaan item wo yang sesuai identifiernya
            '   2.c. filter tbl sch bar untuk mendapatkan semua bar pada work unit ybs
            '   2.d. untuk semua sch bar work ot, yang berada sama span yang diremove, ubah kembali menjadi free ot
            '   2.e. dari min date - max date
            '       2.e.1. jika ketemu sch bar work ot, maka balikin jadi free ot
            '       2.e.2. jika ketemu bolong, maka isikan sch bar free yang bersesuian
            For Each PrimaryWU As Double In ListPrimaryWU
                '   2.a. catat nilai max dan min date untuk kombinasi work unit dan identifier wo
                MaxDate = Date.MinValue
                MinDate = Date.MaxValue
                vTBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + vIdentifierWO + "' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString + _
                " AND [primary work unit] = " + PrimaryWU.ToString
                For ij = 0 To vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(0).Row
                    If rowSchBar.primary_work_unit = PrimaryWU Then
                        MinDate = rowSchBar.plan_start_time
                        MaxDate = rowSchBar.plan_end_time

                        '   2.b. buang sch bar yang merepresentasikan proses pengerjaan item wo yang sesuai identifiernya
                        vTBLSchBar.Rows.Remove(rowSchBar)

                        Exit For
                    End If
                Next

                If MaxDate = Date.MinValue Or MinDate = Date.MaxValue Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '   2.c. filter tbl sch bar untuk mendapatkan semua bar pada work unit ybs
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"

                '   2.d. untuk semua sch bar work ot, yang berada sama span yang diremove, ubah kembali menjadi free ot
                ij = 0
                While ij <= vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(ij).Row
                    If rowSchBar.plan_start_time >= MaxDate Then
                        Exit While
                    Else
                        If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time_ot Then
                            If rowSchBar.plan_start_time >= MinDate And rowSchBar.plan_end_time <= MaxDate And rowSchBar.identifier_wo = vIdentifierWO Then
                                'balikan kembali ke free ot
                                rowSchBar.a = FreeBarColor.A
                                rowSchBar.r = FreeBarColor.R
                                rowSchBar.g = FreeBarColor.G
                                rowSchBar.b = FreeBarColor.B

                                rowSchBar.identifier_jo = ""
                                rowSchBar.identifier_wo = ""
                                rowSchBar.jo_no = ""
                                rowSchBar.mo_no = ""
                                rowSchBar.primary_mo = -1
                                rowSchBar.primary_process = -1
                                rowSchBar.process_code = ""
                                rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot
                                rowSchBar.wo_no = ""
                                rowSchBar.work_unit_code = ""
                                rowSchBar.qty_produce = 0
                                rowSchBar.rate = 0
                            End If
                        End If
                    End If
                    ij += 1
                End While

                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                For ij = 0 To vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(ij).Row
                    '   2.e. dari min date - max date
                    '       2.e.1. jika ketemu sch bar work ot, maka balikin jadi free ot
                    '       2.e.2. jika ketemu bolong, maka isikan sch bar free yang bersesuian
                    If ij = vTBLSchBar.DefaultView.Count - 1 Then
                        'jika baris terakhir, pastikan end time> max time, klo ga, maka tambah sch bar free sampe max time
                        If MaxDate > rowSchBar.plan_end_time Then
                            Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, MaxDate, "", "", -1, -1, PrimaryWU, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                        End If
                    Else
                        'hanya modifikasi sch bar yang berada dalam span work time saja
                        If rowSchBar.plan_end_time = MinDate Then
                            'sch bar terakhir, sebelum masuk ke span work time
                            If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                                'khusus untuk work time dan work time ot, agak berbeda, karena spannya bisa panjang,
                                'kemungkinan sch selanjutnya, bukan yang pas di bawahnya
                                jk = 0
                                Lanjut = True
                                rowSchBar2 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                While rowSchBar2.plan_start_time <= rowSchBar.plan_end_time
                                    jk += 1
                                    rowSchBar2 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                    If ij + jk = vTBLSchBar.DefaultView.Count - 1 Then
                                        'sampe baris terakhir, tambahkan free time jika memang perlu
                                        If rowSchBar2.plan_end_time < MaxDate Then
                                            Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar2.plan_end_time, MaxDate, "", "", -1, -1, PrimaryWU, EnumTipeSchBar.free_time, _
                                            0, 0, 0, 0, 0, -1, -1, -1, -1)
                                            Lanjut = False
                                            Exit While
                                        End If
                                    End If
                                    'jika si di tengah2 ada yang free ot, maka tidak jadi tambahkan gap
                                    'karena asumsinya semua yang diskip adalah off time
                                    If rowSchBar2.tipe = CInt(EnumTipeSchBar.free_ot) Then
                                        Lanjut = False
                                        Exit While
                                    End If
                                End While
                                rowSchBar1 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                If Lanjut = True Then
                                    If rowSchBar1.plan_start_time > rowSchBar.plan_end_time Then
                                        'tambahkan tabel gap -> free time, start:rowSchBar.plan_end_time; end: rowSchBar1.plan_start_time
                                        Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, rowSchBar1.plan_start_time, "", "", -1, -1, PrimaryWU, _
                                        EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                                    End If
                                End If
                            ElseIf rowSchBar.tipe <> ClsTrans_Schedule.EnumTipeSchBar.work_time_ot Then
                                rowSchBar1 = vTBLSchBar.DefaultView(ij + 1).Row
                                If rowSchBar1.plan_start_time > rowSchBar.plan_end_time Then
                                    'tambahkan tabel gap -> free time, start:rowSchBar.plan_end_time; end: rowSchBar1.plan_start_time
                                    Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, rowSchBar1.plan_start_time, "", "", -1, -1, _
                                    PrimaryWU, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                                End If
                            End If
                        ElseIf rowSchBar.plan_start_time >= MinDate And rowSchBar.plan_end_time <= MaxDate Then
                            If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time_ot Or rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                                'khusus untuk work time dan work time ot, agak berbeda, karena spannya bisa panjang,
                                'kemungkinan sch selanjutnya, bukan yang pas di bawahnya
                                jk = 0
                                Lanjut = True
                                rowSchBar2 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                While rowSchBar2.plan_start_time <= rowSchBar.plan_end_time
                                    jk += 1
                                    rowSchBar2 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                    If ij + jk = vTBLSchBar.DefaultView.Count - 1 Then
                                        'sampe baris terakhir, tambahkan free time jika memang perlu
                                        If rowSchBar2.plan_end_time < MaxDate Then
                                            If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                                                Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar2.plan_end_time, MaxDate, "", "", -1, -1, PrimaryWU, _
                                                EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                                            End If
                                            Lanjut = False
                                            Exit While
                                        End If
                                    End If
                                End While
                                rowSchBar1 = vTBLSchBar.DefaultView(ij + jk + 1).Row
                                If Lanjut = True Then
                                    If rowSchBar1.plan_start_time > rowSchBar.plan_end_time Then
                                        'tambahkan tabel gap -> free time, start:rowSchBar.plan_end_time; end: rowSchBar1.plan_start_time
                                        If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                                            Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, rowSchBar1.plan_start_time, "", "", -1, -1, _
                                            PrimaryWU, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                                            'Else
                                            'Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, rowSchBar1.plan_start_time, "", "", -1, -1, PrimaryWU, EnumTipeSchBar.free_ot)
                                        End If
                                    End If
                                End If
                            Else
                                rowSchBar1 = vTBLSchBar.DefaultView(ij + 1).Row
                                If rowSchBar1.plan_start_time > rowSchBar.plan_end_time Then
                                    'tambahkan tabel gap -> free time, start:rowSchBar.plan_end_time; end: rowSchBar1.plan_start_time
                                    Me.AddSchBar(vTBLSchBar, FreeBarColor, rowSchBar.plan_end_time, rowSchBar1.plan_start_time, "", "", -1, -1, _
                                    PrimaryWU, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                                End If
                            End If
                        ElseIf rowSchBar.plan_start_time >= MaxDate And rowSchBar.plan_end_time > MaxDate Then
                            'kalo sampe sini, maka sudah kelewat
                            Exit For
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWU = Nothing
        End Try
    End Sub

    'buang semua item wo sch, yang identifier wo-nya sama dengan identifier rujukan
    Private Sub UnAllocateScheduleItem(ByVal vIdentifierWO As String, ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable)
        Dim rowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            For ij As Double = vTBLItemWO_ItemSch.Rows.Count - 1 To 0 Step -1
                rowItemWO_ItemSch = vTBLItemWO_ItemSch.Rows(ij)
                If rowItemWO_ItemSch.identifier_wo = vIdentifierWO Then
                    vTBLItemWO_ItemSch.Rows.Remove(rowItemWO_ItemSch)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTBLJmlWU(ByVal vTBL_JmlWU As DseTrans_Schedule.TblWUDataTable, ByVal vDummyIdentifier As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, Optional ByVal PlusOne As Boolean = True, Optional ByVal vAddPlus As Integer = 0)
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim Catat As Boolean = False
        Dim row_JmlWU As DseTrans_Schedule.TblWURow

        Try
            vTBL_JmlWU.Rows.Clear()

            For Each rowItemWO In vTBLItemWO.Rows
                If rowItemWO.Identifier = vDummyIdentifier Then
                    Catat = True
                End If
                If Catat = True Then
                    row_JmlWU = vTBL_JmlWU.NewTblWURow
                    With row_JmlWU
                        .Identifier_WO = rowItemWO.Identifier
                        .WU_Allocated = rowItemWO.WU_Allocated
                        If vDummyIdentifier = rowItemWO.Identifier And PlusOne = True Then
                            .WU_Allocated += (1 + vAddPlus)
                        End If
                    End With
                    vTBL_JmlWU.Rows.Add(row_JmlWU)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'mengembalikan identifier wo, yang alokasi wu < max wu dan rate-nya paling kecil
    Private Function GetIdentifierWO_RateMin(ByRef vIdentifierHasil As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, ByRef vFinishDateRef As Date, ByVal vListBannedIdentifier As List(Of String)) As Boolean
        Dim DummyFilter, DummySort As String
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim Ketemu As Boolean = False

        Try
            'catat filter dan sort lama
            DummyFilter = vTBLItemWO.DefaultView.RowFilter
            DummySort = vTBLItemWO.DefaultView.Sort

            'sort berdasarkan rate prod total dan jml wu allocated
            vTBLItemWO.DefaultView.RowFilter = ""
            vTBLItemWO.DefaultView.Sort = "[Plan Prod Rate], [WU Allocated]"

            'ambil item pertama yang total allocated wu < total wu available
            For Each DV As Data.DataRowView In vTBLItemWO.DefaultView
                rowItemWO = DV.Row

                If rowItemWO.WU_Available > rowItemWO.WU_Allocated Then
                    If vListBannedIdentifier.Contains(rowItemWO.Identifier) = False Then
                        Ketemu = True
                        vIdentifierHasil = rowItemWO.Identifier
                        vFinishDateRef = rowItemWO.Plan_End_Time
                        Exit For
                    End If
                End If
            Next

            'kembalikan filter dan sort lama
            DummyFilter = vTBLItemWO.DefaultView.RowFilter
            DummySort = vTBLItemWO.DefaultView.Sort

            If Ketemu = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'kembalikan filter dan sort lama
            DummyFilter = vTBLItemWO.DefaultView.RowFilter
            DummySort = vTBLItemWO.DefaultView.Sort
        End Try
    End Function

    Private Function GetPlanedStartDate(ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vIdentifierWO As String) As Date
        Dim MinStartDate As Date = Date.MaxValue

        Try
            For Each rowWU As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow In vTBLItemWO_ItemSch.Rows
                If rowWU.identifier_wo = vIdentifierWO Then
                    If rowWU.plan_start_time < MinStartDate Then
                        MinStartDate = rowWU.plan_start_time
                    End If
                End If
            Next

            Return MinStartDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Private Function GetPlanedFinishDate(ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vIdentifierWO As String) As Date
        Dim MaxEndDate As Date = #1/1/1950#

        Try
            For Each rowWU As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow In vTBLItemWO_ItemSch.Rows
                If rowWU.identifier_wo = vIdentifierWO Then
                    If rowWU.plan_end_time > MaxEndDate Then
                        If rowWU.out_of_schedule = 0 Then
                            MaxEndDate = rowWU.plan_end_time
                        End If
                    End If
                End If
            Next

            If MaxEndDate = #1/1/1950# Then
                Return #1/1/2200#
            Else
                Return MaxEndDate
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Private Function GetPlanedFinishDate(ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable) As Date
        Dim MaxEndDate As Date = #1/1/1950#

        Try
            For Each rowWU As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow In vTBLItemWO_ItemSch.Rows
                If rowWU.plan_end_time > MaxEndDate Then
                    MaxEndDate = rowWU.plan_end_time
                End If
            Next

            Return MaxEndDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    ''' <summary>
    ''' copy item jadwal dari sumber ke target
    ''' </summary>
    ''' <param name="vTBLTarget"></param>
    ''' <param name="vTBLSumber"></param>
    ''' <param name="vIdentifierWO"></param>
    ''' <remarks></remarks>
    Private Sub CopyItemJadwal(ByVal vTBLTarget As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vTBLSumber As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vIdentifierWO As String)
        Dim rowSumber, rowTarget As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            'buang baris pada tbl target yang sesuai identifiernya
            vTBLTarget.DefaultView.RowFilter = "[identifier wo] = '" + vIdentifierWO + "'"
            While vTBLTarget.DefaultView.Count > 0
                rowTarget = vTBLTarget.DefaultView(0).Row
                vTBLTarget.Rows.Remove(rowTarget)
            End While

            'masukkan baris baru
            For Each rowSumber In vTBLSumber.Rows
                rowTarget = vTBLTarget.Newtbltrans_mo_item_wo_item_schRow
                For ij As Integer = 0 To vTBLTarget.Columns.Count - 1
                    rowTarget.Item(ij) = rowSumber.Item(ij)
                Next
                rowTarget.identifier_wo = vIdentifierWO
                vTBLTarget.Rows.Add(rowTarget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini mengembalikan tanggal selesai produksi paling cepat dan list primary work unit yang digunakan untuk produksi
    ''' </summary>
    ''' <param name="vTBLSchBar"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductionFinishDate(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTblCombination As DseTrans_Schedule.TBlCombinationDataTable, _
    ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
    ByVal vTBLSchItem_Dummy As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vObjCD As Core_Module.ClsCekData, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer) As Date
        Dim RetVal As Date = Date.MaxValue
        Dim QtyTargetProduce As Decimal = 0
        Dim rowCombination As DseTrans_Schedule.TBlCombinationRow
        Dim TBLSchItem As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim MaxDate As Date
        Dim JmlWU As Integer = 0
        Dim DummyWU As Integer = 0
        Dim QtyResultDummy As Decimal = 0

        Try
            QtyTargetProduce = vrowItemWO.Qty_Target_Prod + vrowItemWO.Qty_Target_Prod_Add
            'untuk setiap baris kombinasi
            For Each rowCombination In vTblCombination.Rows
                '1. generate baris2 pada tabel jadwal sesuai kombinasi yang diberikan -> di urutkan berdasarkan kode
                TBLSchItem.Rows.Clear()
                Me.GenerateRowSch(rowCombination, vTBLWU_Available, TBLSchItem, vrowItemWO.Primary_MO, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                vrowItemWO.Primary_Process, va, vr, vg, vb, vrowItemWO.Primary_Unit)

                '2. hitung kombinasi yang bersangkutan
                Me.GetProductionFinishDate(vTBLSchBar, vrowItemWO, TBLSchItem, vObjCD)
                JmlWU = 0
                DummyWU = 0
                QtyResultDummy = 0
                For Each rowSchItem In TBLSchItem.Rows
                    If rowSchItem.qty_produce > 0 Then
                        JmlWU += 1
                    End If
                Next
                For Each rowSchItem In TBLSchItem.Rows
                    If rowSchItem.qty_produce > 0 Then
                        DummyWU += 1
                        rowSchItem.plan_good_percentage = vrowItemWO.Plan_Good_Perct
                        If DummyWU < JmlWU Then
                            rowSchItem.plan_result = Math.Round((rowSchItem.plan_good_percentage * rowSchItem.qty_produce) / 100, CInt(rowSchItem._Decimal), MidpointRounding.AwayFromZero)
                            QtyResultDummy += rowSchItem.plan_result
                        Else
                            rowSchItem.plan_result = Math.Round(vrowItemWO.Qty_Target_Result - QtyResultDummy, CInt(rowSchItem._Decimal), MidpointRounding.AwayFromZero)
                        End If
                    End If
                Next

                '3. update nilai ke kolom end date
                MaxDate = #1/1/1950#
                For Each rowSchItem In TBLSchItem.Rows
                    If rowSchItem.plan_end_time > MaxDate Then
                        MaxDate = rowSchItem.plan_end_time
                    End If
                Next
                rowCombination.End_Date = MaxDate

                '4. tandai sebagai kombinasi paling baik, jika memenuhi kriteri
                If rowCombination.End_Date < RetVal Then
                    RetVal = rowCombination.End_Date

                    'isikan tbl sch item dummy, untuk digunakan allocate
                    vTBLSchItem_Dummy.Rows.Clear()
                    Me.CopyItemJadwal(vTBLSchItem_Dummy, TBLSchItem, "")
                End If
            Next


            'return nilai yang paling tepat
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return RetVal
        Finally
            TBLSchItem.Dispose()
        End Try
    End Function

    Private Sub GenerateRowSch2(ByVal vTblWUAcuan As DseTrans_Schedule.TblWU1DataTable, _
    ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryProses As Double, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer, ByVal vPrimaryUnit As Double, ByVal vPrimaryWUConfig As Double)
        Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowWUAcuan As DseTrans_Schedule.TblWU1Row

        Try
            TBLSchItem.Rows.Clear()

            For Each rowWUAcuan In vTblWUAcuan.Rows
                vTBLWU_Available.DefaultView.RowFilter = "[primary work unit] = " + rowWUAcuan.Primary_WU.ToString
                If vTBLWU_Available.DefaultView.Count = 0 Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    TBLSchItem.Rows.Clear()
                    Exit Sub
                Else
                    vrowWU = vTBLWU_Available.DefaultView(0).Row
                    rowSchItem = TBLSchItem.Newtbltrans_mo_item_wo_item_schRow
                    With rowSchItem
                        'disini argb diset no dulu, nanti pas alocate baru di set warna yg sebenarnya
                        .a = va
                        .r = vr
                        .g = vg
                        .b = vb
                        'kolom2 yang ga dipake di set ke nilai default dulu, disini yang penting itungannya
                        .actual_down_time_duration = 0
                        .actual_end_time = #1/1/1950#
                        .actual_prod_rate = 0
                        .actual_setup_duration = 0
                        .actual_start_time = #1/1/1950#
                        .actual_total_duration = 0
                        .actual_work_duration = 0
                        .plan_setup_duration = 0
                        .plan_total_duration = 0
                        .primary = -1
                        .status = 0
                        .progress = 0

                        'berikut data yang dipake untuk perhitungan
                        .identifier_jo = vIdentifierJO
                        .identifier_wo = vIdentifierWO
                        .primary_mo = vPrimaryMO
                        .primary_process = vPrimaryProses

                        .primary_work_unit = vrowWU.primary_work_unit
                        .plan_prod_rate = vrowWU.effective_rate
                        .plan_start_time = vrowWU.Start_Date
                        .plan_end_time = #1/1/1950#
                        .plan_end_time_minimum = #1/1/1950#
                        .plan_current_time = .plan_start_time
                        .plan_current_time2 = .plan_start_time
                        .qty_produce = 0
                        .qty_produce_add = 0
                        .plan_work_duration = 0
                        .plan_work_duration_add = 0
                        ._Decimal = vrowWU._Decimal
                        .plan_prod_rate_normal = vrowWU.rate_per_hour
                        .plan_prod_rate_efficiency = vrowWU.efficiency
                        .plan_good_percentage = 0
                        .plan_downtime = 0
                        .tipe_dependency = 0
                        .delay_dependency = 0
                        .plan_result = 0
                        .Primary_Unit = vPrimaryUnit
                        .primary_wu_config = vPrimaryWUConfig
                        .primary_input_location = vrowWU.primary_input_location
                        .primary_output_location = vrowWU.primary_output_location
                    End With
                    TBLSchItem.Rows.Add(rowSchItem)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateRowSch(ByVal rowCombination As DseTrans_Schedule.TBlCombinationRow, _
    ByVal vTBLWU_Available As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryProses As Double, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer, ByVal vPrimaryUnit As Double)
        Dim ListWU As String()
        Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            ListWU = rowCombination.Combination.Split("-")
            For Each PrimaryWU As String In ListWU
                vTBLWU_Available.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU
                If vTBLWU_Available.DefaultView.Count = 0 Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    TBLSchItem.Rows.Clear()
                    Exit Sub
                Else
                    vrowWU = vTBLWU_Available.DefaultView(0).Row
                    rowSchItem = TBLSchItem.Newtbltrans_mo_item_wo_item_schRow
                    With rowSchItem
                        'disini argb diset no dulu, nanti pas alocate baru di set warna yg sebenarnya
                        .a = va
                        .r = vr
                        .g = vg
                        .b = vb
                        'kolom2 yang ga dipake di set ke nilai default dulu, disini yang penting itungannya
                        .actual_down_time_duration = 0
                        .actual_end_time = #1/1/1950#
                        .actual_prod_rate = 0
                        .actual_setup_duration = 0
                        .actual_start_time = #1/1/1950#
                        .actual_total_duration = 0
                        .actual_work_duration = 0
                        .plan_setup_duration = 0
                        .plan_total_duration = 0
                        .plan_downtime = 0
                        .primary = -1
                        .status = 0
                        .progress = 0

                        .identifier_jo = vIdentifierJO
                        .identifier_wo = vIdentifierWO
                        .primary_mo = vPrimaryMO
                        .primary_process = vPrimaryProses

                        'berikut data yang dipake untuk perhitungan
                        .primary_work_unit = vrowWU.primary_work_unit
                        .Primary_Unit = vPrimaryUnit
                        .plan_prod_rate = vrowWU.effective_rate
                        .plan_start_time = vrowWU.Start_Date
                        .plan_end_time = #1/1/1950#
                        .plan_end_time_minimum = #1/1/1950#
                        .plan_current_time = .plan_start_time
                        .plan_current_time2 = .plan_start_time
                        .qty_produce = 0
                        .qty_produce_add = 0
                        .plan_work_duration = 0
                        .plan_work_duration_add = 0
                        ._Decimal = vrowWU._Decimal
                        .plan_prod_rate_normal = vrowWU.rate_per_hour
                        .plan_prod_rate_efficiency = vrowWU.efficiency
                        .plan_result = 0
                        .plan_good_percentage = 0
                        .tipe_dependency = 0
                        .delay_dependency = 0
                        .primary_wu_config = -1

                        .primary_input_location = vrowWU.primary_input_location
                        .primary_output_location = vrowWU.primary_output_location
                    End With
                    TBLSchItem.Rows.Add(rowSchItem)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListWU = Nothing
        End Try
    End Sub

    Private Sub GetProductionFinishDate3(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
   ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vObjCD As Core_Module.ClsCekData, ByVal vTBLWUAcuan As DseTrans_Schedule.TblWU1DataTable)
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowWUAcuan As DseTrans_Schedule.TblWU1Row
        Dim WorkDuration As Double = 0
        Dim Lewat As Boolean = False

        Try
            'jika tabel sch item terdiri lebih dari 1 baris, maka leveling currenct time dahulu
            For Each rowSchItem In TBLSchItem.Rows
                'cari item wu yang bersesuaian
                For Each rowWUAcuan In vTBLWUAcuan.Rows
                    If rowWUAcuan.Primary_WU = rowSchItem.primary_work_unit Then
                        'naikkan per qty saja
                        rowSchItem.qty_produce = rowWUAcuan.Qty_Produce
                        WorkDuration = ObjCD.RoundUp1((rowSchItem.qty_produce * 60) / rowSchItem.plan_prod_rate, 0)
                        rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)

                        While rowSchItem.plan_current_time = #1/1/2200# And rowSchItem.qty_produce > 0
                            rowSchItem.qty_produce -= 1
                            WorkDuration = ObjCD.RoundUp1((rowSchItem.qty_produce * 60) / rowSchItem.plan_prod_rate, 0)
                            rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                            Lewat = True
                        End While
                        If Lewat = True Then
                            'tandain, supaya lain kali baris work unit ini tidak digunakan lagi
                            'rowSchItem.plan_current_time = #1/1/2200#
                            rowSchItem.out_of_schedule = 1
                        End If

                        rowSchItem.plan_work_duration = WorkDuration
                        Exit For
                    End If
                Next
            Next

            'beres2 nilai:
            '1. end time = current time
            '2. end time minimum = end time
            For Each rowSchItem In TBLSchItem.Rows
                rowSchItem.plan_end_time = rowSchItem.plan_current_time
                rowSchItem.plan_end_time_minimum = rowSchItem.plan_current_time
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetProductionFinishDate(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vObjCD As Core_Module.ClsCekData)
        Dim JmlAktif As Integer = 0
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim QtySisaProduksi As Decimal = 0
        Dim Duration_Increment_Hr As Integer = 8
        Dim QtyIncrement As Decimal = 0

        Try
            'jika tabel sch item terdiri lebih dari 1 baris, maka leveling currenct time dahulu
            For Each rowSchItem In TBLSchItem.Rows
                If rowSchItem.plan_start_time <> #1/1/2200# And rowSchItem.out_of_schedule = 0 Then
                    JmlAktif += 1
                End If
            Next
            If JmlAktif > 1 Then
                QtySisaProduksi = vrowItemWO.Qty_Target_Prod + vrowItemWO.Qty_Target_Prod_Add - Me.GetQtyProduce(TBLSchItem)
                Me.LevelCurrentTime(vTBLSchBar, vrowItemWO, TBLSchItem, QtySisaProduksi, vObjCD)
            End If

            'selama masih memungkinkan naikan qty produce 8 jam + leveling current time s.d 1 jam
            While Me.IncreaseProdQty(TBLSchItem, 8, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 7, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 6, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 5, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 4, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 3, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 2, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While
            While Me.IncreaseProdQty(TBLSchItem, 1, vTBLSchBar, vrowItemWO, vObjCD) = True
            End While

            'naikkan per qty saja
            Me.IncreaseProdQty(TBLSchItem, vTBLSchBar, vrowItemWO, vObjCD)

            'beres2 nilai:
            '1. end time = current time
            '2. end time minimum = end time
            For Each rowSchItem In TBLSchItem.Rows
                rowSchItem.plan_end_time = rowSchItem.plan_current_time
                rowSchItem.plan_end_time_minimum = rowSchItem.plan_current_time
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IncreaseProdQty(ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
   ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal vObjCD As Core_Module.ClsCekData)
        'set increment qty = 0
        'selama qty sisa produksi > 0
        'cari baris yang current time 2 paling kecil, tambahkan qty-nya
        'hitung ulang current time berdasarkan qty produce add

        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim QtySisaProduksi As Decimal = 0
        Dim TargetTime As Date = Date.MaxValue
        Dim WorkDuration As Double = 0
        Dim Lewat As Boolean = False

        Try
            'set increment qty = 0 dan curent time2 = current time
            For Each rowSchItem In vTBLSchItem.Rows
                rowSchItem.Increment_Qty = 0
                rowSchItem.plan_current_time2 = rowSchItem.plan_current_time
            Next
            'selama qty sisa produksi > 0
            QtySisaProduksi = vrowItemWO.Qty_Target_Prod + vrowItemWO.Qty_Target_Prod_Add - Me.GetQtyProduce(vTBLSchItem)

            While QtySisaProduksi > 0
                'cari target time
                TargetTime = Date.MaxValue
                For Each rowSchItem In vTBLSchItem.Rows
                    If rowSchItem.out_of_schedule = 0 Then
                        If rowSchItem.plan_current_time2 < TargetTime And rowSchItem.out_of_schedule = 0 Then
                            TargetTime = rowSchItem.plan_current_time2
                        End If
                    End If
                Next

                'kalo sampe target time = 1 jan 2200 atau date.maxvalue, artinya tidak ada lagi work unit yang bisa ditambahkan datanya
                If TargetTime = Date.MaxValue Or TargetTime = #1/1/2200# Then
                    For Each rowSchItem In vTBLSchItem.Rows
                        If rowSchItem.out_of_schedule = 1 Then
                            'rowSchItem.plan_end_time = #1/1/2200#
                            rowSchItem.plan_current_time = #1/1/2200#
                        End If
                    Next
                    Exit While
                End If

                'cari baris yang current time 2 = target time, tambahkan qty-nya, hitung ulang qty sisa produksi, hitung add work duration current time 2
                For Each rowSchItem In vTBLSchItem.Rows
                    If rowSchItem.out_of_schedule = 0 Then
                        Lewat = False
                        If rowSchItem.plan_current_time2 = TargetTime Then
                            If QtySisaProduksi >= 2 ^ (rowSchItem.Increment_Qty + 2) Then
                                rowSchItem.Increment_Qty += 2
                                rowSchItem.qty_produce_add += (2 ^ (rowSchItem.Increment_Qty))
                                QtySisaProduksi -= (2 ^ (rowSchItem.Increment_Qty))
                            ElseIf QtySisaProduksi >= 2 ^ (rowSchItem.Increment_Qty + 1) Then
                                rowSchItem.Increment_Qty += 1
                                rowSchItem.qty_produce_add += (2 ^ (rowSchItem.Increment_Qty))
                                QtySisaProduksi -= (2 ^ (rowSchItem.Increment_Qty))
                            ElseIf QtySisaProduksi >= 2 ^ (rowSchItem.Increment_Qty) Then
                                rowSchItem.Increment_Qty = rowSchItem.Increment_Qty
                                rowSchItem.qty_produce_add += (2 ^ (rowSchItem.Increment_Qty))
                                QtySisaProduksi -= (2 ^ (rowSchItem.Increment_Qty))
                            Else
                                If QtySisaProduksi < 1 Then
                                    'rowSchItem.Increment_Qty = rowSchItem.Increment_Qty - ij
                                    rowSchItem.qty_produce_add += QtySisaProduksi
                                    QtySisaProduksi = 0
                                    Exit For
                                Else
                                    For ij As Integer = 1 To rowSchItem.Increment_Qty
                                        If QtySisaProduksi >= 2 ^ (rowSchItem.Increment_Qty - ij) Then
                                            rowSchItem.Increment_Qty = rowSchItem.Increment_Qty - ij
                                            rowSchItem.qty_produce_add += (2 ^ (rowSchItem.Increment_Qty))
                                            QtySisaProduksi -= (2 ^ (rowSchItem.Increment_Qty))
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                            'rowSchItem.qty_produce_add += (2 ^ (rowSchItem.Increment_Qty))
                            'QtySisaProduksi -= (2 ^ (rowSchItem.Increment_Qty))

                            WorkDuration = ObjCD.RoundUp1((rowSchItem.qty_produce_add * 60) / rowSchItem.plan_prod_rate, 0)
                            rowSchItem.plan_current_time2 = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)

                            'pastikan qty produksi hanya terjadi pada waktu produksi yang sudah di alocate
                            While rowSchItem.plan_current_time2 = #1/1/2200# And rowSchItem.qty_produce_add > 0
                                rowSchItem.qty_produce_add -= 1
                                QtySisaProduksi += 1
                                WorkDuration = ObjCD.RoundUp1((rowSchItem.qty_produce_add * 60) / rowSchItem.plan_prod_rate, 0)
                                rowSchItem.plan_current_time2 = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                                Lewat = True
                            End While
                            If Lewat = True Then
                                'tandain, supaya lain kali baris work unit ini tidak digunakan lagi
                                'rowSchItem.plan_current_time2 = #1/1/2200#
                                rowSchItem.out_of_schedule = 1
                            End If
                            Exit For
                        End If
                    End If
                Next
            End While

            'hitung ulang current time dan work duration berdasarkan qty produce add -> kosongkan qty produce add dan work duration add
            For Each rowSchItem In vTBLSchItem.Rows
                If rowSchItem.qty_produce_add > 0 Then
                    WorkDuration = ObjCD.RoundUp1((rowSchItem.qty_produce_add * 60) / rowSchItem.plan_prod_rate, 0)
                    rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                    rowSchItem.qty_produce = rowSchItem.qty_produce + rowSchItem.qty_produce_add
                    rowSchItem.qty_produce_add = 0
                    rowSchItem.plan_work_duration = rowSchItem.plan_work_duration + WorkDuration
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IncreaseProdQty(ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vDuration_Increment_Hr As Integer, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal vObjCD As Core_Module.ClsCekData) As Boolean
        '1. cek apakah masih boleh increase, jika tidak boleh, maka return false
        '2. catat span antara current time max dan min
        '3. increase produce qty
        '4. bandingkan span antara current time max dan min, jika berubah, maka jalankan fungsi leveling current time

        Dim QtyIncrease As Decimal = 0
        Dim QtyIncrease1 As Decimal = 0
        Dim QtyProduce As Decimal = 0
        Dim QtySisaProduksi As Decimal = 0
        Dim Span As Double = 0
        Dim MinTime As Date = Date.MaxValue
        Dim MaxTime As Date = Date.MinValue
        Dim Span1 As Double = 0
        Dim MinTime1 As Date = Date.MaxValue
        Dim MaxTime1 As Date = Date.MinValue
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Lewat As Boolean = False
        Dim TargetTime As Date = Date.MinValue
        Dim WorkDuration As Double = 0
        Dim DummyEndTime As Date = #1/1/1950#

        Try
            '1. cek apakah masih boleh increase, jika tidak boleh, maka return false
            QtyIncrease = Me.GetQtyIncrement(vTBLSchItem, vDuration_Increment_Hr)
            QtyProduce = Me.GetQtyProduce(vTBLSchItem)
            If QtyIncrease + QtyProduce > vrowItemWO.Qty_Target_Prod + vrowItemWO.Qty_Target_Prod_Add Then
                Return False
            End If

            '2. catat span antara current time max dan min
            If vTBLSchItem.Rows.Count > 1 Then
                For Each rowSchItem In vTBLSchItem.Rows
                    If rowSchItem.plan_current_time <> #1/1/2200# Then
                        'hanya consider yang aktif saja
                        If rowSchItem.plan_current_time > MaxTime Then
                            MaxTime = rowSchItem.plan_current_time
                        End If
                        If rowSchItem.plan_current_time < MinTime Then
                            MinTime = rowSchItem.plan_current_time
                        End If
                    End If
                Next
                Span = (MaxTime - MinTime).TotalMinutes
            End If


            '3. increase produce qty
            For Each rowSchItem In vTBLSchItem.Rows
                If rowSchItem.plan_current_time <> #1/1/2200# And rowSchItem.out_of_schedule = 0 Then
                    'rowSchItem.plan_current_time = #12/6/2016 9:30:00 AM#
                    DummyEndTime = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, vDuration_Increment_Hr * 60, vTBLSchBar)
                    If DummyEndTime < #1/1/2200# Then
                        rowSchItem.qty_produce += rowSchItem.plan_prod_rate * vDuration_Increment_Hr
                        rowSchItem.plan_current_time = DummyEndTime
                        rowSchItem.plan_work_duration += vDuration_Increment_Hr * 60
                        QtyIncrease1 += rowSchItem.plan_prod_rate * vDuration_Increment_Hr
                    Else
                        'pastikan produksi hanya dilakukan pada workshift yang sudah terjadwal
                        TargetTime = Me.GetEndFreeTime(rowSchItem.primary_work_unit, vTBLSchBar)
                        If TargetTime <> #1/1/2200# Then
                            WorkDuration = Me.GetWorkDuration(rowSchItem.plan_current_time, TargetTime, rowSchItem.primary_work_unit, vTBLSchBar)
                            QtyProduce = ObjCD.RoundDown1((WorkDuration * rowSchItem.plan_prod_rate) / 60, rowSchItem._Decimal)

                            rowSchItem.qty_produce += QtyProduce
                            rowSchItem.plan_work_duration += WorkDuration
                            'tandain, supaya lain kali baris work unit ini tidak digunakan lagi
                            rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                            QtyIncrease1 += QtyProduce
                            rowSchItem.out_of_schedule = 1
                        Else
                            'tandain, supaya lain kali baris work unit ini tidak digunakan lagi
                            'rowSchItem.plan_current_time = #1/1/2200#
                            rowSchItem.out_of_schedule = 1
                        End If
                    End If
                End If
            Next

            '4. bandingkan span antara current time max dan min, jika berubah, maka jalankan fungsi leveling current time
            If QtyIncrease = QtyIncrease1 Then
                If vTBLSchItem.Rows.Count > 1 Then
                    For Each rowSchItem In vTBLSchItem.Rows
                        If rowSchItem.plan_current_time <> #1/1/2200# Then
                            'hanya consider yang aktif saja
                            If rowSchItem.plan_current_time > MaxTime1 Then
                                MaxTime1 = rowSchItem.plan_current_time
                            End If
                            If rowSchItem.plan_current_time < MinTime1 Then
                                MinTime1 = rowSchItem.plan_current_time
                            End If
                        End If
                    Next
                    Span1 = (MaxTime1 - MinTime1).TotalMinutes
                    If Span1 <> Span Then
                        QtySisaProduksi = vrowItemWO.Qty_Target_Prod + vrowItemWO.Qty_Target_Prod_Add - Me.GetQtyProduce(vTBLSchItem)
                        Me.LevelCurrentTime(vTBLSchBar, vrowItemWO, vTBLSchItem, QtySisaProduksi, vObjCD)
                    End If
                End If


                '5. sampe sini, fungsi berhasil dilaksanakan, return true
                Return True
            Else
                'jika target tidak tercapai karena workshift tidak cukup, return false
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyIncrement(ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vDuration_Increment_Hr As Integer) As Decimal
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowSchItem In vTBLSchItem.Rows
                RetVal += (rowSchItem.plan_prod_rate * vDuration_Increment_Hr)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyIncrement(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, ByVal vTargetDate As Date, _
    ByVal ObjCD As Core_Module.ClsCekData) As Decimal
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim RetVal As Decimal = 0
        Dim WorkDuration As Double = 0

        Try
            For Each rowSchItem In vTBLSchItem.Rows
                If rowSchItem.plan_current_time < vTargetDate Then
                    WorkDuration = Me.GetWorkDuration(rowSchItem.plan_current_time, vTargetDate, rowSchItem.primary_work_unit, vTBLSchBar)
                    RetVal += ObjCD.RoundDown1((rowSchItem.plan_prod_rate * WorkDuration) / 60, rowSchItem._Decimal)
                Else
                    WorkDuration = 0
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyProduce(ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable) As Decimal
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowSchItem In vTBLSchItem.Rows
                RetVal += (rowSchItem.qty_produce + rowSchItem.qty_produce_add)
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetFilterString_IdentifierWOParent(ByVal vIdentifierWO As String, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable) As String
        Dim RetVal As String = ""
        Dim row As DseTrans_Schedule.tbltrans_mo_item_wo1Row

        Try
            vTBLItemWO.DefaultView.RowFilter = "[Identifier WO Parent] = '" + vIdentifierWO + "'"
            For Each dv As Data.DataRowView In vTBLItemWO.DefaultView
                row = dv.Row
                If RetVal = "" Then
                    RetVal = "[identifier wo] = '" + row.Identifier + "'"
                Else
                    RetVal = RetVal + " or [identifier wo] = '" + row.Identifier + "'"
                End If
            Next

            If RetVal = "" Then
                Return "[identifier wo] = '-'"
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "[identifier wo] = '-'"
        End Try
    End Function

    Public Function GetStartTime(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String) As Date
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim row As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_identifier(TBL, vIdentifierJO, vIdentifierWO)
            If TBL.Rows.Count = 0 Then
                Return #1/1/1950#
            Else
                row = TBL.Rows(0)
                Return row.plan_start_time
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function GetStartTime1(ByVal vIdentifierJO As String, ByVal vPrimaryRM As String) As Date
        Dim DTA As New DseTrans_ScheduleTableAdapters.DataTable7TableAdapter
        Dim TBL As New DseTrans_Schedule.DataTable7DataTable
        Dim row As DseTrans_Schedule.DataTable7Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vIdentifierJO, vPrimaryRM)
            If TBL.Rows.Count = 0 Then
                Return #1/1/1950#
            Else
                row = TBL.Rows(0)
                Return row.plan_start_time
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan start time berdasarkan tipe dependency dan delay dependency
    ''' </summary>
    ''' <param name="vIdentifierJO"></param>
    ''' <param name="vIdentifierWO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStartTime(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWU As Double, _
    ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo2DataTable, _
    ByVal vTBLSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch2DataTable) As Date
        Dim ListIdentifier As New List(Of String)
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo2Row
        Dim DV As Data.DataRowView
        Dim MaxEndDate As Date = Date.MinValue
        Dim rowSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch2Row
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency
        Dim DelayDependency As Double = 0

        Try
            For Each rowItemWO In vTBLItemWO.Rows
                If rowItemWO.Identifier_WO = vIdentifierWO Then
                    TipeDependency = rowItemWO.tipe_dependency
                    DelayDependency = rowItemWO.delay_dependency
                    Exit For
                End If
            Next

            'cari semua identifier child
            vTBLItemWO.DefaultView.RowFilter = "[Identifier WO Parent] = '" + vIdentifierWO + "'"
            If vTBLItemWO.DefaultView.Count = 0 Then
                Return #1/1/1950#
            Else
                'untuk semua child, cari akhir end datenya
                For Each DV In vTBLItemWO.DefaultView
                    rowItemWO = DV.Row
                    ListIdentifier.Add(rowItemWO.Identifier_WO)
                Next

                If TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1 Then
                    vTBLSch.DefaultView.Sort = "[plan end time] DESC"
                    For Each Identifier As String In ListIdentifier
                        vTBLSch.DefaultView.RowFilter = "[identifier wo] = '" + Identifier + "'"
                        rowSch = vTBLSch.DefaultView(0).Row
                        If MaxEndDate < rowSch.plan_end_time Then
                            MaxEndDate = rowSch.plan_end_time
                        End If
                    Next
                Else
                    vTBLSch.DefaultView.Sort = "[plan start time] DESC"
                    For Each Identifier As String In ListIdentifier
                        vTBLSch.DefaultView.RowFilter = "[identifier wo] = '" + Identifier + "'"
                        rowSch = vTBLSch.DefaultView(0).Row
                        If MaxEndDate < rowSch.plan_start_time Then
                            MaxEndDate = rowSch.plan_start_time
                        End If
                    Next
                End If
            End If


            'tentukan start time berdasarkan kombinasi end time child, tipe dependency, dan delay dependency
            If DelayDependency = 0 Then
                Return MaxEndDate
            Else
                MaxEndDate = Me.GetEndTime(vPrimaryWU, MaxEndDate, DelayDependency, vTBLSchBarTanpaWO)
                Return MaxEndDate
            End If

            'kalo sampe sini, maka tidak ketemu start time yang dimaksud
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListIdentifier = Nothing
        End Try
    End Function

    Private Sub SetStartTime(ByVal TBLWU As DseTrans_Schedule.tblmaster_product_item_workunitDataTable, _
    ByVal vIdentifierWO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vFimrPlanDate As Date, ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim RetVal As Date = Date.MinValue
        Dim rowSch As DseTrans_Schedule.tbl_sch_barRow
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim AdaSch As Boolean = False
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim DelayDependency As Double = 0
        Dim vRowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim EndTime As Date = Today

        Try
            'yang menentukan start time, yang paling besar dari
            '1. dari end time parent process -> lihat dari tabel schedule
            '2. bandingkan dengan firm plan
            '3. end time terakhir dari work time untuk mesin ybs -> dari db
            '4. end item terahir dari mesin ybs, dari yang sudah terjadwal

            'Pastikan start time ada di dalam work time
            'jika tidak ada dalam work time, maka set ke #1/1/2200#, tandai mesin tersebut tidak dimasukkan ke dalam perhitungan

            'yang menentukan start time, yang paling besar dari
            '1. dari end time parent process -> lihat dari tabel schedule
            vTBLSchBar.DefaultView.RowFilter = Me.GetFilterString_IdentifierWOParent(vIdentifierWO, vTBLItemWO)
            For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
                rowSch = dv.Row
                For Each vRowItemWO In vTBLItemWO
                    If vRowItemWO.Identifier = vIdentifierWO Then
                        DelayDependency = vRowItemWO.delay_dependency
                        TipeDependency = vRowItemWO.tipe_dependency
                        Exit For
                    End If
                Next
                If TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1 Then
                    If DelayDependency = 0 Then
                        If rowSch.plan_end_time > RetVal Then
                            RetVal = rowSch.plan_end_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(rowSch.primary_work_unit, rowSch.plan_end_time, DelayDependency, vTBLSchBarDummy)
                        If EndTime > RetVal Then
                            RetVal = EndTime
                        End If
                    End If
                ElseIf TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start1 Then
                    If DelayDependency = 0 Then
                        If rowSch.plan_start_time > RetVal Then
                            RetVal = rowSch.plan_start_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(rowSch.primary_work_unit, rowSch.plan_start_time, DelayDependency, vTBLSchBarDummy)
                        If EndTime > rowSch.plan_end_time Then
                            EndTime = rowSch.plan_end_time
                        End If
                        RetVal = EndTime
                    End If
                ElseIf TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start_Strict1 Then
                    If DelayDependency = 0 Then
                        If rowSch.plan_start_time > RetVal Then
                            RetVal = rowSch.plan_start_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(rowSch.primary_work_unit, rowSch.plan_start_time, DelayDependency, vTBLSchBarDummy)
                        RetVal = EndTime
                    End If
                End If
            Next


            '2. bandingkan dengan firm plan
            If vFimrPlanDate + TimeSpan.FromDays(1) > RetVal Then
                RetVal = vFimrPlanDate + TimeSpan.FromDays(1)
            End If

            '3. end time terakhir dari work time untuk mesin ybs
            'Pastikan start time ada di dalam work time
            For Each rowWU In TBLWU.Rows
                If rowWU.IsStart_DateNull = True Then
                    rowWU.Start_Date = #1/1/1950#
                End If

                vTBLSchBar.DefaultView.RowFilter = "tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString + _
                " AND [primary work unit] = " + rowWU.primary_work_unit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

                If vTBLSchBar.DefaultView.Count = 0 Then
                    If rowWU.Start_Date < RetVal Then
                        AdaSch = False
                        rowWU.Start_Date = RetVal
                        vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
                        " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
                        ")  AND [primary work unit] = " + rowWU.primary_work_unit.ToString
                        vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

                        For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                            rowSch = DV.Row
                            If rowSch.plan_start_time <= rowWU.Start_Date And rowSch.plan_end_time >= rowWU.Start_Date Then
                                AdaSch = True
                                Exit For
                            End If
                        Next

                        If AdaSch = False Then
                            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                                rowSch = DV.Row
                                If rowSch.plan_start_time >= rowWU.Start_Date Then
                                    rowWU.Start_Date = rowSch.plan_start_time
                                    AdaSch = True
                                    Exit For
                                End If
                            Next
                        End If

                        If AdaSch = False Then
                            'kalo sampe sini artinya tidak ada workshift yang sudah allocate untuk mesin ini
                            'set start date = 1 Jan 2200
                            rowWU.Start_Date = #1/1/2200#
                        End If
                    End If
                Else
                    '4. end item terahir dari mesin ybs, dari yang sudah terjadwal
                    rowSch = vTBLSchBar.DefaultView(vTBLSchBar.DefaultView.Count - 1).Row
                    If rowSch.plan_end_time < RetVal Then
                        If rowWU.Start_Date < RetVal Then
                            AdaSch = False
                            rowWU.Start_Date = RetVal
                            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
                            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
                            ")  AND [primary work unit] = " + rowWU.primary_work_unit.ToString
                            vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

                            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                                rowSch = DV.Row
                                If rowSch.plan_start_time <= rowWU.Start_Date And rowSch.plan_end_time >= rowWU.Start_Date Then
                                    AdaSch = True
                                    Exit For
                                End If
                            Next

                            If AdaSch = False Then
                                For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                                    rowSch = DV.Row
                                    If rowSch.plan_start_time >= rowWU.Start_Date Then
                                        rowWU.Start_Date = rowSch.plan_start_time
                                        AdaSch = True
                                        Exit For
                                    End If
                                Next
                            End If

                            If AdaSch = False Then
                                'kalo sampe sini artinya tidak ada workshift yang sudah allocate untuk mesin ini
                                'set start date = 1 Jan 2200
                                rowWU.Start_Date = #1/1/2200#
                            End If
                        End If
                    Else
                        rowWU.Start_Date = rowSch.plan_end_time
                    End If
                End If

            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetProduceQty(ByVal vDuration_Hour As Integer, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable)
        Try
            MsgBox("belum")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddProduceQty(ByVal vDuration_Hour As Integer, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable)
        Try
            MsgBox("belum")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LevelCurrentTime(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal vQtySisaProduksi As Decimal, ByVal vObjCD As Core_Module.ClsCekData)
        Dim TargetTime As Date = Date.MinValue
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowSch, rowSchB4 As DseTrans_Schedule.tbl_sch_barRow
        Dim TargetTime2 As Date = Date.MinValue
        Dim AdaSch As Boolean = False
        Dim WorkDuration As Decimal = 0
        Dim QtyProduce As Decimal = 0
        'Dim TBLSchItem_Dummy As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        'Dim Lewat1 As Boolean

        Try
            '1. Tentukan Target Time -> mengabaikan yang current time = 1/1/2200
            For Each rowSchItem In TBLSchItem.Rows
                If rowSchItem.plan_current_time <> #1/1/2200# And rowSchItem.out_of_schedule = 0 Then
                    If TargetTime < rowSchItem.plan_current_time Then
                        If Me.GetQtyIncrement(vTBLSchBar, TBLSchItem, rowSchItem.plan_current_time, vObjCD) <= vQtySisaProduksi Then
                            TargetTime = rowSchItem.plan_current_time
                        Else
                            rowSchItem.out_of_schedule = 1
                        End If
                    End If
                End If
            Next

            '2. jika target time = date.MinValue  -> tidak ada mesin yang aktif, maka keluar langsung dari fungsi
            If TargetTime = Date.MinValue Then
                Exit Sub
            End If

            '3. Untuk setiap baris
            For Each rowSchItem In TBLSchItem.Rows
                'a. jika QtySisaProduksi =0, maka abaikan
                If vQtySisaProduksi > 0 Then
                    '   a. jika current time => target time -> abaikan
                    If rowSchItem.plan_current_time < TargetTime And rowSchItem.out_of_schedule = 0 Then
                        TargetTime2 = TargetTime
                        AdaSch = False

                        '   b. jika target time tidak ada dalam work hour, maka target time adalah plan end date terakhir dari sch, yang masih lebih besar dari current time
                        vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
                        " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
                        ") AND [primary work unit] = " + rowSchItem.primary_work_unit.ToString
                        vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

                        For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                            rowSch = DV.Row
                            If rowSch.plan_start_time <= TargetTime2 And rowSch.plan_end_time >= TargetTime2 Then
                                AdaSch = True
                                Exit For
                            End If
                        Next
                        If AdaSch = False Then
                            'jika tidak ada dalam jadwal maka, cari jadwal terakhir sebelum off time yang paling deket. itu jadi target time
                            For i As Double = 1 To vTBLSchBar.DefaultView.Count - 1
                                rowSch = vTBLSchBar.DefaultView(i).Row
                                rowSchB4 = vTBLSchBar.DefaultView(i - 1).Row

                                If rowSch.plan_end_time >= TargetTime2 Then
                                    If rowSch.plan_start_time <= TargetTime2 Then
                                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                        Exit Sub
                                        'ga mungkin ke sini, klo kesini maka sudah di handle di atas, ketika jatuh di work time
                                        'AdaSch = True
                                        'Exit For
                                    Else
                                        TargetTime2 = rowSchB4.plan_end_time
                                        AdaSch = True
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        'jika current time => target time 2 -> abaikan
                        If rowSchItem.plan_current_time < TargetTime2 Then
                            '   c. jika target time ada dalam work hour, maka hitung work hournya, lalu hitung jumlah produksi yang bisa dilakukan untuk work hour tersebut,
                            '      lalu tambahkan qty produce dan hitung current time, pastikan qty-nya tidak mengakibatkan qty produce terlewati
                            WorkDuration = Me.GetWorkDuration(rowSchItem.plan_current_time, TargetTime2, rowSchItem.primary_work_unit, vTBLSchBar)
                            QtyProduce = ObjCD.RoundDown1((rowSchItem.plan_prod_rate * WorkDuration) / 60, rowSchItem._Decimal)
                            'If QtyProduce < vQtySisaProduksi Then
                            '    QtyProduce = vQtySisaProduksi
                            'End If
                            If QtyProduce > 0 Then
                                If vQtySisaProduksi >= QtyProduce Then
                                    WorkDuration = ObjCD.RoundUp1((QtyProduce * 60) / rowSchItem.plan_prod_rate, 0)
                                    vQtySisaProduksi -= QtyProduce
                                    rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                                    rowSchItem.qty_produce += QtyProduce
                                    rowSchItem.plan_work_duration += WorkDuration
                                Else
                                    QtyProduce = vQtySisaProduksi
                                    WorkDuration = ObjCD.RoundUp1((QtyProduce * 60) / rowSchItem.plan_prod_rate, 0)
                                    vQtySisaProduksi -= QtyProduce
                                    rowSchItem.plan_current_time = Me.GetEndTime(rowSchItem.primary_work_unit, rowSchItem.plan_current_time, WorkDuration, vTBLSchBar)
                                    rowSchItem.qty_produce += QtyProduce
                                    rowSchItem.plan_work_duration += WorkDuration
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'TBLSchItem_Dummy.Dispose()
        End Try
    End Sub

    'mengembalikan free time + ot free terakhir yang sudah dialokasi untuk work unit ybs
    Private Function GetEndFreeTime(ByVal vPrimaryWorkUnit As Double, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Date
        Dim RetVal As Date = Date.MinValue
        Dim rowSch As DseTrans_Schedule.tbl_sch_barRow

        Try
            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                rowSch = DV.Row
                If rowSch.plan_end_time > RetVal Then
                    RetVal = rowSch.plan_end_time
                End If
            Next

            'kalo sampe sini, maka tidak cukup sch bar yang butuh di alocate
            If RetVal <> Date.MinValue Then
                Return RetVal
            Else
                Return #1/1/2200#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetStartTime(ByVal vPrimaryWorkUnit As Double, ByVal vEndDate As Date, ByVal vWorkDuration As Decimal, _
   ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Date
        Dim RetVal As Date = #1/1/2200#
        Dim rowSch, rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim CurrentTime As Date = vEndDate
        Dim WD As Double = 0
        Dim TotalWD As Double = 0

        Try
            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time]"

            For ijk As Integer = vTBLSchBar.DefaultView.Count - 1 To 0 Step -1
                rowSch = vTBLSchBar.DefaultView(ijk).Row
                WD = 0

                If rowSch.plan_start_time <= CurrentTime And rowSch.plan_end_time >= CurrentTime Then
                    WD = (CurrentTime - rowSch.plan_start_time).TotalMinutes
                End If
                If vWorkDuration <= TotalWD + WD Then
                    WD = vWorkDuration - TotalWD
                    If WD >= 0 Then
                        Return CurrentTime - TimeSpan.FromMinutes(WD)
                    Else
                        Return CurrentTime + TimeSpan.FromMinutes(WD * (-1))
                    End If
                Else
                    TotalWD += WD
                    If (ijk - 1) >= 0 Then
                        rowSch1 = vTBLSchBar.DefaultView(ijk - 1).Row
                        If WD > 0 Then
                            CurrentTime = rowSch1.plan_end_time
                        ElseIf rowSch.plan_start_time = CurrentTime Then
                            CurrentTime = rowSch1.plan_end_time
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next

            'kalo sampe sini, maka tidak cukup sch bar yang sudah di alocate
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'mengembalikan waktu selesai produksi, sesuai parameter yang tersedia
    Public Function GetEndTime(ByVal vPrimaryWorkUnit As Double, ByVal vStartDate As Date, ByVal vWorkDuration As Decimal, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Date
        Dim RetVal As Date = #1/1/2200#
        Dim rowSch, rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim CurrentTime As Date = vStartDate

        Try
            ''jika start date = start date off time, karena pengerjaan pas abis di akhir shift, maka start date = start date pertama setelah off time tsb
            ''maka cari start time pertama yang free time atau free ot
            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.scheduled_dt).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString + " AND [plan start time] = #" + Format(vStartDate, "dd MMM yyyy HH:mm") + "#"
            While vTBLSchBar.DefaultView.Count > 0
                rowSch = vTBLSchBar.DefaultView(0).Row
                vStartDate = rowSch.plan_end_time
                CurrentTime = vStartDate
                vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time).ToString + _
                " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.scheduled_dt).ToString + _
                ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString + " AND [plan start time] = #" + Format(vStartDate, "dd MMM yyyy HH:mm") + "#"
                vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"
            End While

            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

            For ijk As Integer = 0 To vTBLSchBar.DefaultView.Count - 1
                rowSch = vTBLSchBar.DefaultView(ijk).Row

                'If ijk = 0 And rowSch.plan_start_time > CurrentTime Then -> bagian ini yg di ganti dari fungsi GetEndTime yg original
                '    CurrentTime = rowSch.plan_start_time
                'End If
                If rowSch.plan_start_time > CurrentTime Then
                    CurrentTime = rowSch.plan_start_time
                End If

                If rowSch.plan_start_time <= CurrentTime And rowSch.plan_end_time >= CurrentTime Then
                    If (rowSch.plan_end_time - CurrentTime).TotalMinutes < vWorkDuration Then
                        If ijk = vTBLSchBar.DefaultView.Count - 1 Then
                            'jika baris terakhir maka, return #1/1/2200#
                            'artinya tidak cukup durasinya
                            Return #1/1/2200#
                        Else
                            'kurangi work duration, sejumlah schedule bar tsb, start time adalah start time schedule bar selanjutnya
                            vWorkDuration -= (rowSch.plan_end_time - CurrentTime).TotalMinutes
                            rowSch1 = vTBLSchBar.DefaultView(ijk + 1).Row
                            CurrentTime = rowSch1.plan_start_time
                        End If
                    Else
                        'Return rowSch.plan_start_time + TimeSpan.FromMinutes((CurrentTime - rowSch.plan_start_time).TotalMinutes)
                        Return CurrentTime + TimeSpan.FromMinutes(vWorkDuration)
                    End If
                End If
            Next

            'kalo sampe sini, maka tidak cukup sch bar yang butuh di alocate
            Return #1/1/2200#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetWorkDuration(ByVal vProdRate As Decimal, ByVal vProdQty As Decimal) As Double
        Try
            Return ObjCD.RoundUp1((vProdQty * 60) / vProdRate, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetWorkDuration(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryWorkUnit As Double, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Double
        Dim RetVal As Double = 0
        Dim rowSch As DseTrans_Schedule.tbl_sch_barRow
        Dim CurrentTime As Date = vStartDate

        Try
            'jika start date = start date off time, karena pengerjaan pas abis di akhir shift, maka start date = start date pertama setelah off time tsb
            'maka cari start time pertama yang free time atau free ot
            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.scheduled_dt).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString + " AND [plan start time] = #" + Format(vStartDate, "dd MMM yyyy HH:mm") + "#"
            While vTBLSchBar.DefaultView.Count > 0
                rowSch = vTBLSchBar.DefaultView(0).Row
                vStartDate = rowSch.plan_end_time
                CurrentTime = vStartDate
                vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time).ToString + _
                " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.scheduled_dt).ToString + _
                ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString + " AND [plan start time] = #" + Format(vStartDate, "dd MMM yyyy HH:mm") + "#"
                vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"
            End While

            vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
            " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
            ") AND [primary work unit] = " + vPrimaryWorkUnit.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"
            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                rowSch = DV.Row
                If rowSch.plan_start_time <= CurrentTime And rowSch.plan_end_time >= CurrentTime Then
                    If rowSch.plan_end_time < vEndDate Then
                        RetVal += (rowSch.plan_end_time - CurrentTime).TotalMinutes
                        CurrentTime = rowSch.plan_end_time
                    Else
                        RetVal += (vEndDate - CurrentTime).TotalMinutes
                        Exit For
                    End If
                ElseIf rowSch.plan_start_time >= CurrentTime And rowSch.plan_end_time >= CurrentTime Then
                    If rowSch.plan_start_time >= vEndDate And rowSch.plan_end_time >= vEndDate Then
                        Exit For
                    Else
                        If rowSch.plan_end_time <= vEndDate Then
                            'ambil keseluruhan item sch
                            RetVal += (rowSch.plan_end_time - rowSch.plan_start_time).TotalMinutes
                            CurrentTime = rowSch.plan_end_time
                        ElseIf rowSch.plan_end_time > vEndDate And rowSch.plan_start_time <= vEndDate Then
                            'ambil sampai end date, lalu exit fungsi, karena sudah sampai end date
                            RetVal += (vEndDate - rowSch.plan_start_time).TotalMinutes
                            Return RetVal
                        End If
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AllocateSchedule(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
    ByVal vrowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal vPrimaryWUConfig As Double)
        Dim rowSch, rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            For Each rowSchItem In vTBLSchItem.Rows
                If rowSchItem.qty_produce + rowSchItem.qty_produce_add > 0 Then
                    'buang free time
                    vTBLSchBar.DefaultView.RowFilter = "(tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString + _
                    " OR tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot).ToString + _
                   ") AND [primary work unit] = " + rowSchItem.primary_work_unit.ToString
                    vTBLSchBar.DefaultView.Sort = "[plan start time] ASC"

                    'pastikan end time pada baris terakhir dari sch bar, mencakup end time dari jadwal yang mau dialokasikan
                    'rowSch = vTBLSchBar.DefaultView(vTBLSchBar.DefaultView.Count - 1).Row
                    'If rowSch.plan_end_time < rowSchItem.plan_end_time Then
                    'modifikasi sch bar free time dan free ot
                    For ij As Double = 0 To vTBLSchBar.DefaultView.Count - 1
                        rowSch = vTBLSchBar.DefaultView(ij).Row

                        If rowSch.plan_start_time <= rowSchItem.plan_start_time And rowSch.plan_end_time >= rowSchItem.plan_end_time Then
                            'include (full atau partial)
                            If rowSch.plan_start_time = rowSchItem.plan_start_time And rowSch.plan_end_time > rowSchItem.plan_end_time Then
                                '1. kiri
                                'geser start time
                                rowSch.plan_start_time = rowSchItem.plan_end_time
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                            ElseIf rowSch.plan_start_time < rowSchItem.plan_start_time And rowSch.plan_end_time = rowSchItem.plan_end_time Then
                                '2. kanan
                                'geser end time
                                rowSch.plan_end_time = rowSchItem.plan_start_time
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                            ElseIf rowSch.plan_start_time = rowSchItem.plan_start_time And rowSch.plan_end_time = rowSchItem.plan_end_time Then
                                '3. full
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                                'buang sch bar
                                vTBLSchBar.Rows.Remove(rowSch)
                                ij -= 1
                            ElseIf rowSch.plan_start_time < rowSchItem.plan_start_time And rowSch.plan_end_time > rowSchItem.plan_end_time Then
                                '4. tengah
                                'tambahkan sisa bagian belakang dari free time ot free ot sch bar
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(rowSch.a, rowSch.r, rowSch.g, rowSch.b), rowSchItem.plan_end_time, rowSch.plan_end_time, "", "", _
                                -1, -1, rowSch.primary_work_unit, 0, _
                                0, 0, 0, 0, 0, -1, -1, -1, -1)
                                'geser end time
                                rowSch.plan_end_time = rowSchItem.plan_start_time
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                            End If
                        ElseIf rowSch.plan_start_time <= rowSchItem.plan_start_time And rowSch.plan_end_time >= rowSchItem.plan_start_time Then
                            'nyebrang, cari semua sch bar yang harus dibuang
                            'berhenti jika sch bar terakhir atau plan end time sch item berada dalam range sch bar
                            While ij <= vTBLSchBar.DefaultView.Count - 2
                                rowSch1 = vTBLSchBar.DefaultView(ij + 1).Row
                                If rowSch1.plan_end_time < rowSchItem.plan_end_time Then
                                    'full
                                    If rowSch1.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                        'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                        Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSch1.plan_start_time, rowSch1.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                        vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                    End If
                                    'buang sch bar
                                    vTBLSchBar.Rows.Remove(rowSch1)
                                    'tidak perlu modifikasi ij, karena rowsch1 dibuang -> lanjut evaluasi sch bar selanjutny (bukan rowsch yang dibuang)
                                ElseIf rowSch1.plan_end_time = rowSchItem.plan_end_time Then
                                    'full
                                    If rowSch1.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                        'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                        Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSch1.plan_start_time, rowSch1.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                        vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                    End If
                                    'buang sch bar
                                    vTBLSchBar.Rows.Remove(rowSch1)
                                    'tidak perlu modifikasi ij, karena rowsch1 dibuang -> lanjut evaluasi sch bar selanjutny (bukan rowsch yang dibuang)

                                    'exit while, karena ini adalah sch bar1 terakhir
                                    Exit While
                                ElseIf rowSch1.plan_end_time > rowSchItem.plan_end_time Then
                                    'kiri
                                    If rowSch1.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                        'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                        Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSch1.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                        vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                    End If
                                    'geser start time
                                    rowSch1.plan_start_time = rowSchItem.plan_end_time

                                    'exit while, karena ini adalah sch bar1 terakhir
                                    Exit While
                                End If
                            End While

                            'modifikasi current sch bar
                            If rowSch.plan_start_time < rowSchItem.plan_start_time Then
                                'kanan
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSch.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                                'geser end time
                                rowSch.plan_end_time = rowSchItem.plan_start_time
                            ElseIf rowSch.plan_start_time = rowSchItem.plan_start_time Then
                                'full
                                If rowSch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                    'tandai bahwa itu adalah overtime, dengan menambahkan sch bar item work time ot
                                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSch.plan_start_time, rowSch.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time_ot, 0, 0, 0, 0, 0, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                                End If
                                'buang sch bar
                                vTBLSchBar.Rows.Remove(rowSch)
                                ij -= 1
                            End If
                        ElseIf rowSch.plan_start_time > rowSchItem.plan_start_time And rowSch.plan_start_time >= rowSchItem.plan_end_time Then
                            'udah ewat
                            Exit For
                        End If
                    Next

                    'tambahkan work time
                    Me.AddSchBar(vTBLSchBar, Color.FromArgb(a, r, g, b), rowSchItem.plan_start_time, rowSchItem.plan_end_time, vrowItemWO.Identifier_JO, vrowItemWO.Identifier, _
                    vrowItemWO.Primary_MO, vrowItemWO.Primary_Process, rowSchItem.primary_work_unit, EnumTipeSchBar.work_time, rowSchItem.qty_produce, rowSchItem.plan_prod_rate, _
                    rowSchItem.plan_work_duration, rowSchItem.progress, rowSchItem._Decimal, vrowItemWO.Primary_Unit, vPrimaryWUConfig, rowSchItem.primary_input_location, rowSchItem.primary_output_location)
                    'Else
                    '    rowSchItem.plan_end_time = #1/1/2200#
                    'End If
                End If

                'If rowSchItem.primary_work_unit = 5 Then
                '    vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + rowSchItem.primary_work_unit.ToString
                'End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function SimpanJadwalBaru(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, _
   ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
   ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer, ByVal vPrimaryWUConfig As Double) As Boolean
        '1. update item jo -> set color, a,r,g,b
        '2. update semua item wo -> set start time, end time, setup time, work time, delay time, 
        '   total time, allocated wu
        '3. hapus item sch bar berdasartkan identifier jo, insert semua item wo item sch
        Dim DTAJO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item1TableAdapter
        Dim DTAWO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim DTAWO_ItemSCH As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter

        Try
            DTAJO.Connection.ConnectionString = Core_Module.GetConstring
            DTAWO.Connection.ConnectionString = DTAJO.Connection.ConnectionString
            DTAWO_ItemSCH.Connection.ConnectionString = DTAJO.Connection.ConnectionString

            '1. update item jo -> set sudah terjadwal, set color, a,r,g,b
            DTAJO.UpdateQuery(va, vr, vg, vb, vPrimaryWUConfig, vIdentifierJO)

            '2. update semua item wo -> set start time, end time, setup time, work time, delay time, 
            '   plan rate, effectiveness, normal rate, end time, qty produce, allocated wu
            For Each rowItemWO In vTBLItemWO.Rows
                With rowItemWO
                    .Plan_Total_Time = .Plan_Work_Hour + .Plan_Setup_Time + .Plan_Delay_Time
                    DTAWO.UpdateQuery(.Plan_Start_Time, .Plan_End_Time, .Plan_Setup_Time, _
                    .Plan_Work_Hour, .Plan_Delay_Time, .Plan_Total_Time, .Plan_Prod_Rate, .WU_Allocated, CInt(ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled), _
                    .tipe_dependency, .delay_dependency, .Primary)
                    'plan qty produce, plan rate, qty result plan, plan GP
                    '.Qty_Target_Prod = (.Qty_Target_Result * 100) / .Plan_Good_Perct
                End With
            Next

            '3. hapus item sch bar berdasartkan identifier jo, insert semua item wo item sch
            DTAWO_ItemSCH.DeleteQuery_ByIdentifierJO(vIdentifierJO)
            Me.Simpan_Item_UsingScript(vTBLItemWO_ItemSch, va, vr, vg, vb)

            '4. simpan item rm detail dan wip detail
            Me.SimpanItemRM_WIP_Detail(vIdentifierJO)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAJO.Dispose()
            DTAWO.Dispose()
            DTAWO_ItemSCH.Dispose()
        End Try
    End Function

    Private Sub SimpanItemRM_WIP_Detail(ByVal vIdentifierJO As String)
        Dim DTAWIPDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wip_detailTableAdapter
        Dim TBLWIPDetail As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable
        Dim DTARMDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_detailTableAdapter
        Dim TBLRMDetail As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable
        Dim DTAItemWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo5TableAdapter
        Dim TBLItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo5DataTable
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo5Row
        Dim DTAItemWOSch As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBLItemWOSch As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo_item_schDataTable
        Dim rowItemWOSch As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo_item_schRow
        Dim DTAItemRM As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm1TableAdapter
        Dim TBLItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm1DataTable
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm1Row
        Dim DV1, DV2 As Data.DataRowView
        Dim JmlWU As Integer = 0
        Dim Idx As Integer = 0
        Dim JmlAlokasi As Decimal = 0

        Try
            DTAWIPDetail.Connection.ConnectionString = Core_Module.GetConstring
            DTARMDetail.Connection.ConnectionString = DTAWIPDetail.Connection.ConnectionString
            DTAItemWO.Connection.ConnectionString = DTAWIPDetail.Connection.ConnectionString
            DTAItemWOSch.Connection.ConnectionString = DTAWIPDetail.Connection.ConnectionString
            DTAItemRM.Connection.ConnectionString = DTAWIPDetail.Connection.ConnectionString

            'hapus item wip detail dan rm detail atas indeitifer jo ybs
            DTAWIPDetail.Delete(vIdentifierJO)
            DTARMDetail.DeleteQuery(vIdentifierJO)

            'kumpulkan semua baris item wo, item wo sch dan item rm, atas identifier jo ybs
            DTAItemWO.Fill(TBLItemWO, vIdentifierJO)
            DTAItemWOSch.Fill(TBLItemWOSch, vIdentifierJO)
            DTAItemRM.Fill(TBLItemRM, vIdentifierJO)
            For Each RowItemRM In TBLItemRM.Rows
                RowItemRM.Allocated_RM = 0
                RowItemRM.Allocated_WIP = 0
            Next
            'untuk setiap item wo, jika qty target produce > 0
            'bagi qty rm dan wip untuk wo tersebut sesuai dengan alokasi work unitnya, weighing average
            For Each rowItemWO In TBLItemWO.Rows
                If rowItemWO.Balance_Produce > 0 Then
                    TBLItemWOSch.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier_WO + "'"
                    TBLItemRM.DefaultView.RowFilter = "[Identifier WO] = '" + rowItemWO.Identifier_WO + "'"

                    Idx = 0
                    JmlWU = TBLItemWOSch.DefaultView.Count

                    For Each DV1 In TBLItemWOSch.DefaultView
                        Idx += 1
                        rowItemWOSch = DV1.Row
                        If Idx = JmlWU Then
                            'alokasikan semua rm ke dalam wu ybs
                            For Each DV2 In TBLItemRM.DefaultView
                                RowItemRM = DV2.Row
                                If RowItemRM.is_wip = 0 Then
                                    JmlAlokasi = RowItemRM.Required_Qty + RowItemRM.Additional_Qty - RowItemRM.Allocated_RM - RowItemRM.Replace_Qty
                                    If JmlAlokasi > 0 Then
                                        Me.TambahItemRMDetail(TBLRMDetail, vIdentifierJO, RowItemRM.Identifier_WO, rowItemWOSch.primary_work_unit, _
                                        JmlAlokasi, rowItemWOSch.primary_input_location, rowItemWOSch.primary_mo, RowItemRM.Primary_Product_Code_RM)
                                        RowItemRM.Allocated_RM += JmlAlokasi
                                    End If
                                Else
                                    'alokasikan sisa WIP
                                    If RowItemRM.Allocated_WIP < RowItemRM.WIP_Qty Then
                                        JmlAlokasi = RowItemRM.WIP_Qty - RowItemRM.Allocated_WIP
                                        Me.TambahWIPDetail(TBLWIPDetail, vIdentifierJO, RowItemRM.Identifier_WO, rowItemWOSch.primary_work_unit, _
                                        JmlAlokasi, rowItemWOSch.primary_input_location, rowItemWOSch.primary_mo, RowItemRM.Primary_Product_Code_RM)
                                        RowItemRM.Allocated_WIP += JmlAlokasi
                                    End If
                                    'alokasikan sisa RM
                                    'JmlAlokasi = RowItemRM.Required_Qty + RowItemRM.Additional_Qty - RowItemRM.Allocated_RM - RowItemRM.WIP_Qty
                                    'If JmlAlokasi > 0 Then
                                    '    Me.TambahItemRMDetail(TBLRMDetail, vIdentifierJO, RowItemRM.Identifier_WO, rowItemWOSch.primary_work_unit, _
                                    '    JmlAlokasi, rowItemWOSch.primary_input_location, rowItemWOSch.primary_mo, RowItemRM.Primary_Product_Code_RM)
                                    '    RowItemRM.Allocated_RM += JmlAlokasi
                                    'End If
                                End If
                            Next
                        Else
                            'alokasikan berdasarkan weighing average
                            For Each DV2 In TBLItemRM.DefaultView
                                RowItemRM = DV2.Row
                                If RowItemRM.is_wip = 0 Then
                                    JmlAlokasi = Math.Round((RowItemRM.Required_Qty + RowItemRM.Additional_Qty - RowItemRM.Replace_Qty) * _
                                    (rowItemWOSch.qty_produce + rowItemWOSch.qty_produce_add) / rowItemWO.Balance_Produce, RowItemRM._Decimal, MidpointRounding.AwayFromZero)
                                    If JmlAlokasi > 0 Then
                                        Me.TambahItemRMDetail(TBLRMDetail, vIdentifierJO, RowItemRM.Identifier_WO, rowItemWOSch.primary_work_unit, _
                                        JmlAlokasi, rowItemWOSch.primary_input_location, rowItemWOSch.primary_mo, RowItemRM.Primary_Product_Code_RM)
                                        RowItemRM.Allocated_RM += JmlAlokasi
                                    End If
                                Else
                                    If RowItemRM.Allocated_WIP < RowItemRM.WIP_Qty Then
                                        JmlAlokasi = Math.Round(RowItemRM.WIP_Qty * _
                                        (rowItemWOSch.qty_produce + rowItemWOSch.qty_produce_add) / rowItemWO.Balance_Produce, RowItemRM._Decimal, MidpointRounding.AwayFromZero)
                                        If JmlAlokasi > 0 Then
                                            Me.TambahWIPDetail(TBLWIPDetail, vIdentifierJO, RowItemRM.Identifier_WO, rowItemWOSch.primary_work_unit, _
                                            JmlAlokasi, rowItemWOSch.primary_input_location, rowItemWOSch.primary_mo, RowItemRM.Primary_Product_Code_RM)
                                            RowItemRM.Allocated_WIP += JmlAlokasi
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            'simpan item rm detail dan wip detail ke dalam database 
            Me.Simpan_ItemRMDetail_UsingScript(TBLRMDetail)
            Me.Simpan_ItemWIPDetail_UsingScript(TBLWIPDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_ItemRMDetail_UsingScript(ByVal vTBL As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemRMDetail(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function Simpan_ItemWIPDetail_UsingScript(ByVal vTBL As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemWIPDetail(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemRMDetail(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable) As String
        Dim retval As String = ""
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_mo_item_rm_detail` (`Primary MO`,`Identifier JO`,`Identifier WO`,`Primary WU`,`Primary Product Code RM`,`Target Qty`," + _
            "`Used Qty`,`Balance Qty`,`Remarks`,`Primary Input Location`,`is wip`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(.Primary_MO).ToString + ",'" _
                        + .Identifier_JO + "'" + ",'" _
                        + .Identifier_WO + "'," _
                        + CDbl(.Primary_WU).ToString + "," _
                        + CDbl(.Primary_Product_Code_RM).ToString + "," _
                        + CDec(.Target_Qty).ToString + "," _
                        + CDec(.Used_Qty).ToString + "," _
                        + CDec(.Balance_Qty).ToString + ",'" _
                        + .Remarks + "'" + "," _
                        + CDbl(.Primary_Input_Location).ToString + "," _
                        + CDbl(.is_wip).ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next

            If retval.EndsWith(",") Then
                retval = retval.Substring(0, retval.Length - 1) + ";"
            End If

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemWIPDetail(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable) As String
        Dim retval As String = ""
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_mo_item_wip_detail` (`Primary MO`,`Primary WU`,`Identifier JO`,`Identifier WO`,`Primary Product Code WIP`,`Qty`," + _
            "`Used Qty`,`Balance Qty`,`Primary Input Location`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(.Primary_MO).ToString + "," _
                        + CDbl(.Primary_WU).ToString + ",'" _
                        + .Identifier_JO + "'" + ",'" _
                        + .Identifier_WO + "'," _
                        + CDbl(.Primary_Product_Code_WIP).ToString + "," _
                        + CDec(.Qty).ToString + "," _
                        + CDec(.Used_Qty).ToString + "," _
                        + CDec(.Balance_Qty).ToString + "," _
                        + CDbl(.Primary_Input_Location).ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next

            If retval.EndsWith(",") Then
                retval = retval.Substring(0, retval.Length - 1) + ";"
            End If

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub TambahWIPDetail(ByVal TBLWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable, _
    ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWU As Double, ByVal vQty As Decimal, _
    ByVal vPrimaryInputLocation As Double, ByVal vPrimaryMO As Double, ByVal vPrimaryWIP As Double)
        Dim rowWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailRow
        Try
            rowWIPDetail = TBLWIPDetail.Newtbltrans_mo_item_wip_detailRow
            With rowWIPDetail
                .Balance_Qty = vQty
                .Identifier_JO = vIdentifierJO
                .Identifier_WO = vIdentifierWO
                .Primary = -1
                .Primary_Input_Location = vPrimaryInputLocation
                .Primary_MO = vPrimaryMO
                .Primary_Product_Code_WIP = vPrimaryWIP
                .Primary_WU = vPrimaryWU
                .Qty = vQty
                .Used_Qty = 0
            End With
            TBLWIPDetail.Rows.Add(rowWIPDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahItemRMDetail(ByVal TBLRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable, _
    ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWU As Double, ByVal vQty As Decimal, _
    ByVal vPrimaryInputLocation As Double, ByVal vPrimaryMO As Double, ByVal vPrimaryRM As Double)
        Dim rowRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailRow
        Try
            rowRMDetail = TBLRMDetail.Newtbltrans_mo_item_rm_detailRow()
            With rowRMDetail

                .Balance_Qty = vQty
                .Identifier_JO = vIdentifierJO
                .Identifier_WO = vIdentifierWO
                .is_wip = 0
                .Primary = -1
                .Primary_Input_Location = vPrimaryInputLocation
                .Primary_MO = vPrimaryMO
                .Primary_Product_Code_RM = vPrimaryRM
                .Primary_WU = vPrimaryWU
                .Remarks = ""
                .Target_Qty = vQty
                .Used_Qty = 0
            End With
            TBLRMDetail.Rows.Add(rowRMDetail)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_Item_UsingScript(ByVal vTBL As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vTBL, va, vr, vg, vb)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
   ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer) As String
        Dim retval As String = ""
        Dim row As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_mo_item_wo_item_sch` (`primary mo`,`identifier jo`,`identifier wo`,`primary work unit`,`plan start time`,`plan end time`," + _
            "`plan end time minimum`,`plan setup duration`,`plan work duration`,`plan total duration`,`plan prod rate normal`,`plan prod rate efficiency`,`plan prod rate`," + _
            "`qty produce`,`a`,`r`,`g`,`b`,`primary process`,`plan good percentage`,`plan result`,`plan downtime`,`balance produce`,`balance result`,`status`,`tipe dependency`,`delay dependency`,`primary input location`,`primary output location`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'If row.qty_produce + row.qty_produce_add > 0 Then
                    'hanya simpan yang ada qty producenya saja
                    If row.qty_produce + row.qty_produce_add > 0 Then
                        'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                        If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                            Delimiter = ";"
                        End If

                        With row
                            .plan_total_duration = .plan_work_duration + .plan_work_duration_add + .plan_setup_duration + .plan_downtime

                            retval = retval + vbCrLf + "(" + CDbl(.primary_mo).ToString + ",'" + .identifier_jo + "'" + ",'" + .identifier_wo + "'," _
                            + CDbl(.primary_work_unit).ToString + ",'" + Format(CDate(.plan_start_time), "yyyy-MM-dd HH:mm:ss") + "'," _
                            + "'" + Format(CDate(.plan_end_time), "yyyy-MM-dd HH:mm:ss") + "'," + "'" + Format(CDate(.plan_end_time_minimum), "yyyy-MM-dd HH:mm:ss") + "'," _
                            + CDbl(.plan_setup_duration).ToString + "," + CDbl(.plan_work_duration).ToString + "," _
                            + CDbl(.plan_total_duration).ToString + "," + CDec(.plan_prod_rate_normal).ToString + "," _
                            + CDbl(.plan_prod_rate_efficiency).ToString + "," + CDec(.plan_prod_rate).ToString + "," _
                            + CDec(.qty_produce).ToString + "," + CInt(va).ToString + "," _
                            + CInt(vr).ToString + "," + CInt(vg).ToString + "," _
                            + CInt(vb).ToString + "," + CDbl(.primary_process).ToString + "," _
                            + CDec(.plan_good_percentage).ToString + "," _
                            + CDec(.plan_result).ToString + "," _
                            + CDec(.plan_downtime).ToString + "," _
                            + CDec(.qty_produce).ToString + "," _
                            + CDec(.plan_result).ToString + "," _
                            + CDbl(ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled).ToString + "," _
                            + CInt(.tipe_dependency).ToString + "," _
                            + CDbl(.delay_dependency).ToString + "," _
                            + CDbl(.primary_input_location).ToString + "," _
                            + CDbl(.primary_output_location).ToString + "" _
                            + ")" + Delimiter

                            'plan gp, plan qty result, plan downtime
                            '.plan_result = (.plan_good_percentage * .qty_produce) / 100
                        End With
                    End If
                Else
                    Exit For
                End If
            Next

            If retval.EndsWith(",") Then
                retval = retval.Substring(0, retval.Length - 1) + ";"
            End If

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Sub LoadScheduleBar(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable)
        Dim DTA_WS_Item As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL_WS_Item As New DseTrans_Schedule.tblmaster_workshift_itemDataTable
        Dim DTA_WS_ItemHoliday As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL_WS_ItemHoliday As New DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable

        Try
            DTA_WS_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WS_ItemHoliday.Connection.ConnectionString = DTA_WS_Item.Connection.ConnectionString

            DTA_WS_Item.Fill(TBL_WS_Item)
            DTA_WS_ItemHoliday.Fill(TBL_WS_ItemHoliday, vStartDate, vEndDate)

            'kosongkan tabel
            vTBLSchBar.Rows.Clear()

            'isikan data
            '1. allocate workshift yang ada 
            Me.LoadSchBar_AllocatedWS(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_AllocatedWS, TBL_WS_Item, TBL_WS_ItemHoliday, vTBLWU)
            '2. scheduled down time
            Me.LoadSchBar_SchDT(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_SchDT, vTBLWU)
            '3. scheduled overtime
            Me.LoadSchBar_SchOT(vStartDate, vTBLSchBar, Me.GetColor_SchOT, TBL_WS_Item, TBL_WS_ItemHoliday, vTBLWU)
            '4. item work order yang sudah terjadwal
            Me.LoadSchBar_ItemWO(vStartDate, vEndDate, vTBLSchBar, TBL_WS_Item, TBL_WS_ItemHoliday, vTBLWU)
            '5. off time untuk mengisikan data antaranya
            Me.LoadSchBar_Offtime(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_Offtime, TBL_WS_Item, TBL_WS_ItemHoliday)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Item.Dispose()
            TBL_WS_Item.Dispose()
            DTA_WS_ItemHoliday.Dispose()
            TBL_WS_ItemHoliday.Dispose()
        End Try
    End Sub

    'mengisikan tabel schedule bar, tanpa item yang sudah terjadwal
    Public Sub LoadScheduleBar_TanpaScheduledItem(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim DTA_WS_Item As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL_WS_Item As New DseTrans_Schedule.tblmaster_workshift_itemDataTable
        Dim DTA_WS_ItemHoliday As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL_WS_ItemHoliday As New DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable

        Try
            DTA_WS_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WS_ItemHoliday.Connection.ConnectionString = DTA_WS_Item.Connection.ConnectionString

            DTA_WS_Item.Fill(TBL_WS_Item)
            DTA_WS_ItemHoliday.Fill(TBL_WS_ItemHoliday, vStartDate, vEndDate)

            'kosongkan tabel
            vTBLSchBar.Rows.Clear()

            'isikan data
            '1. allocate workshift yang ada 
            Me.LoadSchBar_AllocatedWS(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_AllocatedWS, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing)
            '2. scheduled down time
            Me.LoadSchBar_SchDT(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_SchDT, Nothing)
            '3. scheduled overtime
            Me.LoadSchBar_SchOT(vStartDate, vTBLSchBar, Me.GetColor_SchOT, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing)
            '4. off time untuk mengisikan data antaranya
            Me.LoadSchBar_Offtime(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_Offtime, TBL_WS_Item, TBL_WS_ItemHoliday)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Item.Dispose()
            TBL_WS_Item.Dispose()
            DTA_WS_ItemHoliday.Dispose()
            TBL_WS_ItemHoliday.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' klo yang ini tanpa off time
    ''' </summary>
    ''' <param name="vStartDate"></param>
    ''' <param name="vEndDate"></param>
    ''' <param name="vTBLSchBar"></param>
    ''' <remarks></remarks>
    Public Sub LoadScheduleBar_TanpaScheduledItem2(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim DTA_WS_Item As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL_WS_Item As New DseTrans_Schedule.tblmaster_workshift_itemDataTable
        Dim DTA_WS_ItemHoliday As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL_WS_ItemHoliday As New DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable

        Try
            DTA_WS_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WS_ItemHoliday.Connection.ConnectionString = DTA_WS_Item.Connection.ConnectionString

            DTA_WS_Item.Fill(TBL_WS_Item)
            DTA_WS_ItemHoliday.Fill(TBL_WS_ItemHoliday, vStartDate, vEndDate)

            'kosongkan tabel
            vTBLSchBar.Rows.Clear()

            'isikan data
            '1. allocate workshift yang ada 
            Me.LoadSchBar_AllocatedWS(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_AllocatedWS, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing)
            '2. scheduled down time
            Me.LoadSchBar_SchDT(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_SchDT, Nothing)
            '3. scheduled overtime
            Me.LoadSchBar_SchOT(vStartDate, vTBLSchBar, Me.GetColor_SchOT, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Item.Dispose()
            TBL_WS_Item.Dispose()
            DTA_WS_ItemHoliday.Dispose()
            TBL_WS_ItemHoliday.Dispose()
        End Try
    End Sub

    Public Sub LoadScheduleBar_WOItemOnly(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim DTA_WO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL_WO As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim row_WO As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim TipeBar As EnumTipeSchBar = EnumTipeSchBar.work_time

        Try
            vTBLSchBar.Rows.Clear()

            DTA_WO.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WO.Fill(TBL_WO, vStartDate)

            For Each row_WO In TBL_WO.Rows
                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, row_WO.plan_prod_rate, _
                row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location)

                row_WO.inserted = 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WO.Dispose()
            TBL_WO.Dispose()
        End Try
    End Sub

    Public Sub LoadScheduleBar_WOItemOnly_EndDateMoreThan(ByVal vDateReference As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim DTA_WO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL_WO As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim row_WO As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim TipeBar As EnumTipeSchBar = EnumTipeSchBar.work_time

        Try
            vTBLSchBar.Rows.Clear()

            DTA_WO.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WO.FillBy_EndDateMoreThan(TBL_WO, vDateReference)

            For Each row_WO In TBL_WO.Rows
                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, row_WO.plan_prod_rate, _
                row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location)

                row_WO.inserted = 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WO.Dispose()
            TBL_WO.Dispose()
        End Try
    End Sub

    Public Sub LoadScheduleBar_TanpaScheduledItem(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vPrimaryWU As Double)
        Dim DTA_WS_Item As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL_WS_Item As New DseTrans_Schedule.tblmaster_workshift_itemDataTable
        Dim DTA_WS_ItemHoliday As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL_WS_ItemHoliday As New DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable

        Try
            DTA_WS_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WS_ItemHoliday.Connection.ConnectionString = DTA_WS_Item.Connection.ConnectionString

            DTA_WS_Item.Fill(TBL_WS_Item)
            DTA_WS_ItemHoliday.Fill(TBL_WS_ItemHoliday, vStartDate, vEndDate)

            'kosongkan tabel
            vTBLSchBar.Rows.Clear()

            'isikan data
            '1. allocate workshift yang ada 
            Me.LoadSchBar_AllocatedWS(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_AllocatedWS, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing, vPrimaryWU)
            '2. scheduled down time
            Me.LoadSchBar_SchDT(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_SchDT, Nothing, vPrimaryWU)
            '3. scheduled overtime
            Me.LoadSchBar_SchOT(vStartDate, vTBLSchBar, Me.GetColor_SchOT, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing, vPrimaryWU)
            '4. off time untuk mengisikan data antaranya
            Me.LoadSchBar_Offtime(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_Offtime, TBL_WS_Item, TBL_WS_ItemHoliday)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Item.Dispose()
            TBL_WS_Item.Dispose()
            DTA_WS_ItemHoliday.Dispose()
            TBL_WS_ItemHoliday.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' yang ini tanpa off time
    ''' </summary>
    ''' <param name="vStartDate"></param>
    ''' <param name="vEndDate"></param>
    ''' <param name="vTBLSchBar"></param>
    ''' <param name="vPrimaryWU"></param>
    ''' <remarks></remarks>
    Public Sub LoadScheduleBar_TanpaScheduledItem2(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vPrimaryWU As Double)
        Dim DTA_WS_Item As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL_WS_Item As New DseTrans_Schedule.tblmaster_workshift_itemDataTable
        Dim DTA_WS_ItemHoliday As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL_WS_ItemHoliday As New DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable

        Try
            DTA_WS_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WS_ItemHoliday.Connection.ConnectionString = DTA_WS_Item.Connection.ConnectionString

            DTA_WS_Item.Fill(TBL_WS_Item)
            DTA_WS_ItemHoliday.Fill(TBL_WS_ItemHoliday, vStartDate, vEndDate)

            'kosongkan tabel
            vTBLSchBar.Rows.Clear()

            'isikan data
            '1. allocate workshift yang ada 
            Me.LoadSchBar_AllocatedWS(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_AllocatedWS, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing, vPrimaryWU)
            '2. scheduled down time
            Me.LoadSchBar_SchDT(vStartDate, vEndDate, vTBLSchBar, Me.GetColor_SchDT, Nothing, vPrimaryWU)
            '3. scheduled overtime
            Me.LoadSchBar_SchOT(vStartDate, vTBLSchBar, Me.GetColor_SchOT, TBL_WS_Item, TBL_WS_ItemHoliday, Nothing, vPrimaryWU) 

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Item.Dispose()
            TBL_WS_Item.Dispose()
            DTA_WS_ItemHoliday.Dispose()
            TBL_WS_ItemHoliday.Dispose()
        End Try
    End Sub

    Private Sub LoadSchBar_Offtime(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vColor As Color, _
    ByVal TBL_WS_Item As DseTrans_Schedule.tblmaster_workshift_itemDataTable, ByVal TBL_WS_ItemHoliday As DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable)
        Dim ListPrimaryWorkUnit As New List(Of Double)
        Dim vRowSchBar, vRowSchBar_Next As DseTrans_Schedule.tbl_sch_barRow
        Dim DV As Data.DataRowView
        Dim PrimaryWorkUnit As Double = -1
        Dim Cur_Index As Integer = 0

        Try
            'scan untuk mendapatkan primary work unit yang ada
            vTBLSchBar.DefaultView.RowFilter = ""
            vTBLSchBar.DefaultView.Sort = "[primary work unit], [plan start time]"
            For Each DV In vTBLSchBar.DefaultView
                vRowSchBar = DV.Row
                If PrimaryWorkUnit <> vRowSchBar.primary_work_unit Then
                    ListPrimaryWorkUnit.Add(vRowSchBar.primary_work_unit)
                    PrimaryWorkUnit = vRowSchBar.primary_work_unit
                End If
            Next

            'mulai dari start time sampai dengan end time, isikan dengan sch bar yang bertipe off time, jika ada jeda
            For Each PrimaryWorkUnit In ListPrimaryWorkUnit
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWorkUnit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                Cur_Index = 0

                While Cur_Index <= vTBLSchBar.DefaultView.Count - 1
                    vRowSchBar = vTBLSchBar.DefaultView(Cur_Index).Row
                    If Cur_Index = 0 Then
                        If vTBLSchBar.DefaultView.Count = 1 Then
                            'isikan jeda before item sch
                            If vStartDate < vRowSchBar.plan_start_time Then
                                Me.AddSchBar(vTBLSchBar, vColor, vStartDate, vRowSchBar.plan_start_time, "", _
                                "", -1, -1, PrimaryWorkUnit, EnumTipeSchBar.off_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                            End If
                            'isikan jeda after item sch
                            If vEndDate > vRowSchBar.plan_end_time Then
                                Me.AddSchBar(vTBLSchBar, vColor, vRowSchBar.plan_end_time, vEndDate, "", _
                                "", -1, -1, PrimaryWorkUnit, EnumTipeSchBar.off_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                            End If
                            'beres, cek work unit selanjutnya
                            Exit While
                        Else
                            If vStartDate < vRowSchBar.plan_start_time Then
                                'jeda antara start date dan item sch bar pertama
                                Me.AddSchBar(vTBLSchBar, vColor, vStartDate, vRowSchBar.plan_start_time, "", _
                                "", -1, -1, PrimaryWorkUnit, EnumTipeSchBar.off_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                            End If
                        End If
                        Cur_Index += 1
                    ElseIf Cur_Index = vTBLSchBar.DefaultView.Count - 1 Then
                        'jeda antara item sch bar dan end date
                        If vEndDate > vRowSchBar.plan_end_time Then
                            Me.AddSchBar(vTBLSchBar, vColor, vRowSchBar.plan_end_time, vEndDate, "", _
                            "", -1, -1, PrimaryWorkUnit, EnumTipeSchBar.off_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                        End If
                        'beres, cek work unit selanjutnya
                        Exit While
                    Else
                        vRowSchBar_Next = vTBLSchBar.DefaultView(Cur_Index + 1).Row
                        'jeda antara item sch bar
                        If vRowSchBar.plan_end_time < vRowSchBar_Next.plan_start_time Then
                            Me.AddSchBar(vTBLSchBar, vColor, vRowSchBar.plan_end_time, vRowSchBar_Next.plan_start_time, "", _
                            "", -1, -1, PrimaryWorkUnit, EnumTipeSchBar.off_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                        End If
                        Cur_Index += 1
                    End If
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryWorkUnit = Nothing
        End Try
    End Sub

    Private Sub LoadSchBar_ItemWO(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal TBL_WS_Item As DseTrans_Schedule.tblmaster_workshift_itemDataTable, ByVal TBL_WS_ItemHoliday As DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable)
        Dim DTA_WO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL_WO As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim row_WO As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Cur_Index As Integer = 0
        Dim row_sch As DseTrans_Schedule.tbl_sch_barRow
        Dim DummyTime As Date = Now
        Dim TipeBar As EnumTipeSchBar = EnumTipeSchBar.work_time
        Dim TipeBar1 As EnumTipeSchBar = EnumTipeSchBar.work_time_ot
        Dim indexLeft As Integer = 0
        Dim indexRight As Integer = 0
        'Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunit1Row
        Dim ListPrimaryWU1 As New List(Of Double)
        Dim ListPrimaryWU2 As New List(Of Double)
        Dim PrimaryWU As Double = -2
        Dim DV As Data.DataRowView

        Try
            DTA_WO.Connection.ConnectionString = Core_Module.GetConstring
            DTA_WO.Fill(TBL_WO, vStartDate)

            If IsNothing(vTBLWU) = False Then
                If vTBLWU.Rows.Count > 0 Then
                    'buang item allocated ws, atas work unit yang tidak ada didalam list wu
                    'While indexRight <= TBL_WO.Rows.Count - 1
                    '    If indexLeft <= vTBLWU.Rows.Count - 1 Then
                    '        vrowWU = vTBLWU.Rows(indexLeft)
                    '        row_WO = TBL_WO.Rows(indexRight)
                    '        If vrowWU.primary_work_unit > row_WO.primary_work_unit Then
                    '            TBL_WO.Rows.RemoveAt(indexRight)
                    '        ElseIf vrowWU.primary_work_unit = row_WO.primary_work_unit Then
                    '            indexRight += 1
                    '        Else
                    '            indexLeft += 1
                    '        End If
                    '    Else
                    '        TBL_WO.Rows.RemoveAt(indexRight)
                    '    End If
                    'End While

                    TBL_WO.DefaultView.Sort = "[primary work unit]"
                    TBL_WO.DefaultView.RowFilter = ""
                    For Each DV In TBL_WO.DefaultView
                        row_WO = DV.Row
                        If row_WO.primary_work_unit <> PrimaryWU Then
                            PrimaryWU = row_WO.primary_work_unit
                            ListPrimaryWU1.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU1
                        vTBLWU.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        If vTBLWU.DefaultView.Count = 0 Then
                            'tidak ada di tabel wu, catat
                            ListPrimaryWU2.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU2
                        TBL_WO.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        While TBL_WO.DefaultView.Count > 0
                            TBL_WO.Rows.Remove(TBL_WO.DefaultView(0).Row)
                        End While
                    Next
                End If
            End If

            For Each row_WO In TBL_WO.Rows
                row_WO.inserted = 0
            Next
            For Each row_WO In TBL_WO.Rows
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + row_WO.primary_work_unit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                Cur_Index = 0

                While Cur_Index <= vTBLSchBar.DefaultView.Count - 1
                    row_sch = vTBLSchBar.DefaultView(Cur_Index).Row
                    'hanya bisa melakukan penjadwal pada item allocated ws atau sch ot
                    If row_sch.tipe = EnumTipeSchBar.free_ot Or _
                    row_sch.tipe = EnumTipeSchBar.free_time Then
                        'If row_sch.tipe = EnumTipeSchBar.free_ot Then
                        '    TipeBar = EnumTipeSchBar.work_time_ot
                        'ElseIf row_sch.tipe = EnumTipeSchBar.free_time Then
                        '    TipeBar = EnumTipeSchBar.work_time
                        'End If

                        If row_sch.plan_end_time <= row_WO.plan_start_time And row_sch.plan_start_time < row_WO.plan_start_time Then
                            'item jadwal/overtime tidak bersinggungan sama sekali dengan item ws
                            'majukan indexnya, untuk mengecek item sch selanjutnya
                            Cur_Index += 1
                        ElseIf row_sch.plan_start_time > row_WO.plan_start_time And row_sch.plan_start_time >= row_WO.plan_end_time Then
                            'item jadwal/overtime tidak bersinggungan sama sekali dengan item ws
                            'exit while karena tidak ada sch untuk work unit ini, semua item sch untuk work unit ybs sudah lewat
                            Exit While
                        ElseIf row_sch.plan_end_time >= row_WO.plan_start_time And row_sch.plan_end_time <= row_WO.plan_end_time And _
                        row_sch.plan_start_time < row_WO.plan_start_time Then
                            'item work order bersinggungan dengan item allocated ws/sch ot, pada bagian akhir
                            'untuk sch bar: geser end time item sch ke start time item wo
                            DummyTime = row_sch.plan_end_time
                            row_sch.plan_end_time = row_WO.plan_start_time
                            'tambahkan sch bar baru: start time = start time item wo; end time = endtime item sch bar
                            'tipe sch baru baru tergantung tipe sch bar asal
                            If row_WO.inserted = 0 Then
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, row_WO.plan_prod_rate, _
                                row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location) ', row_sch.primary_workshift)
                            End If

                            If row_sch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                'jika tipe-nya free ot, maka tambahkan bar, work ot
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, DummyTime, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar1, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_sch.primary_workshift)
                            End If
                            row_WO.inserted = 1
                            'index +1
                            Cur_Index += 1
                        ElseIf row_sch.plan_start_time >= row_WO.plan_start_time And row_sch.plan_start_time <= row_WO.plan_end_time And _
                        row_sch.plan_end_time > row_WO.plan_end_time Then
                            'item work order bersinggungan pada bagian depan
                            'untuk sch bar: geser start time ke end time untuk item wo
                            DummyTime = row_sch.plan_start_time
                            row_sch.plan_start_time = row_WO.plan_end_time
                            'tambahkan sch bar baru, start time = start time lama item sch bar, end time = end time item wo
                            'tipe sch bar tergantung item sch bar
                            If row_WO.inserted = 0 Then
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), DummyTime, row_WO.plan_end_time, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, _
                                row_WO.plan_prod_rate, row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location) ', row_sch.primary_workshift)
                            End If
                            If row_sch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), DummyTime, row_WO.plan_end_time, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar1, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                            End If
                            row_WO.inserted = 1
                            'index +1
                            Cur_Index += 1
                        ElseIf row_sch.plan_start_time >= row_WO.plan_start_time And row_sch.plan_end_time <= row_WO.plan_end_time Then
                            'item work order melingkupi item sch/sama besar
                            'ganti tipe sch bar, tergantung sch bar sebelumnya
                            If row_WO.inserted = 0 Then
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, row_WO.plan_prod_rate, _
                                row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location)
                                row_WO.inserted = 1
                                'index +=1
                                Cur_Index += 1
                            End If
                            If row_sch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                With row_sch
                                    .tipe = TipeBar1
                                    'lengkapi informasi primary mo, identifier jo, identifier wo, primary work unit, primary process
                                    .primary_mo = row_WO.primary_mo
                                    .identifier_jo = row_WO.identifier_jo
                                    .identifier_wo = row_WO.identifier_wo
                                    .primary_work_unit = row_WO.primary_work_unit
                                    .primary_process = row_WO.primary_process
                                    .a = row_WO.a
                                    .r = row_WO.r
                                    .g = row_WO.g
                                    .b = row_WO.b
                                    .qty_produce = 0
                                    .rate = 0
                                End With
                            ElseIf row_sch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_time Then
                                vTBLSchBar.Rows.Remove(row_sch)
                                'disini tidak nambah index, cek baris pengganti diposisi yang baru saja di remove
                            End If
                        ElseIf row_sch.plan_start_time < row_WO.plan_start_time And row_sch.plan_end_time > row_WO.plan_end_time Then
                            'item sch melingkupi item wo
                            'untuk sch bar: geser end time ke start time item wo
                            DummyTime = row_sch.plan_end_time
                            row_sch.plan_end_time = row_WO.plan_start_time

                            'masukkan item sch baru baru, start time dan end time = start time dan end time item wo
                            'tipe sch bar tergantung tipe sch bar, bisa wo biasa atau wo pada saat ot
                            Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                            row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar, row_WO.qty_produce, row_WO.plan_prod_rate, _
                            row_WO.plan_work_duration, row_WO.progress, row_WO._Decimal, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location) ', row_sch.primary_workshift)
                            If row_sch.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                                Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_WO.a, row_WO.r, row_WO.g, row_WO.b), row_WO.plan_start_time, row_WO.plan_end_time, row_WO.identifier_jo, _
                                row_WO.identifier_wo, row_WO.primary_mo, row_WO.primary_process, row_WO.primary_work_unit, TipeBar1, 0, 0, 0, 0, 0, -1, -1, -1, -1)
                            End If
                            'tambahkan sch baru baru, start time = end time item wo, end time = end time lama item sch bar
                            'tipe sch bar sama persis dengan tipe sch bar lama
                            Me.AddSchBar(vTBLSchBar, Color.FromArgb(row_sch.a, row_sch.r, row_sch.g, row_sch.b), row_WO.plan_end_time, DummyTime, row_sch.identifier_jo, _
                            row_sch.identifier_wo, row_sch.primary_mo, row_sch.primary_process, row_sch.primary_work_unit, row_sch.tipe, row_sch.qty_produce, row_sch.rate, _
                            row_sch.total_work_hour, row_sch.progress, row_sch.decimal_point, row_WO.Primary_Unit, row_WO.primary_wu_config, row_WO.primary_input_location, row_WO.primary_output_location) ', row_sch.primary_workshift)
                            row_WO.inserted = 1

                            'exit while, karena alokasi item wo ybs sudah tercapai semua, cek item wo selanjutnya
                            Exit While
                        End If
                    Else
                        'majukan indexnya, untuk mengecek item sch selanjutnya
                        Cur_Index += 1
                    End If
                End While

                row_WO.inserted = 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message + " -  cur Idx: " + Cur_Index.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WO.Dispose()
            TBL_WO.Dispose()
            ListPrimaryWU1 = Nothing
            ListPrimaryWU2 = Nothing
        End Try
    End Sub

    Private Sub LoadSchBar_SchOT(ByVal vStartDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vColor As Color, _
    ByVal TBL_WS_Item As DseTrans_Schedule.tblmaster_workshift_itemDataTable, ByVal TBL_WS_ItemHoliday As DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable, Optional ByVal vPrimaryWU As Double = -1)
        Dim DTA_OT As New DseTrans_ScheduleTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim TBL_OT As New DseTrans_Schedule.tbltrans_overtime_itemDataTable
        Dim row_OT As DseTrans_Schedule.tbltrans_overtime_itemRow
        Dim indexLeft As Integer = 0
        Dim indexRight As Integer = 0
        'Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunit1Row
        Dim ListPrimaryWU1 As New List(Of Double)
        Dim ListPrimaryWU2 As New List(Of Double)
        Dim PrimaryWU As Double = -2
        Dim DV As Data.DataRowView

        Try
            DTA_OT.Connection.ConnectionString = Core_Module.GetConstring

            If vPrimaryWU = -1 Then
                DTA_OT.Fill(TBL_OT, vStartDate)
            Else
                DTA_OT.FillBy_PrimaryWorkUnit(TBL_OT, vStartDate, vPrimaryWU)
            End If

            If IsNothing(vTBLWU) = False Then
                If vTBLWU.Rows.Count > 0 Then
                    'buang item allocated ws, atas work unit yang tidak ada didalam list wu
                    'While indexRight <= TBL_OT.Rows.Count - 1
                    '    If indexLeft <= vTBLWU.Rows.Count - 1 Then
                    '        vrowWU = vTBLWU.Rows(indexLeft)
                    '        row_OT = TBL_OT.Rows(indexRight)
                    '        If vrowWU.primary_work_unit > row_OT.primary_workunit Then
                    '            TBL_OT.Rows.RemoveAt(indexRight)
                    '        ElseIf vrowWU.primary_work_unit = row_OT.primary_workunit Then
                    '            indexRight += 1
                    '        Else
                    '            indexLeft += 1
                    '        End If
                    '    Else
                    '        TBL_OT.Rows.RemoveAt(indexRight)
                    '    End If
                    'End While

                    TBL_OT.DefaultView.Sort = "[primary workunit]"
                    TBL_OT.DefaultView.RowFilter = ""
                    For Each DV In TBL_OT.DefaultView
                        row_OT = DV.Row
                        If row_OT.primary_workunit <> PrimaryWU Then
                            PrimaryWU = row_OT.primary_workunit
                            ListPrimaryWU1.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU1
                        vTBLWU.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        If vTBLWU.DefaultView.Count = 0 Then
                            'tidak ada di tabel wu, catat
                            ListPrimaryWU2.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU2
                        TBL_OT.DefaultView.RowFilter = "[primary workunit] = " + PrimaryWU.ToString
                        While TBL_OT.DefaultView.Count > 0
                            TBL_OT.Rows.Remove(TBL_OT.DefaultView(0).Row)
                        End While
                    Next
                End If
            End If

            For Each row_OT In TBL_OT.Rows
                Me.AddSchBar(vTBLSchBar, vColor, row_OT.start, row_OT._end, "", "", -1, -1, row_OT.primary_workunit, EnumTipeSchBar.free_ot, 0, 0, 0, 0, 0, -1, -1, -1, -1)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_OT.Dispose()
            TBL_OT.Dispose()
            ListPrimaryWU1 = Nothing
            ListPrimaryWU2 = Nothing
        End Try
    End Sub

    Private Sub LoadSchBar_SchDT(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vColor As Color, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable, Optional ByVal vPrimaryWU As Double = -1)
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_sch_downtimeTableAdapter
        Dim TBL As New DseTrans_Schedule.tbltrans_sch_downtimeDataTable
        Dim row As DseTrans_Schedule.tbltrans_sch_downtimeRow
        Dim indexLeft As Integer = 0
        Dim indexRight As Integer = 0
        'Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunit1Row
        Dim ListPrimaryWU1 As New List(Of Double)
        Dim ListPrimaryWU2 As New List(Of Double)
        Dim PrimaryWU As Double = -2
        Dim DV As Data.DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimaryWU = -1 Then
                DTA.Fill(TBL, vStartDate)
            Else
                DTA.FillBy_PrimaryWorkUnit(TBL, vStartDate, vPrimaryWU)
            End If

            If IsNothing(vTBLWU) = False Then
                If vTBLWU.Rows.Count > 0 Then
                    'buang item allocated ws, atas work unit yang tidak ada didalam list wu
                    'While indexRight <= TBL.Rows.Count - 1
                    '    If indexLeft <= vTBLWU.Rows.Count - 1 Then
                    '        vrowWU = vTBLWU.Rows(indexLeft)
                    '        row = TBL.Rows(indexRight)
                    '        If vrowWU.primary_work_unit > row.primary_work_unit Then
                    '            TBL.Rows.RemoveAt(indexRight)
                    '        ElseIf vrowWU.primary_work_unit = row.primary_work_unit Then
                    '            indexRight += 1
                    '        Else
                    '            indexLeft += 1
                    '        End If
                    '    Else
                    '        TBL.Rows.RemoveAt(indexRight)
                    '    End If
                    'End While

                    TBL.DefaultView.Sort = "[primary work unit]"
                    TBL.DefaultView.RowFilter = ""
                    For Each DV In TBL.DefaultView
                        row = DV.Row
                        If row.primary_work_unit <> PrimaryWU Then
                            PrimaryWU = row.primary_work_unit
                            ListPrimaryWU1.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU1
                        vTBLWU.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        If vTBLWU.DefaultView.Count = 0 Then
                            'tidak ada di tabel wu, catat
                            ListPrimaryWU2.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU2
                        TBL.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        While TBL.DefaultView.Count > 0
                            TBL.Rows.Remove(TBL.DefaultView(0).Row)
                        End While
                    Next
                End If
            End If

            For Each row In TBL.Rows
                With row
                    Me.AddSchBar(vTBLSchBar, vColor, row.start_time, row.end_time, "", "", -1, -1, .primary_work_unit, EnumTipeSchBar.scheduled_dt, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', -1)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListPrimaryWU1 = Nothing
            ListPrimaryWU2 = Nothing
        End Try
    End Sub

    Private Sub LoadSchBar_AllocatedWS(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vColor As Color, _
    ByVal TBL_WS_Item As DseTrans_Schedule.tblmaster_workshift_itemDataTable, ByVal TBL_WS_ItemHoliday As DseTrans_Schedule.tblmaster_workshift_itemholidayDataTable, _
    ByVal vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable, Optional ByVal vPrimaryWU As Double = -1)
        Dim DTA_WS_Allocate As New DseTrans_ScheduleTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBL_WS_Allocate As New DseTrans_Schedule.tbltrans_mo_workshift_allocateDataTable
        Dim row_WS_Allocate As DseTrans_Schedule.tbltrans_mo_workshift_allocateRow
        Dim row_WS_Item As DseTrans_Schedule.tblmaster_workshift_itemRow
        Dim dow As DayOfWeek = DayOfWeek.Friday
        Dim tipe_work_day As Core_Module.ClsMasterWorkShift.TipeWorkdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday
        Dim indexLeft As Integer = 0
        Dim indexRight As Integer = 0
        'Dim vrowWU As DseTrans_Schedule.tblmaster_product_item_workunit1Row
        Dim ListPrimaryWU1 As New List(Of Double)
        Dim ListPrimaryWU2 As New List(Of Double)
        Dim PrimaryWU As Double = -2
        Dim DV As Data.DataRowView

        Try
            DTA_WS_Allocate.Connection.ConnectionString = Core_Module.GetConstring

            If vPrimaryWU = -1 Then
                DTA_WS_Allocate.Fill(TBL_WS_Allocate, vStartDate)
            Else
                DTA_WS_Allocate.FillBy_PrimaryWU(TBL_WS_Allocate, vStartDate, vPrimaryWU)
            End If

            If IsNothing(vTBLWU) = False Then
                If vTBLWU.Rows.Count > 0 Then
                    'buang item allocated ws, atas work unit yang tidak ada didalam list wu
                    'While indexRight <= TBL_WS_Allocate.Rows.Count - 1
                    '    If indexLeft <= vTBLWU.Rows.Count - 1 Then
                    '        vrowWU = vTBLWU.Rows(indexLeft)
                    '        row_WS_Allocate = TBL_WS_Allocate.Rows(indexRight)
                    '        If vrowWU.primary_work_unit > row_WS_Allocate.primary_work_unit Then
                    '            TBL_WS_Allocate.Rows.RemoveAt(indexRight)
                    '        ElseIf vrowWU.primary_work_unit = row_WS_Allocate.primary_work_unit Then
                    '            indexRight += 1
                    '        Else
                    '            indexLeft += 1
                    '        End If
                    '    Else
                    '        TBL_WS_Allocate.Rows.RemoveAt(indexRight)
                    '    End If
                    'End While

                    'cari primary work unit yang harus dibuang
                    TBL_WS_Allocate.DefaultView.Sort = "[primary work unit]"
                    TBL_WS_Allocate.DefaultView.RowFilter = ""
                    For Each DV In TBL_WS_Allocate.DefaultView
                        row_WS_Allocate = DV.Row
                        If row_WS_Allocate.primary_work_unit <> PrimaryWU Then
                            PrimaryWU = row_WS_Allocate.primary_work_unit
                            ListPrimaryWU1.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU1
                        vTBLWU.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        If vTBLWU.DefaultView.Count = 0 Then
                            'tidak ada di tabel wu, catat
                            ListPrimaryWU2.Add(PrimaryWU)
                        End If
                    Next
                    For Each PrimaryWU In ListPrimaryWU2
                        TBL_WS_Allocate.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                        While TBL_WS_Allocate.DefaultView.Count > 0
                            TBL_WS_Allocate.Rows.Remove(TBL_WS_Allocate.DefaultView(0).Row)
                        End While
                    Next
                End If
            End If


            For Each row_WS_Allocate In TBL_WS_Allocate.Rows
                'cek apakah hari tsb libur
                TBL_WS_ItemHoliday.DefaultView.RowFilter = "[primary workshift] = " + row_WS_Allocate.primary_workshift.ToString + _
                " AND date = #" + Format(row_WS_Allocate._date, "MM/dd/yyyy") + "#"
                If TBL_WS_ItemHoliday.DefaultView.Count = 0 Then
                    'jika bukan libur maka lanjut
                    dow = row_WS_Allocate._date.DayOfWeek
                    TBL_WS_Item.DefaultView.RowFilter = "[primary workshift] = " + row_WS_Allocate.primary_workshift.ToString + _
                    " AND [day of week] = " + CInt(dow).ToString
                    If TBL_WS_Item.DefaultView.Count > 0 Then
                        row_WS_Item = TBL_WS_Item.DefaultView(0).Row
                        'hanya masukkan sch weekday, yang holiday di abaikan
                        If row_WS_Item.tipe_workdays = CInt(tipe_work_day) Then
                            'untuk workshift normal, tidak pindah hari
                            If row_WS_Item.break_duration = 0 Then
                                '1 bar saja
                                'masukkan ke dalam sch bar
                                If row_WS_Item.schedule_in < row_WS_Item.schedule_out Then
                                    'untuk jadwal yang tidak pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                Else
                                    'untuk jadwal yang pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date + TimeSpan.FromDays(1), "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                End If
                            Else
                                '2 bar, sebelum break dan stlh break
                                If row_WS_Item.schedule_in < row_WS_Item.schedule_break_out Then
                                    'untuk jadwal yang tidak pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_break_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                Else
                                    'untuk jadwal yang pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date + TimeSpan.FromDays(1), "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_break_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                End If
                                If row_WS_Item.schedule_break_in < row_WS_Item.schedule_out Then
                                    'untuk jadwal yang tidak pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_break_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                Else
                                    'untuk jadwal yang pindah hari
                                    Me.AddSchBar(vTBLSchBar, vColor, CDate(Format(row_WS_Allocate._date, "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_break_in).Replace(".", ":")), _
                                    CDate(Format(row_WS_Allocate._date + TimeSpan.FromDays(1), "dd MMM yyyy") + " " + Me.ObjCD.Convert_ToString(row_WS_Item.schedule_out).Replace(".", ":")), _
                                    "", "", -1, -1, row_WS_Allocate.primary_work_unit, EnumTipeSchBar.free_time, 0, 0, 0, 0, 0, -1, -1, -1, -1) ', row_WS_Allocate.primary_workshift)
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Logic error: workshift item not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_WS_Allocate.Dispose()
            TBL_WS_Allocate.Dispose()
            ListPrimaryWU1 = Nothing
            ListPrimaryWU2 = Nothing
        End Try
    End Sub

    Private Sub AddSchBar(ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vColor As Color, ByVal vPlanStartTime As Date, _
    ByVal vPlanEndTime As Date, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryMO As Double, ByVal vPrimaryProcess As Double, _
    ByVal vPrimaryWorkUnit As Double, ByVal vTipeBar As EnumTipeSchBar, ByVal vQty As Decimal, ByVal vRate As Decimal, ByVal vTotalWH As Double, _
    ByVal vProgress As Decimal, ByVal vDecimalPoint As Integer, ByVal vPrimaryUnit As Double, ByVal vPrimaryWUConfig As Double, _
    ByVal vPrimaryInputLocation As Double, ByVal vPrimaryOutputLocation As Double) ', ByVal vPrimaryWorkshift As Double)
        Dim row As DseTrans_Schedule.tbl_sch_barRow

        Try
            If vPrimaryWorkUnit = -1 Then
                MsgBox("Logic error: invalid primary workunit, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If vTipeBar = EnumTipeSchBar.work_time Then
                If vQty = 0 Or vRate = 0 Or vTotalWH = 0 Then
                    MsgBox("Logic error: invalid qty or rate, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vPrimaryUnit = -1 Then
                    MsgBox("Logic error: invalid unit data, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vPrimaryInputLocation = -1 Then
                    MsgBox("Logic error: invalid input location, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vPrimaryOutputLocation = -1 Then
                    MsgBox("Logic error: invalid output location, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Else
                If vQty <> 0 Or vRate <> 0 Or vTotalWH <> 0 Then
                    MsgBox("Logic error: invalid qty or rate, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If vTipeBar = EnumTipeSchBar.work_time Or vTipeBar = EnumTipeSchBar.work_time_ot Then
                'harus ada: primary workshift, primary mo, identifier jo, identifier wo, primary process, primary work unit
                'If vPrimaryWorkshift = -1 Then
                '    MsgBox("Logic error: invalid primary workshift, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                '    Exit Sub
                'End If
                If vPrimaryMO = -1 Then
                    MsgBox("Logic error: invalid primary manufacturing order, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vIdentifierJO = "" Then
                    MsgBox("Logic error: invalid identifier jo, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vIdentifierWO = "" Then
                    MsgBox("Logic error: invalid identifier wo, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If vPrimaryProcess = -1 Then
                    MsgBox("Logic error: invalid primary process, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            row = vTBLSchBar.Newtbl_sch_barRow
            With row
                .a = vColor.A
                .b = vColor.B
                .g = vColor.G
                .r = vColor.R
                .identifier_jo = vIdentifierJO
                .identifier_wo = vIdentifierWO
                .plan_end_time = vPlanEndTime
                .plan_start_time = vPlanStartTime
                .primary_mo = vPrimaryMO
                .primary_process = vPrimaryProcess
                .primary_work_unit = vPrimaryWorkUnit
                .tipe = CInt(vTipeBar)
                .duration = (.plan_end_time - .plan_start_time).TotalMinutes
                .total_work_hour = vTotalWH
                .qty_produce = vQty
                .rate = vRate
                .progress = vProgress
                .primary_unit = vPrimaryUnit
                .primary_wu_config = vPrimaryWUConfig
                .decimal_point = vDecimalPoint
                .primary_input_location = vPrimaryInputLocation
                .primary_output_location = vPrimaryOutputLocation
                If .plan_end_time.Hour = 0 And .plan_end_time.Minute = 0 And .tipe = CInt(EnumTipeSchBar.free_time) Then
                    'MsgBox("")
                End If
            End With
            vTBLSchBar.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTipeSchBar(ByVal vRefDate As Date, ByVal vPrimaryWU As Double, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As EnumTipeSchBar
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow

        Try
            vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                rowSchBar = DV.Row
                If rowSchBar.plan_start_time <= vRefDate And rowSchBar.plan_end_time >= vRefDate Then
                    Return rowSchBar.tipe
                End If
            Next

            Return EnumTipeSchBar.off_time
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeSchBar.off_time
        End Try
    End Function

    Public Function GetFirstStartTime(ByVal vRefDate As Date, ByVal vPrimaryWU As Double, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Date
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow

        Try
            vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                rowSchBar = DV.Row
                If rowSchBar.plan_start_time > vRefDate Then
                    If rowSchBar.tipe = CInt(ClsTrans_Schedule.EnumTipeSchBar.free_ot) Or _
                    rowSchBar.tipe = CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time) Then
                        Return rowSchBar.plan_start_time
                    End If
                End If
            Next

            Return #1/1/2200#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'mengembalikan jam workshift pertama, untuk primary work unit dan tanggal yang diberikan
    Public Function GetFirstStartTime(ByVal vRefDate As Date, ByVal vPrimaryWU As Double) As Date
        Dim RetVal As Date = #1/1/1950#
        Dim DTA1 As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBL1 As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim row1, row11 As DseTrans_OT_Item.tbltrans_mo_workshift_allocateRow
        Dim DTA2 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL2 As New DseTrans_OT_Item.tblmaster_workshift_itemDataTable
        Dim row2 As DseTrans_OT_Item.tblmaster_workshift_itemRow
        Dim DTA3 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL3 As New DseTrans_OT_Item.tblmaster_workshift_itemholidayDataTable
        Dim MinSchIn As Double = Double.MaxValue
        Dim TglMin As Date = Today
        Dim mn, ij As Double

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA1.Connection.ConnectionString

            DTA1.FillBy_Primary_WorkUnit(TBL1, vRefDate + TimeSpan.FromDays(30), vRefDate, vPrimaryWU)
            DTA2.FillBy2(TBL2)
            DTA3.Fill(TBL3, vRefDate)

            'isikan jadwal berdasarkan item workshift
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    row1.day_of_week = row1._date.DayOfWeek
                Next
                mn = 0
                For Each row1 In TBL1.Rows
                    If row1.sch_in = -1 Then
                        If TBL2.Rows.Count > 0 Then
                            For Each row2 In TBL2.Rows
                                If row2.primary_transaksi = row1.primary_workshift And row2.day_of_week = row1.day_of_week Then
                                    ij = 0
                                    While mn + ij <= TBL1.Rows.Count - 1
                                        row11 = TBL1.Rows(mn + ij)
                                        If row1.primary_workshift = row11.primary_workshift And row1.day_of_week = row11.day_of_week Then
                                            row11.sch_in = row2.schedule_in
                                            row11.sch_out = row2.schedule_out
                                            row11.tipe_workdays = row2.tipe_workdays
                                        End If
                                        ij += 1
                                    End While

                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    mn += 1
                Next

                'isikan jadwal berdasarkan item workshift
                For Each row1 In TBL1.Rows
                    TBL3.DefaultView.RowFilter = "[primary transaksi] = " + row1.primary_workshift.ToString + " AND date = #" + Format(row1._date, "MM/dd/yyyy") + "#"
                    If TBL3.DefaultView.Count > 0 Then
                        row1.tipe_workdays = CInt(Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday)
                    End If
                Next

                'cari item jadwal kerja pertama terhitung sejak ref date
                For Each row1 In TBL1.Rows
                    If row1.tipe_workdays <> Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday Then
                        If MinSchIn > row1.sch_in Then
                            MinSchIn = row1.sch_in
                            TglMin = row1._date
                            Exit For
                        End If
                    End If
                Next
                If MinSchIn = Double.MaxValue Then
                    Return #1/1/1950#
                Else
                    Return TglMin.Date + TimeSpan.FromMinutes(MinSchIn)
                End If
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            DTA3.Dispose()
            TBL3.Dispose()
        End Try
    End Function

    'Public Sub GeserEndTime(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    'ByVal vNewEndDate As Date, ByVal vTblItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable)
    '    Dim RowSchBar_Dummy, RowSchBar_Before As DseTrans_Schedule.tbl_sch_barRow
    '    Dim ij As Double = 0
    '    Dim mn As Double = 0
    '    Dim IdxGeser As Double = -1
    '    Dim MenitGeser As Double = 0
    '    Dim TBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable
    '    Dim TipeDummy As EnumTipeSchBar = EnumTipeSchBar.free_ot

    '    Try
    '        Me.LoadScheduleBar_TanpaScheduledItem(vRowSchBar.plan_start_time.Date - TimeSpan.FromDays(10), #1/1/2200#, TBLSchBarDummy)

    '        vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString
    '        vTBLSchBar.DefaultView.Sort = "[plan start time]"
    '        While ij <= vTBLSchBar.DefaultView.Count - 1
    '            RowSchBar_Dummy = vTBLSchBar.DefaultView(ij).Row
    '            If IdxGeser = -1 Then
    '                If RowSchBar_Dummy.identifier_wo = vRowSchBar.identifier_wo And _
    '                RowSchBar_Dummy.plan_start_time = vRowSchBar.plan_start_time And _
    '                RowSchBar_Dummy.plan_end_time = vRowSchBar.plan_end_time Then
    '                    RowSchBar_Dummy.plan_end_time = vNewEndDate
    '                    'ubah rate dan hal lainnya yang berkaitan
    '                    '1. rate + total WH
    '                    RowSchBar_Dummy.total_work_hour = Me.GetWorkDuration(RowSchBar_Dummy.plan_start_time, RowSchBar_Dummy.plan_end_time, RowSchBar_Dummy.primary_work_unit, TBLSchBarDummy)
    '                    RowSchBar_Dummy.rate = Math.Round(RowSchBar_Dummy.qty_produce * 60 / RowSchBar_Dummy.total_work_hour, CInt(RowSchBar_Dummy.decimal_point), MidpointRounding.AwayFromZero)
    '                    '2. end time untuk item wo ybs

    '                    '3. rate cummualtive + total WH cummulative

    '                    MsgBox("belum ubah hal lain yang berkaitan")
    '                    MsgBox("belum coding berkaitan dengan overtime")
    '                    IdxGeser = ij
    '                End If
    '            End If

    '            If ij >= IdxGeser + 1 And IdxGeser <> -1 And RowSchBar_Dummy.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
    '                mn = 1
    '                RowSchBar_Before = vTBLSchBar.DefaultView(ij - mn).Row
    '                While RowSchBar_Before.tipe <> ClsTrans_Schedule.EnumTipeSchBar.work_time And ij - mn > IdxGeser
    '                    mn += 1
    '                    RowSchBar_Before = vTBLSchBar.DefaultView(ij - mn).Row
    '                End While
    '                If RowSchBar_Before.tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time Then
    '                    If RowSchBar_Before.plan_end_time > RowSchBar_Dummy.plan_start_time Then
    '                        MenitGeser = Me.GetWorkDuration(RowSchBar_Dummy.plan_start_time, RowSchBar_Before.plan_end_time, vRowSchBar.primary_work_unit, TBLSchBarDummy)
    '                        RowSchBar_Dummy.plan_start_time = RowSchBar_Before.plan_end_time
    '                        'jika akhir dari worktime/ot, maka ganti start time ke worktime/free ot paling dekat
    '                        TipeDummy = Me.GetTipeSchBar(RowSchBar_Dummy.plan_start_time + TimeSpan.FromMinutes(1), vRowSchBar.primary_work_unit, TBLSchBarDummy)
    '                        If TipeDummy = EnumTipeSchBar.off_time Or TipeDummy = EnumTipeSchBar.scheduled_dt Then
    '                            RowSchBar_Dummy.plan_start_time = Me.GetFirstStartTime(RowSchBar_Dummy.plan_start_time, vRowSchBar.primary_work_unit, TBLSchBarDummy)
    '                        End If
    '                        RowSchBar_Dummy.plan_end_time = Me.GetEndTime(vRowSchBar.primary_work_unit, RowSchBar_Dummy.plan_end_time, MenitGeser, TBLSchBarDummy)
    '                    Else
    '                        'kalo sampe 1x geser, maka selanjutnya ga geser juga, keluar fungsi aj
    '                        Exit Sub
    '                    End If
    '                End If
    '            End If
    '            ij += 1
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Function ApakahBolehHapusJadwalJO(ByVal vIdentifierJO As String, ByVal vObjMO As ClsTrans_ManufacturingOrder, ByVal vObjLog As ClsTrans_StockBundling) As Boolean
        Dim FirmPlanned As Date = Today

        Try
            'tidak masuk ke masa firm planned
            'bukan dari bagian jo yang sudah selesai
            'belum ada log produksi

            'tidak masuk ke masa firm planned
            FirmPlanned = Me.GetFirmPlan
            If vObjMO.GetEarliestStartTime(vIdentifierJO).Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't delete the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'bukan dari bagian jo yang sudah selesai
            If vObjMO.IsJOFinished(vIdentifierJO) = True Then
                MsgBox("This job order is closed, you can't delete the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'belum ada log produksi
            If vObjLog.GetJml_ByIdentifierJO(vIdentifierJO) > 0 Then
                MsgBox("This job order is refered by 1 or more production log, you can't delete the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehGeserStartTime_FormUtama(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vNewStartDate As Date, ByVal vNewEndDate1 As Date, ByVal vStartDateLama As Date, ByVal vDurasiGeser As Double) As Boolean
        Dim FirmPlanned As Date = Today
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim ObjLog As New ClsTrans_StockBundling
        Dim DTAWS_All As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS_All As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim Ketemu As Boolean = False
        Dim RowSchBar_Dummy, RowSchBar_Dummy1 As DseTrans_Schedule.tbl_sch_barRow
        Dim MaxEndTime As Date

        Try
            'tidak masuk ke masa firm planned
            'bukan dari bagian jo yang sudah selesai -> di form
            'belum ada log untuk item sch tsb -> di form
            'pada start date baru, workshift harus sudah allocate
            'pada end date baru, workshift harus sudah allocate -> di form
            'start date harus dalam work time, free time, free ot
            'klo geser kiri start date tidak boleh tabrakan dengan item wo sch yang sudah ada
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            'tidak masuk ke masa firm planned
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_start_time.Date <= FirmPlanned.Date Or vNewStartDate.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If


            'pada start date baru, workshift harus sudah allocate
            DTAWS_All.FillBy_Primary_WorkUnit(TBLWS_All, vNewStartDate.Date, vNewStartDate.Date, vRowSchBar.primary_work_unit)
            If TBLWS_All.Rows.Count = 0 Then
                MsgBox("Please allocate workshift for this work unit on " + Format(vNewStartDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pada end date baru, workshift harus sudah allocate -> di cek pada form

            'start date harus dalam work time, free time, free ot
            Ketemu = False
            vTBLSchBarTanpaWO.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
            " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
            vTBLSchBarTanpaWO.DefaultView.Sort = "[plan start time]"
            For Each dv As Data.DataRowView In vTBLSchBarTanpaWO.DefaultView
                RowSchBar_Dummy = dv.Row
                If RowSchBar_Dummy.plan_start_time <= vNewStartDate And RowSchBar_Dummy.plan_end_time >= vNewStartDate Then
                    Ketemu = True
                    Exit For
                ElseIf RowSchBar_Dummy.plan_start_time > vNewStartDate Then
                    Exit For
                End If
            Next
            If Ketemu = False Then
                MsgBox("Please set end time on a valid work time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vNewStartDate < vStartDateLama Then
                'start time baru tidak boleh tabrakan dengan item wo sch yang sudah ada
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
                                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
                    RowSchBar_Dummy = dv.Row
                    If RowSchBar_Dummy.plan_start_time <= vNewStartDate And RowSchBar_Dummy.plan_end_time > vNewStartDate Then
                        If RowSchBar_Dummy.identifier_wo <> vRowSchBar.identifier_wo Then
                            MsgBox("Conflicting with another work order on start time, please select another start time.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    ElseIf RowSchBar_Dummy.plan_start_time > vNewStartDate Then
                        Exit For
                    End If
                Next

                'tidak boleh lebih kiri dari item sch yang sebelumnya dia
                If vTBLSchBar.DefaultView.Count > 1 Then
                    For ij As Integer = 1 To vTBLSchBar.DefaultView.Count - 1
                        RowSchBar_Dummy = vTBLSchBar.DefaultView(ij).Row
                        RowSchBar_Dummy1 = vTBLSchBar.DefaultView(ij - 1).Row
                        If RowSchBar_Dummy.identifier_wo = vRowSchBar.identifier_wo Then
                            If RowSchBar_Dummy1.plan_end_time > vNewStartDate Then
                                MsgBox("Conflicting with another work order minimum value you can enter is: " + Format(RowSchBar_Dummy1.plan_end_time, "dd MMM yyyy HH.mm"), MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If


            'end time baru tidak boleh tabrakan dengan item wo sch yang sudah ada
            'vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
            '" AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            'vTBLSchBar.DefaultView.Sort = "[plan start time]"
            'For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
            '    RowSchBar_Dummy = dv.Row
            '    If RowSchBar_Dummy.plan_start_time <= vNewEndDate1 And RowSchBar_Dummy.plan_end_time > vNewEndDate1 Then
            '        If RowSchBar_Dummy.identifier_wo <> vRowSchBar.identifier_wo Then
            '            MsgBox("Conflicting with another work order on end time, please select another start time.", MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    ElseIf RowSchBar_Dummy.plan_start_time > vNewEndDate1 Then
            '        Exit For
            '    End If
            'Next


            'pergeseran start date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            If vDurasiGeser > 0 Then
                MaxEndTime = Me.GetMaxEndTime(vRowSchBar.primary_work_unit)
                If MaxEndTime <> #1/1/1950# Then
                    MaxEndTime = Me.GetEndTime(vRowSchBar.primary_work_unit, MaxEndTime, vDurasiGeser, vTBLSchBarTanpaWO)
                    If MaxEndTime = #1/1/2200# Then
                        MsgBox("Insufficient allocated work hour, please allocate more work hour for this work unit.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjMO = Nothing
            ObjLog = Nothing
        End Try
    End Function

    Public Function GetMaxEndTime(ByVal vPrimaryWU As Double) As Date
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetMaxEndTime(vPrimaryWU)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehUbahRate_FormUtama(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vDurasiGeser As Double, ByVal vNewRate As Decimal) As Boolean
        Dim FirmPlanned As Date = Today
        Dim MaxEndTime As Date

        Try
            'tidak masuk ke firm plan
            'rate harus positive
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            'tidak masuk ke firm plan
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_start_time.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'rate harus positive
            If vNewRate <= 0 Then
                MsgBox("Please specify a positive number", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            If vDurasiGeser > 0 Then
                MaxEndTime = Me.GetMaxEndTime(vRowSchBar.primary_work_unit)
                If MaxEndTime <> #1/1/1950# Then
                    MaxEndTime = Me.GetEndTime(vRowSchBar.primary_work_unit, MaxEndTime, vDurasiGeser, vTBLSchBarTanpaWO)
                    If MaxEndTime = #1/1/2200# Then
                        MsgBox("Insufficient allocated work hour, please allocate more work hour for this work unit.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehUbahWH_FormUtama(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vDurasiGeser As Double) As Boolean
        Dim FirmPlanned As Date = Today
        Dim MaxEndTime As Date

        Try
            'tidak masuk ke firm plan
            'rate harus positive
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            'tidak masuk ke firm plan
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_start_time.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            If vDurasiGeser > 0 Then
                MaxEndTime = Me.GetMaxEndTime(vRowSchBar.primary_work_unit)
                If MaxEndTime <> #1/1/1950# Then
                    MaxEndTime = Me.GetEndTime(vRowSchBar.primary_work_unit, MaxEndTime, vDurasiGeser, vTBLSchBarTanpaWO)
                    If MaxEndTime = #1/1/2200# Then
                        MsgBox("Insufficient allocated work hour, please allocate more work hour for this work unit.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehGeserEndTime_FormUtama(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vNewEndDate As Date, ByVal vEndDateLama As Date, ByVal vStartTimeLama As Date, ByVal vDurasiGeser As Double) As Boolean
        Dim FirmPlanned As Date = Today
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim ObjLog As New ClsTrans_StockBundling
        Dim DTAWS_All As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS_All As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim Ketemu As Boolean = False
        Dim RowSchBar_Dummy As DseTrans_Schedule.tbl_sch_barRow
        Dim MaxEndTime As Date

        Try
            'tidak masuk ke masa firm planned
            'bukan dari bagian jo yang sudah selesai -> di form
            'belum ada log untuk item sch tsb -> di form
            'pada start date baru, workshift harus sudah allocate
            'pada end date baru, workshift harus sudah allocate -> di form
            'start date harus dalam work time, free time, free ot
            'klo geser kiri start date tidak boleh tabrakan dengan item wo sch yang sudah ada
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            'tidak masuk ke masa firm planned
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_start_time.Date <= FirmPlanned.Date Or vNewEndDate.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If


            'pada start date baru, workshift harus sudah allocate
            DTAWS_All.FillBy_Primary_WorkUnit(TBLWS_All, vNewEndDate.Date, vNewEndDate.Date, vRowSchBar.primary_work_unit)
            If TBLWS_All.Rows.Count = 0 Then
                MsgBox("Please allocate workshift for this work unit on " + Format(vNewEndDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pada end date baru, workshift harus sudah allocate -> di cek pada form

            'start date harus dalam work time, free time, free ot
            Ketemu = False
            vTBLSchBarTanpaWO.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
            " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
            vTBLSchBarTanpaWO.DefaultView.Sort = "[plan start time]"
            For Each dv As Data.DataRowView In vTBLSchBarTanpaWO.DefaultView
                RowSchBar_Dummy = dv.Row
                If RowSchBar_Dummy.plan_start_time <= vNewEndDate And RowSchBar_Dummy.plan_end_time >= vNewEndDate Then
                    Ketemu = True
                    Exit For
                ElseIf RowSchBar_Dummy.plan_start_time > vNewEndDate Then
                    Exit For
                End If
            Next
            If Ketemu = False Then
                MsgBox("Please set end time on a valid work time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'end baru tidak boleh tabrakan dengan item wo sch yang sudah ada
            'vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
            '     " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            'vTBLSchBar.DefaultView.Sort = "[plan start time]"
            'For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
            '    RowSchBar_Dummy = dv.Row
            '    If RowSchBar_Dummy.plan_start_time <= vNewEndDate And RowSchBar_Dummy.plan_end_time > vNewEndDate Then
            '        If RowSchBar_Dummy.identifier_wo <> vRowSchBar.identifier_wo Then
            '            MsgBox("Conflicting with another work order on start time, please select another start time.", MsgBoxStyle.Information, "Information")
            '            Return False
            '        End If
            '    ElseIf RowSchBar_Dummy.plan_start_time > vNewEndDate Then
            '        Exit For
            '    End If
            'Next

            'end time baru harus lebih besar dari start time
            If vStartTimeLama >= vNewEndDate Then
                MsgBox("Minimum value you can enter is: " + Format(vStartTimeLama + TimeSpan.FromMinutes(1), "dd MMM yyyy HH.mm") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            If vDurasiGeser > 0 Then
                MaxEndTime = Me.GetMaxEndTime(vRowSchBar.primary_work_unit)
                If MaxEndTime <> #1/1/1950# Then
                    MaxEndTime = Me.GetEndTime(vRowSchBar.primary_work_unit, MaxEndTime, vDurasiGeser, vTBLSchBarTanpaWO)
                    If MaxEndTime = #1/1/2200# Then
                        MsgBox("Insufficient allocated work hour, please allocate more work hour for this work unit.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjMO = Nothing
            ObjLog = Nothing
        End Try
    End Function

    Public Function ApakahBolehGeserStartTime(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vNewStartDate As Date, ByVal vNewEndDate1 As Date, ByVal vStartDateLama As Date, _
    ByVal vTipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable) As Boolean
        Dim FirmPlanned As Date = Today
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim DTAWS_All As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS_All As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim RowSchBar_Dummy As DseTrans_Schedule.tbl_sch_barRow
        Dim Ketemu As Boolean = False

        Try
            DTAWS_All.Connection.ConnectionString = Core_Module.GetConstring
            'tidak masuk ke masa firm planned
            'bukan dari bagian jo yang sudah selesai
            'pada start date baru, workshift harus sudah allocate
            'pada end date baru, workshift harus sudah allocate
            'start date harus dalam work time, free time, free ot
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            'tidak masuk ke masa firm planned
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_start_time.Date <= FirmPlanned.Date Or vNewStartDate.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'bukan dari bagian jo yang sudah selesai
            If ObjMO.ApakahJOSudahSelesai(vRowSchBar.identifier_jo) = True Then
                MsgBox("The specified job order is closed, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pada start date baru, workshift harus sudah allocate
            DTAWS_All.FillBy_Primary_WorkUnit(TBLWS_All, vNewStartDate.Date, vNewStartDate.Date, vRowSchBar.primary_work_unit)
            If TBLWS_All.Rows.Count = 0 Then
                MsgBox("Please allocate workshift for this work unit on " + Format(vNewStartDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'start date baru harus dalam work time, free time, free ot
            Ketemu = False
            vTBLSchBarKosong.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
            vTBLSchBarKosong.DefaultView.Sort = "[plan start time]"
            For Each dv As Data.DataRowView In vTBLSchBarKosong.DefaultView
                RowSchBar_Dummy = dv.Row
                If RowSchBar_Dummy.plan_start_time <= vNewStartDate And RowSchBar_Dummy.plan_end_time >= vNewStartDate Then
                    Ketemu = True
                    Exit For
                ElseIf RowSchBar_Dummy.plan_start_time > vNewStartDate Then
                    Exit For
                End If
            Next
            If Ketemu = False Then
                MsgBox("Please set end time on a valid work time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'klo geser kiri start date tidak boleh tabrakan dengan item wo sch yang sudah ada
            If vStartDateLama > vNewStartDate Then
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + _
                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
                    RowSchBar_Dummy = dv.Row
                    If RowSchBar_Dummy.plan_start_time <= vNewStartDate And RowSchBar_Dummy.plan_end_time > vNewStartDate Then
                        MsgBox("Conflicting with another work order, please select another start time.", MsgBoxStyle.Information, "Information")
                        Return False
                    ElseIf RowSchBar_Dummy.plan_start_time > vNewStartDate Then
                        Exit For
                    End If
                Next

                'klo geser kiri start date tidak boleh tabrakan dengan item wo anak2-nya, dengan mempertimbangkan dependency
                vTBLSchBar.DefaultView.RowFilter = GetFilterString_IdentifierWOParent(vRowSchBar.identifier_wo, vTBLItemWO)
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                For Each dv As Data.DataRowView In vTBLSchBar.DefaultView
                    RowSchBar_Dummy = dv.Row
                    If vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1 Then
                        'tidak boleh sebelum end date
                        'If RowSchBar_Dummy.plan_end_time > vNewStartDate Then
                        '    MsgBox("Conflicting with another work order, please select another start time.", MsgBoxStyle.Information, "Information")
                        '    Return False
                        'End If
                    ElseIf vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start1 Or vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start_Strict1 Then
                        If RowSchBar_Dummy.plan_start_time > vNewStartDate Then
                            MsgBox("Conflicting with another work order, please select another start time.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If


            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            'MsgBox("belum: ini dicoding, setelah coding pergeseran")

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjMO = Nothing
        End Try
    End Function

    Public Function ApakahBolehGeserEndTime(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow, ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vNewEndDate As Date) As Boolean
        Dim DTAWS_All As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS_All As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim RowSchBar_Dummy As DseTrans_Schedule.tbl_sch_barRow
        Dim Ketemu As Boolean = False
        Dim FirmPlanned As Date = Today
        Dim ObjMO As New ClsTrans_ManufacturingOrder

        Try
            'tidak masuk ke masa firm planned
            'bukan dari bagian jo yang sudah selesai
            'start date harus < end date
            'pada end date, workshift harus sudah allocate
            'end date harus dalam work time, free time, free ot
            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja

            DTAWS_All.Connection.ConnectionString = Core_Module.GetConstring

            'tidak masuk ke masa firm planned
            FirmPlanned = Me.GetFirmPlan
            If vRowSchBar.plan_end_time.Date <= FirmPlanned.Date Or vNewEndDate.Date <= FirmPlanned.Date Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'bukan dari bagian jo yang sudah selesai
            If ObjMO.ApakahJOSudahSelesai(vRowSchBar.identifier_jo) = True Then
                MsgBox("The specified job order is closed, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'start date harus < end date
            If vRowSchBar.plan_start_time >= vNewEndDate Then
                MsgBox("Start time must be smaller than end time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pada end date, workshift harus sudah allocate
            DTAWS_All.FillBy_Primary_WorkUnit(TBLWS_All, vNewEndDate.Date, vNewEndDate.Date, vRowSchBar.primary_work_unit)
            If TBLWS_All.Rows.Count = 0 Then
                MsgBox("Please allocate workshift for this work unit on " + Format(vNewEndDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'end date harus dalam work time, free time, free ot
            vTBLSchBarKosong.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar.primary_work_unit.ToString + " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
            vTBLSchBarKosong.DefaultView.Sort = "[plan start time]"
            For Each dv As Data.DataRowView In vTBLSchBarKosong.DefaultView
                RowSchBar_Dummy = dv.Row
                If RowSchBar_Dummy.plan_start_time <= vNewEndDate And RowSchBar_Dummy.plan_end_time >= vNewEndDate Then
                    Ketemu = True
                    Exit For
                ElseIf RowSchBar_Dummy.plan_start_time > vNewEndDate Then
                    Exit For
                End If
            Next
            If Ketemu = False Then
                MsgBox("Please set end time on a valid work time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pergeseran end date, tidak boleh menyebabkan ada work time yang tidak kebagian alokasi waktu kerja
            'MsgBox("belum: ini dicoding, setelah coding pergeseran")

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTAWS_All.Dispose()
            TBLWS_All.Dispose()
            ObjMO = Nothing
        End Try
    End Function

    Public Function ApakahFreeTime(ByVal vRefDate As Date, ByVal vPrimaryWU As Double, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Boolean
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow

        Try
            vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString
            vTBLSchBar.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In vTBLSchBar.DefaultView
                rowSchBar = DV.Row
                If rowSchBar.plan_start_time <= vRefDate And rowSchBar.plan_end_time >= vRefDate Then
                    If rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_time Or _
                    rowSchBar.tipe = ClsTrans_Schedule.EnumTipeSchBar.free_ot Then
                        Return True
                    End If
                ElseIf rowSchBar.plan_start_time > vRefDate Then
                    Return False
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Sub Reschedule_JadwalBaru1(ByVal vIdentifierJO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
   ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
   ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
   ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
   ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vStartTimeAcuan As Date, ByVal vPrimaryWUConfig As Double, Optional ByVal vPrimaryWU_Acuan As Double = -1, _
   Optional ByVal vIdentifierWO_Acuan As String = "")
        Dim ij As New FrmTrans_ProductionSchedule_Progress

        Try
            ij.vIdentifierJO = vIdentifierJO
            ij.vTBLSchBar = vTBLSchBar
            ij.vTBLItemWO = vTBLItemWO
            ij.vTBLItemWO_ItemSch = vTBLItemWO_ItemSch
            ij.a = a
            ij.r = r
            ij.g = g
            ij.b = b
            ij.vTBLSchBarKosong = vTBLSchBarKosong
            ij.vTBLSchBarTanpaWO = vTBLSchBarTanpaWO
            ij.vPrimaryWU_Acuan = vPrimaryWU_Acuan
            ij.vIdentifierWO_Acuan = vIdentifierWO_Acuan
            ij.vStartTimeAcuan = vStartTimeAcuan
            ij.ObjSch = Me
            ij.vPrimaryWUConfig = vPrimaryWUConfig
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub Reschedule_NewRate(ByVal vIdentifierJO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
 ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
 ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
 ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
 ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
 ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
  ByVal vPrimaryWUConfig As Double, ByVal vTBLRefRate As DseTrans_Schedule2.TblRefDataTable)
        Dim ij As New FrmTrans_ProductionSchedule_Progress6

        Try
            ij.vIdentifierJO = vIdentifierJO
            ij.vTBLSchBar = vTBLSchBar
            ij.vTBLItemWO = vTBLItemWO
            ij.vTBLItemWO_ItemSch = vTBLItemWO_ItemSch
            ij.vTBLRefRate = vTBLRefRate
            ij.a = a
            ij.r = r
            ij.g = g
            ij.b = b
            ij.vTBLSchBarKosong = vTBLSchBarKosong
            ij.vTBLSchBarTanpaWO = vTBLSchBarTanpaWO 
            ij.ObjSch = Me
            ij.vPrimaryWUConfig = vPrimaryWUConfig
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub Reschedule_NewRate(ByVal vIdentifierJO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
 ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
 ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
 ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
 ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
 ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, _
 ByVal FrmProgres As FrmTrans_ProductionSchedule_Progress6, ByVal vPrimaryWUConfig As Double, ByVal vTBLRefRate As DseTrans_Schedule2.TblRefDataTable)
        Dim FirmPlanDate, StartDate, EndDate As Date
        Dim vTBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vTBLSchBarDummy2 As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vTBLSchBar_Reserve As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vRowSchBar_Reserve As DseTrans_Schedule.tbl_sch_barRow
        Dim vRowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DV As Data.DataRowView
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim StartTime As Date = Today
        Dim StartTime1 As Date = Today
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim DelayDependency As Double = 0
        Dim TBLSchItem As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim RowSchItem, RowSchItem2 As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Start1, End1 As Date
        Dim JmlWUActive As Integer = 0
        Dim TotalProduce As Decimal = 0
        Dim TotalWH As Decimal = 0
        Dim DummyRate As Decimal = 0
        Dim CurStep As Integer = 0
        Dim TotalStep As Integer = 5
        Dim IdentifierWODummy As String = ""
        Dim rowRefRate As DseTrans_Schedule2.TblRefRow

        Try
            Start1 = Date.MaxValue
            End1 = Date.MinValue

            'pindahkan semua baris atas identifier jo ybs
            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": duplicating old data"
            vTBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + vIdentifierJO + "' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            FrmProgres.prg1.Value = 0
            FrmProgres.prg1.Maximum = vTBLSchBar.DefaultView.Count
            For Each DV In vTBLSchBar.DefaultView
                vRowSchBar = DV.Row
                FrmProgres.prg1.Value += 1
                FrmProgres.LblProgress.Text = "Data " + FrmProgres.prg1.Value.ToString + " of " + FrmProgres.prg1.Maximum.ToString
                vRowSchBar_Reserve = vTBLSchBar_Reserve.Newtbl_sch_barRow
                For ij As Integer = 0 To vTBLSchBar_Reserve.Columns.Count - 1
                    vRowSchBar_Reserve.Item(ij) = vRowSchBar.Item(ij)
                Next
                vTBLSchBar_Reserve.Rows.Add(vRowSchBar_Reserve)
            Next

            'buat tabel sch baru yang siap untuk alokasi ulang
            FirmPlanDate = Me.GetFirmPlan
            StartDate = FirmPlanDate + TimeSpan.FromDays(1)
            EndDate = Me.GetLast_WS_AllocatedDate
            If Me.GetLast_WS_OvertimeDate.Date > EndDate Then
                EndDate = Me.GetLast_WS_OvertimeDate.Date
            End If

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": preparing gantt chart"
            FrmProgres.prg1.Value = 0
            Me.CopyTblSchBar(vTBLSchBarKosong, vTBLSchBarDummy)

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": preparing reference gantt chart"
            FrmProgres.prg1.Value = 0
            Me.CopyTblSchBar(vTBLSchBarTanpaWO, vTBLSchBarDummy2)
            'Me.LoadScheduleBar(StartDate, EndDate, vTBLSchBarDummy, Nothing)
            'Me.LoadScheduleBar_TanpaScheduledItem(StartDate, EndDate, vTBLSchBarDummy2) 'ini masih bisa optimasi lagi kalo emang berasa lambat

            'alokasikan ulang jadwal berdasarkan urutan level
            'basis alokasi adalah: qty, duration, tipe dependency, dan delay dependency
            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": rescheduling data"
            FrmProgres.prg1.Value = 0
            FrmProgres.prg1.Maximum = vTBLItemWO.Rows.Count
            For Each rowItemWO In vTBLItemWO.Rows
                FrmProgres.prg1.Value += 1
                FrmProgres.LblProgress.Text = "Data " + FrmProgres.prg1.Value.ToString + " of " + FrmProgres.prg1.Maximum.ToString + " - Process: " + rowItemWO.Process_Code + " - Code No: " + rowItemWO.Product_Code

                Start1 = Date.MaxValue
                End1 = Date.MinValue
                JmlWUActive = 0
                TotalWH = 0
                TotalProduce = 0

                vTBLSchBar_Reserve.DefaultView.RowFilter = "[identifier wo] = '" + rowItemWO.Identifier + "'"
                vTBLSchBar_Reserve.DefaultView.Sort = "[plan start time]"
                TipeDependency = rowItemWO.tipe_dependency
                DelayDependency = rowItemWO.delay_dependency

                '1. tentukan start time berdasarkan end time proses sebelumnya
                StartTime = #1/1/1950#

                For Each DV In vTBLSchBar_Reserve.DefaultView
                    JmlWUActive += 1

                    vRowSchBar_Reserve = DV.Row

                    'tentukan nilai rate yang baru
                    For Each rowRefRate In vTBLRefRate.Rows
                        If rowRefRate.Primary_Process = vRowSchBar_Reserve.primary_process And _
                        rowRefRate.Primary_WU = vRowSchBar_Reserve.primary_work_unit Then
                            vRowSchBar_Reserve.rate = rowRefRate.Rate
                            Exit For
                        End If
                    Next
                    DummyRate = vRowSchBar_Reserve.rate
                    vRowSchBar_Reserve.total_work_hour = Me.GetWorkDuration(vRowSchBar_Reserve.rate, vRowSchBar_Reserve.qty_produce)
                    TotalWH += vRowSchBar_Reserve.total_work_hour
                    TotalProduce += vRowSchBar_Reserve.qty_produce

                    '1. tentukan start time berdasarkan end time proses sebelumnya
                    If rowItemWO.Identifier <> IdentifierWODummy Then
                        StartTime = Me.GetStartTime_ByDependency(StartTime, TipeDependency, DelayDependency, vTBLSchBarDummy, vTBLSchBarDummy2, rowItemWO, vTBLItemWO)
                        IdentifierWODummy = rowItemWO.Identifier
                    End If
                    '2. tentukan start time berdasarkan firm planned dan work unit
                    StartTime1 = GetStartTime_ByWorkUnit(vTBLSchBarDummy, vRowSchBar_Reserve, FirmPlanDate)
                    If StartTime1 > StartTime Then
                        StartTime = StartTime1
                    End If

                    TBLSchItem.Rows.Clear()
                    Me.CreateRowItemSch(vRowSchBar_Reserve, TBLSchItem, a, r, g, b, vIdentifierJO, rowItemWO, StartTime, vTBLSchBarDummy2)

                    'copykan data start time, end time pada item wo sch
                    For Each RowSchItem In vTBLItemWO_ItemSch.Rows
                        If RowSchItem.identifier_wo = rowItemWO.Identifier And _
                        RowSchItem.primary_work_unit = vRowSchBar_Reserve.primary_work_unit Then
                            RowSchItem2 = TBLSchItem.Rows(0)

                            RowSchItem.plan_start_time = RowSchItem2.plan_start_time
                            RowSchItem.plan_end_time = RowSchItem2.plan_end_time
                            RowSchItem.plan_end_time_minimum = RowSchItem.plan_end_time
                            RowSchItem.plan_prod_rate = vRowSchBar_Reserve.rate
                            RowSchItem.plan_prod_rate_normal = ObjCD.DigitGrouping((RowSchItem.plan_prod_rate * 100) / RowSchItem.plan_prod_rate_efficiency, 2)
                            RowSchItem.plan_work_duration = vRowSchBar_Reserve.total_work_hour

                            If Start1 > RowSchItem.plan_start_time Then
                                Start1 = RowSchItem.plan_start_time
                            End If
                            If End1 < RowSchItem.plan_end_time Then
                                End1 = RowSchItem.plan_end_time
                            End If

                            Exit For
                        End If
                    Next

                    'update variable yang mempengaruhi data item wo
                    'masukkan data rowsch reserve ke tabel sch bar dummy
                    'ubah data: start time, dan end time
                    Me.AllocateSchedule(vTBLSchBarDummy, TBLSchItem, a, r, g, b, rowItemWO, vPrimaryWUConfig)
                Next

                If JmlWUActive = 1 Then
                    rowItemWO.Plan_Prod_Rate = DummyRate
                Else
                    If TotalWH <> 0 Then
                        rowItemWO.Plan_Prod_Rate = Math.Round((TotalProduce * 60) / (TotalWH / JmlWUActive), rowItemWO._Decimal, MidpointRounding.AwayFromZero)
                    Else
                        rowItemWO.Plan_Prod_Rate = Math.Round(0, MidpointRounding.AwayFromZero)
                    End If
                End If

                rowItemWO.Plan_Work_Hour = TotalWH
                rowItemWO.WU_Allocated = JmlWUActive
                'rowItemWO.WU_Available = vTBLWU.Rows.Count

                rowItemWO.Plan_End_Time = End1
                rowItemWO.Plan_Start_Time = Start1
            Next

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": finalizing gantt chart"
            FrmProgres.prg1.Value = 0
            FrmProgres.LblProgress.Text = ""
            Me.CopyTblSchBar(vTBLSchBarDummy, vTBLSchBar)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLSchBarDummy.Rows.Clear()
            vTBLSchBar_Reserve.Rows.Clear()
            vTBLSchBarDummy2.Rows.Clear()
            TBLSchItem.Rows.Clear()
            vTBLSchBarDummy.Dispose()
            vTBLSchBar_Reserve.Dispose()
            vTBLSchBarDummy2.Dispose()
            TBLSchItem.Dispose()
        End Try
    End Sub

    Public Sub Reschedule_JadwalBaru1(ByVal vIdentifierJO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable, _
    ByVal vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
    ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable, ByVal FrmProgres As FrmTrans_ProductionSchedule_Progress, _
    ByVal vStartTimeAcuan As Date, ByVal vPrimaryWUConfig As Double, Optional ByVal vPrimaryWU_Acuan As Double = -1, _
    Optional ByVal vIdentifierWO_Acuan As String = "")
        Dim FirmPlanDate, StartDate, EndDate As Date
        Dim vTBLSchBarDummy As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vTBLSchBarDummy2 As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vTBLSchBar_Reserve As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim vRowSchBar_Reserve As DseTrans_Schedule.tbl_sch_barRow
        Dim vRowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DV As Data.DataRowView
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim StartTime As Date = Today
        Dim StartTime1 As Date = Today
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim DelayDependency As Double = 0
        Dim TBLSchItem As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
        Dim RowSchItem, RowSchItem2 As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Start1, End1 As Date
        Dim JmlWUActive As Integer = 0
        Dim TotalProduce As Decimal = 0
        Dim TotalWH As Decimal = 0
        Dim DummyRate As Decimal = 0
        Dim CurStep As Integer = 0
        Dim TotalStep As Integer = 5
        Dim IdentifierWODummy As String = ""

        Try
            Start1 = Date.MaxValue
            End1 = Date.MinValue

            'pindahkan semua baris atas identifier jo ybs
            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": duplicating old data"
            vTBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + vIdentifierJO + "' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            FrmProgres.prg1.Value = 0
            FrmProgres.prg1.Maximum = vTBLSchBar.DefaultView.Count
            For Each DV In vTBLSchBar.DefaultView
                vRowSchBar = DV.Row
                FrmProgres.prg1.Value += 1
                FrmProgres.LblProgress.Text = "Data " + FrmProgres.prg1.Value.ToString + " of " + FrmProgres.prg1.Maximum.ToString
                vRowSchBar_Reserve = vTBLSchBar_Reserve.Newtbl_sch_barRow
                For ij As Integer = 0 To vTBLSchBar_Reserve.Columns.Count - 1
                    vRowSchBar_Reserve.Item(ij) = vRowSchBar.Item(ij)
                Next
                vTBLSchBar_Reserve.Rows.Add(vRowSchBar_Reserve)
            Next

            'buat tabel sch baru yang siap untuk alokasi ulang
            FirmPlanDate = Me.GetFirmPlan
            StartDate = FirmPlanDate + TimeSpan.FromDays(1)
            EndDate = Me.GetLast_WS_AllocatedDate
            If Me.GetLast_WS_OvertimeDate.Date > EndDate Then
                EndDate = Me.GetLast_WS_OvertimeDate.Date
            End If

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": preparing gantt chart"
            FrmProgres.prg1.Value = 0
            Me.CopyTblSchBar(vTBLSchBarKosong, vTBLSchBarDummy)

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": preparing reference gantt chart"
            FrmProgres.prg1.Value = 0
            Me.CopyTblSchBar(vTBLSchBarTanpaWO, vTBLSchBarDummy2)
            'Me.LoadScheduleBar(StartDate, EndDate, vTBLSchBarDummy, Nothing)
            'Me.LoadScheduleBar_TanpaScheduledItem(StartDate, EndDate, vTBLSchBarDummy2) 'ini masih bisa optimasi lagi kalo emang berasa lambat

            'alokasikan ulang jadwal berdasarkan urutan level
            'basis alokasi adalah: qty, duration, tipe dependency, dan delay dependency
            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": rescheduling data"
            FrmProgres.prg1.Value = 0
            FrmProgres.prg1.Maximum = vTBLItemWO.Rows.Count
            For Each rowItemWO In vTBLItemWO.Rows
                FrmProgres.prg1.Value += 1
                FrmProgres.LblProgress.Text = "Data " + FrmProgres.prg1.Value.ToString + " of " + FrmProgres.prg1.Maximum.ToString + " - Process: " + rowItemWO.Process_Code + " - Code No: " + rowItemWO.Product_Code

                Start1 = Date.MaxValue
                End1 = Date.MinValue
                JmlWUActive = 0
                TotalWH = 0
                TotalProduce = 0

                vTBLSchBar_Reserve.DefaultView.RowFilter = "[identifier wo] = '" + rowItemWO.Identifier + "'"
                vTBLSchBar_Reserve.DefaultView.Sort = "[plan start time]"
                TipeDependency = rowItemWO.tipe_dependency
                DelayDependency = rowItemWO.delay_dependency

                '1. tentukan start time berdasarkan end time proses sebelumnya
                StartTime = #1/1/1950#

                For Each DV In vTBLSchBar_Reserve.DefaultView
                    JmlWUActive += 1

                    vRowSchBar_Reserve = DV.Row
                    DummyRate = vRowSchBar_Reserve.rate
                    TotalWH += vRowSchBar_Reserve.total_work_hour
                    TotalProduce += vRowSchBar_Reserve.qty_produce

                    If vRowSchBar_Reserve.plan_start_time >= vStartTimeAcuan Then
                        '1. tentukan start time berdasarkan end time proses sebelumnya
                        If rowItemWO.Identifier <> IdentifierWODummy Then
                            StartTime = Me.GetStartTime_ByDependency(StartTime, TipeDependency, DelayDependency, vTBLSchBarDummy, vTBLSchBarDummy2, rowItemWO, vTBLItemWO)
                            IdentifierWODummy = rowItemWO.Identifier
                        End If
                        '2. tentukan start time berdasarkan firm planned dan work unit
                        StartTime1 = GetStartTime_ByWorkUnit(vTBLSchBarDummy, vRowSchBar_Reserve, FirmPlanDate)
                        If StartTime1 > StartTime Then
                            StartTime = StartTime1
                        End If
                    Else
                        '3. start time tidak berubah
                        StartTime = vRowSchBar_Reserve.plan_start_time
                    End If

                    TBLSchItem.Rows.Clear()
                    If vPrimaryWU_Acuan = -1 Then
                        Me.CreateRowItemSch(vRowSchBar_Reserve, TBLSchItem, a, r, g, b, vIdentifierJO, rowItemWO, StartTime, vTBLSchBarDummy2)
                    Else
                        If vPrimaryWU_Acuan = vRowSchBar_Reserve.primary_work_unit And vIdentifierWO_Acuan = vRowSchBar_Reserve.identifier_wo Then
                            Me.CreateRowItemSch(vRowSchBar_Reserve, TBLSchItem, a, r, g, b, vIdentifierJO, rowItemWO, vRowSchBar_Reserve.plan_start_time, vTBLSchBarDummy2)
                        Else
                            Me.CreateRowItemSch(vRowSchBar_Reserve, TBLSchItem, a, r, g, b, vIdentifierJO, rowItemWO, StartTime, vTBLSchBarDummy2)
                        End If
                    End If

                    'copykan data start time, end time pada item wo sch
                    For Each RowSchItem In vTBLItemWO_ItemSch.Rows
                        If RowSchItem.identifier_wo = rowItemWO.Identifier And _
                        RowSchItem.primary_work_unit = vRowSchBar_Reserve.primary_work_unit Then
                            RowSchItem2 = TBLSchItem.Rows(0)

                            RowSchItem.plan_start_time = RowSchItem2.plan_start_time
                            RowSchItem.plan_end_time = RowSchItem2.plan_end_time
                            RowSchItem.plan_end_time_minimum = RowSchItem.plan_end_time

                            If Start1 > RowSchItem.plan_start_time Then
                                Start1 = RowSchItem.plan_start_time
                            End If
                            If End1 < RowSchItem.plan_end_time Then
                                End1 = RowSchItem.plan_end_time
                            End If

                            Exit For
                        End If
                    Next

                    'update variable yang mempengaruhi data item wo
                    'masukkan data rowsch reserve ke tabel sch bar dummy
                    'ubah data: start time, dan end time
                    Me.AllocateSchedule(vTBLSchBarDummy, TBLSchItem, a, r, g, b, rowItemWO, vPrimaryWUConfig)
                Next

                If JmlWUActive = 1 Then
                    rowItemWO.Plan_Prod_Rate = DummyRate
                Else
                    If TotalWH <> 0 Then
                        rowItemWO.Plan_Prod_Rate = Math.Round((TotalProduce * 60) / (TotalWH / JmlWUActive), rowItemWO._Decimal, MidpointRounding.AwayFromZero)
                    Else
                        rowItemWO.Plan_Prod_Rate = Math.Round(0, MidpointRounding.AwayFromZero)
                    End If
                End If
                rowItemWO.Plan_Work_Hour = TotalWH
                rowItemWO.WU_Allocated = JmlWUActive
                'rowItemWO.WU_Available = vTBLWU.Rows.Count

                rowItemWO.Plan_End_Time = End1
                rowItemWO.Plan_Start_Time = Start1
            Next

            CurStep += 1
            FrmProgres.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": finalizing gantt chart"
            FrmProgres.prg1.Value = 0
            FrmProgres.LblProgress.Text = ""
            Me.CopyTblSchBar(vTBLSchBarDummy, vTBLSchBar)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLSchBarDummy.Rows.Clear()
            vTBLSchBar_Reserve.Rows.Clear()
            vTBLSchBarDummy2.Rows.Clear()
            TBLSchItem.Rows.Clear()
            vTBLSchBarDummy.Dispose()
            vTBLSchBar_Reserve.Dispose()
            vTBLSchBarDummy2.Dispose()
            TBLSchItem.Dispose()
        End Try
    End Sub

    Private Sub CreateRowItemSch(ByVal vRowSchBar_Reserve As DseTrans_Schedule.tbl_sch_barRow, _
    ByVal TBLSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable, _
    ByVal va As Integer, ByVal vr As Integer, ByVal vg As Integer, ByVal vb As Integer, ByVal vIdentifierJO As String, _
    ByVal rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal StartTime As Date, _
    ByVal vTBLSchBarDummy2 As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim RowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            RowSchItem = TBLSchItem.Newtbltrans_mo_item_wo_item_schRow
            With RowSchItem
                'disini argb diset no dulu, nanti pas alocate baru di set warna yg sebenarnya
                .a = va
                .r = vr
                .g = vg
                .b = vb
                'kolom2 yang ga dipake di set ke nilai default dulu, disini yang penting itungannya
                .actual_down_time_duration = 0
                .actual_end_time = #1/1/1950#
                .actual_prod_rate = 0
                .actual_setup_duration = 0
                .actual_start_time = #1/1/1950#
                .actual_total_duration = 0
                .actual_work_duration = 0
                .plan_setup_duration = 0
                .plan_total_duration = 0
                .plan_downtime = 0
                .primary = -1
                .status = 0
                .progress = 0
                .plan_prod_rate_normal = 0
                .plan_prod_rate_efficiency = 0
                .plan_result = 0
                .plan_good_percentage = 0
                .tipe_dependency = 0
                .delay_dependency = 0

                .identifier_jo = vIdentifierJO
                .identifier_wo = rowItemWO.Identifier
                .primary_mo = rowItemWO.Primary_MO
                .primary_process = vRowSchBar_Reserve.primary_process

                'berikut data yang dipake untuk perhitungan
                .primary_work_unit = vRowSchBar_Reserve.primary_work_unit
                .Primary_Unit = vRowSchBar_Reserve.primary_unit
                .plan_prod_rate = vRowSchBar_Reserve.rate
                .plan_start_time = StartTime
                .plan_end_time = Me.GetEndTime(vRowSchBar_Reserve.primary_work_unit, .plan_start_time, vRowSchBar_Reserve.total_work_hour, vTBLSchBarDummy2)
                .plan_end_time_minimum = .plan_end_time
                .plan_current_time = .plan_start_time
                .plan_current_time2 = .plan_start_time
                .qty_produce = vRowSchBar_Reserve.qty_produce
                .qty_produce_add = 0
                .plan_work_duration = vRowSchBar_Reserve.total_work_hour
                .plan_work_duration_add = 0
                ._Decimal = vRowSchBar_Reserve.decimal_point
                .primary_wu_config = vRowSchBar_Reserve.primary_wu_config
                .primary_input_location = vRowSchBar_Reserve.primary_input_location
                .primary_output_location = vRowSchBar_Reserve.primary_output_location
            End With
            TBLSchItem.Rows.Add(RowSchItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetStartTime_ByDependency(ByVal vStartTime As Date, ByVal vTipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency, _
    ByVal vDelayDependency As Double, ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBarDummy2 As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vRowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row, _
    ByVal vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable) As Date
        Dim vRowSchBar_Dummy As DseTrans_Schedule.tbl_sch_barRow
        Dim EndTime As Date = Today

        Try
            'cek proses sebelumnya hanya jika bergantung pada proses sebelumnya
            vTBLSchBarDummy.DefaultView.RowFilter = GetFilterString_IdentifierWOParent(vRowItemWO.Identifier, vTBLItemWO)
            If vTBLSchBarDummy.DefaultView.Count > 0 Then
                If vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1 Then
                    'cari row sch berdasarkan identifier wo, yang plan end datenya paling bontot
                    vTBLSchBarDummy.DefaultView.Sort = "[plan end time]"
                    vRowSchBar_Dummy = vTBLSchBarDummy.DefaultView(vTBLSchBarDummy.DefaultView.Count - 1).Row

                    If vDelayDependency = 0 Then
                        If vRowSchBar_Dummy.plan_end_time > vStartTime Then
                            vStartTime = vRowSchBar_Dummy.plan_end_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(vRowSchBar_Dummy.primary_work_unit, vRowSchBar_Dummy.plan_end_time, vDelayDependency, vTBLSchBarDummy2)
                        If EndTime > vStartTime Then
                            vStartTime = EndTime
                        End If
                    End If
                ElseIf vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start1 Then
                    vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                    vRowSchBar_Dummy = vTBLSchBarDummy.DefaultView(vTBLSchBarDummy.DefaultView.Count - 1).Row

                    If vDelayDependency = 0 Then
                        If vRowSchBar_Dummy.plan_start_time > vStartTime Then
                            vStartTime = vRowSchBar_Dummy.plan_start_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(vRowSchBar_Dummy.primary_work_unit, vRowSchBar_Dummy.plan_start_time, vDelayDependency, vTBLSchBarDummy2)
                        If EndTime > vStartTime Then
                            vStartTime = EndTime
                        End If
                        If vStartTime > vRowSchBar_Dummy.plan_end_time Then
                            vStartTime = vRowSchBar_Dummy.plan_end_time
                        End If
                    End If
                ElseIf vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start_Strict1 Then
                    vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                    vRowSchBar_Dummy = vTBLSchBarDummy.DefaultView(vTBLSchBarDummy.DefaultView.Count - 1).Row

                    If vDelayDependency = 0 Then
                        If vRowSchBar_Dummy.plan_start_time > vStartTime Then
                            vStartTime = vRowSchBar_Dummy.plan_start_time
                        End If
                    Else
                        EndTime = Me.GetEndTime(vRowSchBar_Dummy.primary_work_unit, vRowSchBar_Dummy.plan_start_time, vDelayDependency, vTBLSchBarDummy2)
                        If EndTime > vStartTime Then
                            vStartTime = EndTime
                        End If
                    End If
                End If
            End If

            Return vStartTime
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetStartTime_ByWorkUnit(ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vRowSchBar_Reserve As DseTrans_Schedule.tbl_sch_barRow, ByVal FirmPlanDate As Date) As Date
        Dim StartTime As Date
        Dim vRowSchBar As DseTrans_Schedule.tbl_sch_barRow

        Try
            '0. dari firm plan
            StartTime = FirmPlanDate
            '1. dari end time work unit tersebut
            vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar_Reserve.primary_work_unit.ToString + _
                                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            vTBLSchBarDummy.DefaultView.Sort = "[plan end time]"
            If vTBLSchBarDummy.DefaultView.Count = 0 Then
                'belum ada work time, maka start time, ambil sch bar pertama setelah firm plan
                vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar_Reserve.primary_work_unit.ToString + _
                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
                vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                For ij As Integer = 0 To vTBLSchBarDummy.DefaultView.Count
                    vRowSchBar = vTBLSchBarDummy.DefaultView(ij).Row
                    If vRowSchBar.plan_start_time.Date > FirmPlanDate Then
                        StartTime = vRowSchBar.plan_start_time
                        Exit For
                    End If
                Next
            Else
                vRowSchBar = vTBLSchBarDummy.DefaultView(vTBLSchBarDummy.DefaultView.Count - 1).Row
                If vRowSchBar.plan_end_time > StartTime Then
                    StartTime = vRowSchBar.plan_end_time
                Else
                    'work time terakhir sebelum firm plan, maka start time, ambil sch bar pertama setelah firm plan
                    vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vRowSchBar_Reserve.primary_work_unit.ToString + _
                    " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
                    vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                    For ij As Integer = 0 To vTBLSchBarDummy.DefaultView.Count
                        vRowSchBar = vTBLSchBarDummy.DefaultView(0).Row
                        If vRowSchBar.plan_start_time.Date > FirmPlanDate Then
                            StartTime = vRowSchBar.plan_start_time
                            Exit For
                        End If
                    Next
                End If
            End If

            Return StartTime
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetStartTime_ByWorkUnit(ByVal vTBLSchBarDummy As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vPrimaryWorkUnit As Double, ByVal FirmPlanDate As Date) As Date
        Dim StartTime As Date
        Dim vRowSchBar As DseTrans_Schedule.tbl_sch_barRow

        Try
            '0. dari firm plan
            StartTime = FirmPlanDate
            '1. dari end time work unit tersebut
            vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWorkUnit.ToString + _
                                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            vTBLSchBarDummy.DefaultView.Sort = "[plan end time]"
            If vTBLSchBarDummy.DefaultView.Count = 0 Then
                'belum ada work time, maka start time, ambil sch bar pertama setelah firm plan
                vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWorkUnit.ToString + _
                " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
                vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                For ij As Integer = 0 To vTBLSchBarDummy.DefaultView.Count
                    vRowSchBar = vTBLSchBarDummy.DefaultView(ij).Row
                    If vRowSchBar.plan_start_time.Date > FirmPlanDate Then
                        StartTime = vRowSchBar.plan_start_time
                        Exit For
                    End If
                Next
            Else
                vRowSchBar = vTBLSchBarDummy.DefaultView(vTBLSchBarDummy.DefaultView.Count - 1).Row
                If vRowSchBar.plan_end_time > StartTime Then
                    StartTime = vRowSchBar.plan_end_time
                Else
                    'work time terakhir sebelum firm plan, maka start time, ambil sch bar pertama setelah firm plan
                    vTBLSchBarDummy.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWorkUnit.ToString + _
                    " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time).ToString
                    vTBLSchBarDummy.DefaultView.Sort = "[plan start time]"
                    For ij As Integer = 0 To vTBLSchBarDummy.DefaultView.Count
                        vRowSchBar = vTBLSchBarDummy.DefaultView(0).Row
                        If vRowSchBar.plan_start_time.Date > FirmPlanDate Then
                            StartTime = vRowSchBar.plan_start_time
                            Exit For
                        End If
                    Next
                End If
            End If

            Return StartTime
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub HapusJadwal(ByVal vIdentifierJO As String, ByVal vFrmProgress As FrmTrans_ProductionSchedule_Progress3)
        Dim DTARef As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch1TableAdapter
        Dim TBLRef As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch1DataTable
        Dim rowRef As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch1Row
        Dim StartModif As Boolean = False
        Dim TBLSchBarTanpaWO As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TBLSchBarTanpaWO1 As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim RowSchBar, RowSchBar1 As DseTrans_Schedule.tbl_sch_barRow
        Dim StartDate As Date = Date.MaxValue
        Dim EndDate As Date = Today
        Dim DTASch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch2TableAdapter
        Dim TBLSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch2DataTable
        Dim TBLSch1 As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch2DataTable
        Dim rowSch, rowSch1 As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch2Row
        Dim DV As Data.DataRowView
        Dim TotalStep As Integer = 3
        Dim CurStep As Integer = 0
        Dim DescLama As String = ""
        Dim CurIdx As Integer = 0
        Dim IsFirst As Boolean = True
        Dim StartDateByDependency As Date = #1/1/1950#
        Dim DTAWOItem As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo2TableAdapter
        Dim DTAWOItem1 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim TBLWOItem As New DseTrans_Schedule.tbltrans_mo_item_wo2DataTable
        'Dim rowWOItem As DseTrans_Schedule.tbltrans_mo_item_wo2Row
        Dim AddDuration As Double = 0
        Dim EndOfWU As Date = #1/1/1950#
        Dim FirmPlan As Date = #1/1/1950#
        Dim DTAWIPDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wip_detailTableAdapter
        Dim DTARMDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_detailTableAdapter

        Try
            DTARef.Connection.ConnectionString = Core_Module.GetConstring
            DTASch.Connection.ConnectionString = DTARef.Connection.ConnectionString
            DTAWOItem.Connection.ConnectionString = DTARef.Connection.ConnectionString
            DTAWOItem1.Connection.ConnectionString = DTARef.Connection.ConnectionString
            DTAWIPDetail.Connection.ConnectionString = DTARef.Connection.ConnectionString
            DTARMDetail.Connection.ConnectionString = DTARef.Connection.ConnectionString

            FirmPlan = Me.GetFirmPlan

            'hapus item wip detail dan rm detail
            DTAWIPDetail.Delete(vIdentifierJO)
            DTARMDetail.DeleteQuery(vIdentifierJO)

            'load item wo sch untuk identifier wo ybs
            'catatkan: primary wu, start time, durasi geser
            CurStep += 1
            vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": reloading gantt chart"
            DTARef.Fill(TBLRef, vIdentifierJO)

            'load gantt chart kosong
            For Each rowRef In TBLRef.Rows
                If StartDate.Date > rowRef.Plan_Start_Time.Date Then
                    StartDate = rowRef.Plan_Start_Time.Date
                End If
            Next
            EndDate = Me.GetLast_WS_AllocatedDate
            Me.LoadScheduleBar_TanpaScheduledItem(StartDate - TimeSpan.FromDays(2), EndDate, TBLSchBarTanpaWO)

            'untuk setiap primary wu, mulai dari item ybs ke bawah, geser sejumlah durasi ybs ke kiri
            CurStep += 1
            CurIdx = 0
            vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": rescheduling"
            DescLama = vFrmProgress.LblDescription.Text
            DTASch.FillBy(TBLSch, StartDate - TimeSpan.FromDays(7))
            TBLSch1.Rows.Clear()
            For Each rowSch In TBLSch.Rows
                rowSch1 = TBLSch1.Newtbltrans_mo_item_wo_item_sch2Row

                For ij As Integer = 0 To TBLSch.Columns.Count - 1
                    rowSch1.Item(ij) = rowSch.Item(ij)
                Next

                TBLSch1.Rows.Add(rowSch1)
            Next
            TBLSch.DefaultView.Sort = "[plan start time]"
            For Each rowRef In TBLRef.Rows
                CurIdx += 1
                vFrmProgress.LblDescription.Text = DescLama + " - work unit " + CurIdx.ToString + " of " + TBLRef.Rows.Count.ToString + " (" + rowRef.work_unit_code + ")"

                TBLSchBarTanpaWO1.Rows.Clear()
                TBLSchBarTanpaWO.DefaultView.RowFilter = "[primary work unit] = " + rowRef.primary_work_unit.ToString
                For Each DV In TBLSchBarTanpaWO.DefaultView
                    RowSchBar = DV.Row
                    RowSchBar1 = TBLSchBarTanpaWO1.Newtbl_sch_barRow
                    For ij As Integer = 0 To TBLSchBarTanpaWO.Columns.Count - 1
                        RowSchBar1.Item(ij) = RowSchBar.Item(ij)
                    Next
                    TBLSchBarTanpaWO1.Rows.Add(RowSchBar1)
                Next

                StartModif = False
                IsFirst = True
                AddDuration = 0
                EndOfWU = #1/1/1950#

                TBLSch.DefaultView.RowFilter = "[primary work unit] = " + rowRef.primary_work_unit.ToString
                vFrmProgress.prg1.Value = 0
                vFrmProgress.prg1.Maximum = TBLSch.DefaultView.Count
                For Each DV In TBLSch.DefaultView
                    vFrmProgress.prg1.Value += 1
                    vFrmProgress.LblProgress.Text = "Data " + vFrmProgress.prg1.Value.ToString + " of " + vFrmProgress.prg1.Maximum.ToString
                    rowSch = DV.Row

                    If rowSch.identifier_jo = vIdentifierJO Then
                        StartModif = True
                    Else
                        If StartModif = False Then
                            EndOfWU = rowSch.plan_end_time
                        End If

                        If StartModif = True Then
                            If rowSch.identifier_jo <> vIdentifierJO Then
                                StartDate = Me.GetStartTime(rowRef.primary_work_unit, rowSch.plan_start_time, rowRef.Duration, TBLSchBarTanpaWO1)
                                If IsFirst = True Then
                                    IsFirst = False
                                    'get tanggal start jika dibandingkan dengan proses sebelumnya
                                    TBLWOItem.Rows.Clear()
                                    DTAWOItem.Fill(TBLWOItem, rowSch.identifier_jo)

                                    StartDateByDependency = Me.GetStartTime(vIdentifierJO, rowSch.identifier_wo, rowRef.primary_work_unit, TBLSchBarTanpaWO, TBLWOItem, TBLSch1)
                                    If StartDateByDependency <> #1/1/1950# Then
                                        If StartDateByDependency < StartDate And StartDateByDependency > EndOfWU Then
                                            AddDuration = Me.GetWorkDuration(StartDateByDependency, StartDate, rowRef.primary_work_unit, TBLSchBarTanpaWO1)
                                        End If
                                    End If
                                End If

                                If rowRef.Duration + AddDuration < rowSch.plan_work_duration Then
                                    EndDate = Me.GetStartTime(rowRef.primary_work_unit, rowSch.plan_end_time, rowRef.Duration + AddDuration, TBLSchBarTanpaWO1)
                                Else
                                    EndDate = Me.GetEndTime(rowRef.primary_work_unit, StartDate, rowSch.plan_work_duration, TBLSchBarTanpaWO1)
                                End If

                                DTASch.UpdateQuery_SetJadwalBaru(StartDate, EndDate, rowSch.primary)
                            End If
                        End If
                    End If
                Next
            Next
            vFrmProgress.LblProgress.Text = ""
            vFrmProgress.prg1.Value = 0

            'hapus item wo sch
            CurStep += 1
            vFrmProgress.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + ": deleting schedule"
            DTASch.DeleteQuery_DeleteByIdentifierJO(vIdentifierJO)

            'update status atas semua item wo
            DTAWOItem1.UpdateQuery_SetStatus(CInt(ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned), vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTARef.Dispose()
            TBLRef.Dispose()
            DTASch.Dispose()
            TBLSch.Dispose()
            DTAWOItem.Dispose()
            TBLWOItem.Dispose()
            TBLSch1.Dispose()
            TBLSchBarTanpaWO1.Dispose()
        End Try
    End Sub

    Public Sub OptimizeSchedule(ByVal vIdentifierJO As String, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
    ByVal vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusJadwal(ByVal vIdentifierJO As String)
        Dim ij As New FrmTrans_ProductionSchedule_Progress3

        Try
            ij.ObjSch = Me
            ij.vIdentifierJO = vIdentifierJO
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub GeserJadwal(ByVal vPrimaryWU As Double, ByVal vDurasiGeser As Double, _
    ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierWOAcuan As String, ByVal vStartDateAcuan As Date, _
    ByVal vIdentifierJOAcuan As String, ByVal vFrm As FrmTrans_ProductionSchedule_Progress4)
        Dim TBLSchBarKosong As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim DTAItemWOSch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
        Dim DTAItemWOSch2 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch2TableAdapter
        Dim TBLItemWOSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4DataTable
        Dim rowItemWOSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4Row
        Dim StartUpdate As Boolean = False
        Dim EndTime As Date = Today
        Dim StartDate, EndDate As Date
        Dim TBLRef As New DseTrans_Schedule.Tbl_RefDataTable
        Dim rowRef As DseTrans_Schedule.Tbl_RefRow
        Dim Ketemu As Boolean = False

        Dim TotalStep As Integer = 3
        Dim CurStep As Integer = 0

        Try
            CurStep += 1
            vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Reloading gantt chart"
            vFrm.prg1.Value = 0
            vFrm.LblProgress.Text = ""

            DTAItemWOSch.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemWOSch2.Connection.ConnectionString = DTAItemWOSch.Connection.ConnectionString

            'load data sch bar
            Me.LoadScheduleBar_TanpaScheduledItem(vStartDateAcuan, #1/1/1950#, TBLSchBarKosong, vPrimaryWU)

            'load item sch wo, berdasarkan primary wu, dan kemudian urutkan berdasarkan plan start time
            DTAItemWOSch.Fill(TBLItemWOSch, (vStartDateAcuan - TimeSpan.FromDays(1)).Date, vPrimaryWU)

            'untuk setiap baris, tentukan start time dan end time, kemudian simpan ke database
            CurStep += 1
            vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Rescheduling"
            vFrm.prg1.Value = 0
            vFrm.LblProgress.Text = ""
            vFrm.prg1.Maximum = TBLItemWOSch.Rows.Count
            For Each rowItemWOSch In TBLItemWOSch
                vFrm.prg1.Value += 1
                vFrm.LblProgress.Text = "Data " + vFrm.prg1.Value.ToString + " of " + vFrm.prg1.Maximum.ToString + " - " + Format(rowItemWOSch.plan_start_time, "dd MMM yyyy")

                If StartUpdate = True Then
                    If vDurasiGeser > 0 Then
                        'geser ke kanan
                        rowItemWOSch.plan_start_time = Me.GetEndTime(vPrimaryWU, rowItemWOSch.plan_start_time, vDurasiGeser, TBLSchBarKosong)
                    Else
                        'geser ke kiri
                        rowItemWOSch.plan_start_time = Me.GetStartTime(vPrimaryWU, rowItemWOSch.plan_start_time, vDurasiGeser * -1, TBLSchBarKosong)
                    End If
                    EndTime = Me.GetEndTime(vPrimaryWU, rowItemWOSch.plan_start_time, rowItemWOSch.plan_work_duration, vTBLSchBarKosong)

                    'update ke db
                    DTAItemWOSch2.UpdateQuery_SetJadwalBaru(rowItemWOSch.plan_start_time, EndTime, rowItemWOSch.primary)
                    'simpan identifier wo, untuk update start time + end time
                    Ketemu = False
                    For Each rowRef In TBLRef.Rows
                        If rowRef.identifier_jo = rowItemWOSch.identifier_jo And rowRef.identifier_wo = rowItemWOSch.identifier_wo Then
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        rowRef = TBLRef.NewTbl_RefRow
                        rowRef.identifier_wo = rowItemWOSch.identifier_wo
                        rowRef.identifier_jo = rowItemWOSch.identifier_jo
                        TBLRef.Rows.Add(rowRef)
                    End If
                Else
                    If rowItemWOSch.identifier_wo = vIdentifierWOAcuan And rowItemWOSch.identifier_jo = vIdentifierJOAcuan Then
                        StartUpdate = True
                    End If
                End If
            Next

            'untuk setiap idetifier wo, tentukan start time dan end time
            CurStep += 1
            vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Updating database"
            vFrm.prg1.Value = 0
            vFrm.LblProgress.Text = ""
            vFrm.prg1.Maximum = TBLRef.Rows.Count
            For Each rowRef In TBLRef.Rows
                vFrm.prg1.Value += 1
                vFrm.LblProgress.Text = "Data " + vFrm.prg1.Value.ToString + " of " + vFrm.prg1.Maximum.ToString

                TBLItemWOSch.Rows.Clear()
                StartDate = Date.MaxValue
                EndDate = Date.MinValue

                DTAItemWOSch.FillBy_IdentifierWO_IdentifierJO(TBLItemWOSch, rowRef.identifier_wo, rowRef.identifier_jo)
                For Each rowItemWOSch In TBLItemWOSch.Rows
                    If StartDate > rowItemWOSch.plan_start_time Then
                        StartDate = rowItemWOSch.plan_start_time
                    End If
                    If EndDate < rowItemWOSch.plan_end_time Then
                        EndDate = rowItemWOSch.plan_end_time
                    End If
                Next

                DTAItemWOSch.UpdateQuery(StartDate, EndDate, rowRef.identifier_wo, rowRef.identifier_jo)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLSchBarKosong.Dispose()
            DTAItemWOSch.Dispose()
            TBLItemWOSch.Dispose()
            TBLRef.Dispose()
            DTAItemWOSch2.Dispose()
        End Try
    End Sub

    Public Sub GeserJadwal(ByVal vPrimaryWU As Double, ByVal vDurasiGeser As Double, _
    ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierWOAcuan As String, ByVal vStartDateAcuan As Date, _
    ByVal vIdentifierJOAcuan As String)
        Dim Frm As New FrmTrans_ProductionSchedule_Progress4

        Try
            Frm.PrimaryWU = vPrimaryWU
            Frm.DurasiGeser = vDurasiGeser
            Frm.TBLSchBarKosong = vTBLSchBarKosong
            Frm.IdentifierWOAcuan = vIdentifierWOAcuan
            Frm.StartDateAcuan = vStartDateAcuan
            Frm.IdentifierJOAcuan = vIdentifierJOAcuan
            Frm.ObjSch = Me

            Frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub

    ' Public Sub GeserJadwal_EndTime(ByVal vPrimaryWU As Double, ByVal vDurasiGeser As Double, _
    'ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierWOAcuan As String, ByVal vStartDateAcuan As Date, _
    'ByVal vIdentifierJOAcuan As String)
    '     Dim Frm As New FrmTrans_ProductionSchedule_Progress5

    '     Try
    '         Frm.PrimaryWU = vPrimaryWU
    '         Frm.DurasiGeser = vDurasiGeser
    '         Frm.TBLSchBarKosong = vTBLSchBarKosong
    '         Frm.IdentifierWOAcuan = vIdentifierWOAcuan
    '         Frm.StartDateAcuan = vStartDateAcuan
    '         Frm.IdentifierJOAcuan = vIdentifierJOAcuan
    '         Frm.ObjSch = Me

    '         Frm.ShowDialog()
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         Frm.Dispose()
    '     End Try
    ' End Sub

    ' Public Sub GeserJadwal_EndTime(ByVal vPrimaryWU As Double, ByVal vDurasiGeser As Double, _
    ' ByVal vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierWOAcuan As String, ByVal vStartDateAcuan As Date, _
    ' ByVal vIdentifierJOAcuan As String, ByVal vFrm As FrmTrans_ProductionSchedule_Progress5)
    '     Dim TBLSchBarKosong As New DseTrans_Schedule.tbl_sch_barDataTable
    '     Dim DTAItemWOSch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
    '     Dim DTAItemWOSch2 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch2TableAdapter
    '     Dim TBLItemWOSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4DataTable
    '     Dim rowItemWOSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4Row
    '     Dim StartUpdate As Boolean = False
    '     Dim EndTime As Date = Today
    '     Dim StartDate, EndDate As Date
    '     Dim TBLRef As New DseTrans_Schedule.Tbl_RefDataTable
    '     Dim rowRef As DseTrans_Schedule.Tbl_RefRow
    '     Dim Ketemu As Boolean = False

    '     Dim TotalStep As Integer = 3
    '     Dim CurStep As Integer = 0

    '     Try
    '         CurStep += 1
    '         vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Reloading gantt chart"
    '         vFrm.prg1.Value = 0
    '         vFrm.LblProgress.Text = ""

    '         DTAItemWOSch.Connection.ConnectionString = Core_Module.GetConstring
    '         DTAItemWOSch2.Connection.ConnectionString = DTAItemWOSch.Connection.ConnectionString

    '         'load data sch bar
    '         Me.LoadScheduleBar_TanpaScheduledItem(vStartDateAcuan, #1/1/1950#, TBLSchBarKosong, vPrimaryWU)

    '         'load item sch wo, berdasarkan primary wu, dan kemudian urutkan berdasarkan plan start time
    '         DTAItemWOSch.Fill(TBLItemWOSch, (vStartDateAcuan - TimeSpan.FromDays(1)).Date, vPrimaryWU)

    '         'untuk setiap baris, tentukan start time dan end time, kemudian simpan ke database
    '         CurStep += 1
    '         vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Rescheduling"
    '         vFrm.prg1.Value = 0
    '         vFrm.LblProgress.Text = ""
    '         vFrm.prg1.Maximum = TBLItemWOSch.Rows.Count
    '         For Each rowItemWOSch In TBLItemWOSch
    '             vFrm.prg1.Value += 1
    '             vFrm.LblProgress.Text = "Data " + vFrm.prg1.Value.ToString + " of " + vFrm.prg1.Maximum.ToString + " - " + Format(rowItemWOSch.plan_start_time, "dd HHH yyyy HH.mm")

    '             If StartUpdate = True Then
    '                 If vDurasiGeser > 0 Then
    '                     'geser ke kanan
    '                     rowItemWOSch.plan_start_time = Me.GetEndTime(vPrimaryWU, rowItemWOSch.plan_start_time, vDurasiGeser, TBLSchBarKosong)
    '                 Else
    '                     'geser ke kiri
    '                     rowItemWOSch.plan_start_time = Me.GetStartTime(vPrimaryWU, rowItemWOSch.plan_start_time, vDurasiGeser * -1, TBLSchBarKosong)
    '                 End If
    '                 EndTime = Me.GetEndTime(vPrimaryWU, rowItemWOSch.plan_start_time, rowItemWOSch.plan_work_duration, vTBLSchBarKosong)

    '                 'update ke db
    '                 DTAItemWOSch2.UpdateQuery_SetJadwalBaru(rowItemWOSch.plan_start_time, EndTime, rowItemWOSch.primary)
    '                 'simpan identifier wo, untuk update start time + end time
    '                 Ketemu = False
    '                 For Each rowRef In TBLRef.Rows
    '                     If rowRef.identifier_jo = rowItemWOSch.identifier_jo And rowRef.identifier_wo = rowItemWOSch.identifier_wo Then
    '                         Ketemu = True
    '                         Exit For
    '                     End If
    '                 Next
    '                 If Ketemu = False Then
    '                     rowRef = TBLRef.NewTbl_RefRow
    '                     rowRef.identifier_wo = rowItemWOSch.identifier_wo
    '                     rowRef.identifier_jo = rowItemWOSch.identifier_jo
    '                     TBLRef.Rows.Add(rowRef)
    '                 End If
    '             Else
    '                 If rowItemWOSch.identifier_wo = vIdentifierWOAcuan And rowItemWOSch.identifier_jo = vIdentifierJOAcuan Then
    '                     StartUpdate = True
    '                 End If
    '             End If
    '         Next

    '         'untuk setiap idetifier wo, tentukan start time dan end time
    '         CurStep += 1
    '         vFrm.LblDescription.Text = "Step " + CurStep.ToString + " of " + TotalStep.ToString + " - Updating database"
    '         vFrm.prg1.Value = 0
    '         vFrm.LblProgress.Text = ""
    '         vFrm.prg1.Maximum = TBLRef.Rows.Count
    '         For Each rowRef In TBLRef.Rows
    '             vFrm.prg1.Value += 1
    '             vFrm.LblProgress.Text = "Data " + vFrm.prg1.Value.ToString + " of " + vFrm.prg1.Maximum.ToString

    '             TBLItemWOSch.Rows.Clear()
    '             StartDate = Date.MaxValue
    '             EndDate = Date.MinValue

    '             DTAItemWOSch.FillBy_IdentifierWO_IdentifierJO(TBLItemWOSch, rowRef.identifier_wo, rowRef.identifier_jo)
    '             For Each rowItemWOSch In TBLItemWOSch.Rows
    '                 If StartDate > rowItemWOSch.plan_start_time Then
    '                     StartDate = rowItemWOSch.plan_start_time
    '                 End If
    '                 If EndDate < rowItemWOSch.plan_end_time Then
    '                     EndDate = rowItemWOSch.plan_end_time
    '                 End If
    '             Next

    '             DTAItemWOSch.UpdateQuery(StartDate, EndDate, rowRef.identifier_wo, rowRef.identifier_jo)
    '         Next
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         TBLSchBarKosong.Dispose()
    '         DTAItemWOSch.Dispose()
    '         TBLItemWOSch.Dispose()
    '         TBLRef.Dispose()
    '         DTAItemWOSch2.Dispose()
    '     End Try
    ' End Sub

    Public Sub UpdateStartEndTime_ItemWO(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String)
        Dim DTAItemWOSch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
        Dim TBLItemWOSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4DataTable
        Dim rowItemWOSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4Row
        Dim StartDate, EndDate As Date

        Try
            DTAItemWOSch.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = Date.MaxValue
            EndDate = Date.MinValue

            DTAItemWOSch.FillBy_IdentifierWO_IdentifierJO(TBLItemWOSch, vIdentifierWO, vIdentifierJO)
            For Each rowItemWOSch In TBLItemWOSch.Rows
                If StartDate > rowItemWOSch.plan_start_time Then
                    StartDate = rowItemWOSch.plan_start_time
                End If
                If EndDate < rowItemWOSch.plan_end_time Then
                    EndDate = rowItemWOSch.plan_end_time
                End If
            Next

            DTAItemWOSch.UpdateQuery(StartDate, EndDate, vIdentifierWO, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemWOSch.Dispose()
        End Try
    End Sub

    Public Function GetProdEfficiency(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String) As Integer
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            RetVal = DTA.ScalarQuery_GetProdRateEfficiency(vIdentifierJO, vIdentifierWO)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDecimalPoint(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String) As Integer
        Dim DTA As New DseTrans_ScheduleTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseTrans_Schedule.tblmaster_productDataTable
        Dim row As DseTrans_Schedule.tblmaster_productRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vIdentifierJO, vIdentifierWO)
            If TBL.Rows.Count = 0 Then
                Return 0
            Else
                row = TBL.Rows(0)
                Return row._Decimal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub UpdateStartEndTime_ItemWO2(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String)
        Dim DTAItemWOSch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch4TableAdapter
        Dim TBLItemWOSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4DataTable
        Dim rowItemWOSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch4Row
        Dim StartDate, EndDate As Date
        Dim TotalWH, ProduceQty, RateKumulatif As Decimal

        Try
            TotalWH = 0
            RateKumulatif = 0
            ProduceQty = 0

            DTAItemWOSch.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = Date.MaxValue
            EndDate = Date.MinValue

            DTAItemWOSch.FillBy_IdentifierWO_IdentifierJO(TBLItemWOSch, vIdentifierWO, vIdentifierJO)
            For Each rowItemWOSch In TBLItemWOSch.Rows
                If StartDate > rowItemWOSch.plan_start_time Then
                    StartDate = rowItemWOSch.plan_start_time
                End If
                If EndDate < rowItemWOSch.plan_end_time Then
                    EndDate = rowItemWOSch.plan_end_time
                End If
                TotalWH += rowItemWOSch.plan_work_duration
                ProduceQty += rowItemWOSch.qty_produce
            Next

            DTAItemWOSch.UpdateQuery1(StartDate, EndDate, TotalWH, _
            Math.Round((ProduceQty * 60) / TotalWH, Me.GetDecimalPoint(vIdentifierJO, vIdentifierWO), MidpointRounding.AwayFromZero), vIdentifierWO, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemWOSch.Dispose()
        End Try
    End Sub

    Function UpdatePlanStartEndSchedule(ByVal vNewPlanStart As Date, ByVal vNewPlanEnd As Date, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkunit As Double) As Boolean
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_UpdateSchedule(vNewPlanStart, vNewPlanEnd, vIdentifierJO, vIdentifierWO, vPrimaryWorkunit)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function UpdateInputOutputLocation(ByVal vIdentiferJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkUnit As Double, ByVal vNewPrimaryInputLocation As Double, ByVal vNewPrimaryOutputLocatio As Double) As Boolean
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_InputOutputLocation(vNewPrimaryInputLocation, vNewPrimaryOutputLocatio, vIdentiferJO, vIdentifierWO, vPrimaryWorkUnit) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ChangeWorkUnit(ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkUnit_old As Double, ByVal vPrimaryWorkUnit_new As Double, _
    ByVal vPlanStart As Date, ByVal vPlanEnd As Date, ByVal vPlanWorkDuration As Double, ByVal vProdRate As Decimal, ByVal vEfficiencyRate As Decimal, ByVal vEffectiveRate As Decimal, _
    ByVal vPrimaryInputLocation As Double, ByVal vPrimaryOutputLocation As Double) As Boolean
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_ChangeWorkUnit(vPrimaryWorkUnit_new, vPlanStart, vPlanEnd, vPlanWorkDuration, vProdRate, vEfficiencyRate, vEffectiveRate, vPrimaryInputLocation, vPrimaryOutputLocation, vPrimaryMO, _
            vIdentifierJO, vIdentifierWO, vPrimaryWorkUnit_old) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetProgress(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkUnit As Double) As Decimal
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim progress As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            progress = dta.ScalarQuery_GetProgress(vIdentifierJO, vIdentifierWO, vPrimaryWorkUnit)
            If progress.HasValue Then
                Return progress.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Function UpdateScheduleAndQty(ByVal vNewPlanStart As Date, ByVal vNewPlanEnd As Date, ByVal vQtyProduce As Decimal, ByVal vTotalWH As Double, _
    ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkunit As Double) As Integer
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Return dta.UpdateQuery_UpdateScheduleAndQty(vNewPlanStart, vNewPlanEnd, vQtyProduce, vTotalWH, vTotalWH, vIdentifierJO, vIdentifierWO, vPrimaryWorkunit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function DeleteZeroQty(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkunit As Double) As Integer
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Return dta.DeleteQuery_ZeroQty(vIdentifierJO, vIdentifierWO, vPrimaryWorkunit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function InsertNewSchedule(ByVal vRowSchBar As DseTrans_Schedule.tbl_sch_barRow) As Boolean
        Dim dta As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim DelayDependency As Double = 0
        Dim objWorkunit As New Core_Module.ClsMasterWorkUnit
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.GetProductDependency(vRowSchBar.primary_product, TipeDependency, DelayDependency)
            With vRowSchBar
                dta.InsertQuery_Schedule(.primary_mo, .identifier_jo, .identifier_wo, .primary_work_unit, .plan_start_time, .plan_end_time, _
                .plan_end_time, 0, .total_work_hour, .total_work_hour, .rate, 100, .rate, #1/1/1950#, #1/1/1950#, 0, 0, 0, 0, 0, 0, 0, _
                .qty_produce, 0, 0, .a, .r, .g, .b, ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled, .primary_process, 0, .qty_produce, .qty_produce, _
                100, 0, 0, .qty_produce, TipeDependency, DelayDependency, objWorkunit.GetPrimaryInputLocation(.primary_work_unit), objWorkunit.GetPrimaryOutputLocation(.primary_work_unit))
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetProductDependency(ByVal vPrimaryProduct As Double, ByRef vTipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency, ByRef vDelayDependency As Double) As Boolean
        Dim dtaDep As New DseTrans_ScheduleTableAdapters.tblmaster_product1TableAdapter
        Dim tblDep As New DseTrans_Schedule.tblmaster_product1DataTable
        Dim rowDep As DseTrans_Schedule.tblmaster_product1Row
        Try
            vTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
            vDelayDependency = 0
            dtaDep.Connection.ConnectionString = Core_Module.GetConstring
            If dtaDep.Fill(tblDep, vPrimaryProduct) > 0 Then
                rowDep = tblDep.Rows(0)
                vTipeDependency = rowDep.tipe_dependency
                vDelayDependency = rowDep.delay_dependency
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdatePlanStartEndScheduleWO(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String) As Integer
        Dim DtaItemSch As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch8TableAdapter
        Dim TblItemSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch8DataTable
        Dim RowItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch8Row
        Dim DtaItemWo As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim planStartTime, planEndTime As Date
        Try
            DtaItemSch.Connection.ConnectionString = Core_Module.GetConstring
            If DtaItemSch.Fill(TblItemSch, vIdentifierJO, vIdentifierWO) > 0 Then
                TblItemSch.DefaultView.RowFilter = ""
                TblItemSch.DefaultView.Sort = "[plan start time] ASC"
                RowItemSch = TblItemSch.DefaultView(0).Row
                planStartTime = RowItemSch.plan_start_time
                TblItemSch.DefaultView.Sort = "[plan end time] DESC"
                RowItemSch = TblItemSch.DefaultView(0).Row
                planEndTime = RowItemSch.plan_end_time
                Return DtaItemWo.UpdateQuery_PlanStartTimePlanEndTime(planStartTime, planEndTime, vIdentifierJO, vIdentifierWO)
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan true jika ada salah satu hal dibawah ini:
    ''' 1. plan start <= start date dan plan end >= end date
    ''' 2. plan start >= start date dan plan start <= end date
    ''' 3. plan end >= start date dan plan end <= end date
    ''' </summary>
    ''' <param name="vPrimaryWU"></param>
    ''' <param name="vStartDate"></param>
    ''' <param name="vEndDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahAda_PlanSchedule(ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim Jml As Nullable(Of Long)
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter

        Try
            Jml = DTA.ScalarQuery_GetJml1(vEndDate, vStartDate, vPrimaryWU)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                End If
            End If

            Jml = DTA.ScalarQuery_GetJml2(vPrimaryWU, vEndDate, vStartDate)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                End If
            End If

            Jml = DTA.ScalarQuery_GetJml3(vStartDate, vPrimaryWU, vEndDate)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
