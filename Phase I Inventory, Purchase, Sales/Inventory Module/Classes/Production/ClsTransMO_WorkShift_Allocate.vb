Option Explicit On
Option Strict Off
Public Class ClsTransMO_WorkShift_Allocate
    Dim ObjMySQL As New Core_Module.ClsMySQLReference
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Public Enum EnumTipeBundling As Integer
        Inventory
        Purchase1
        Produksi1
    End Enum

    Public Function ApakahAmanDrJadwalTabrakan(ByVal vPrimaryWS As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal vListPrimaryWU As List(Of Double), ByRef vDummyStr As String) As Boolean
        Dim TipeWorksDay As Core_Module.ClsMasterWorkShift.TipeWorkdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday
        Dim TBLWS1 As New DseTransMO_WorkShift_Allocate2.tblmaster_workshift_itemDataTable
        Dim DTAWS1 As New DseTransMO_WorkShift_Allocate2TableAdapters.tblmaster_workshift_itemTableAdapter
        Dim rowWS1 As DseTransMO_WorkShift_Allocate2.tblmaster_workshift_itemRow
        Dim TBLWS2 As New DseTransMO_WorkShift_Allocate2.tblmaster_workshift_itemDataTable
        Dim rowWS2 As DseTransMO_WorkShift_Allocate2.tblmaster_workshift_itemRow
        Dim TBLConflict As New DseTransMO_WorkShift_Allocate2.DataTable1DataTable
        Dim rowConflict As DseTransMO_WorkShift_Allocate2.DataTable1Row
        Dim DOW As Integer = 0
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLAllocate As New DseTransMO_WorkShift_Allocate2.tbltrans_mo_workshift_allocateDataTable
        Dim rowAllocate As DseTransMO_WorkShift_Allocate2.tbltrans_mo_workshift_allocateRow
        Dim TBLAllocateD As New DseTransMO_WorkShift_Allocate2.tbltrans_mo_workshift_allocateDataTable
        Dim rowAllocateD As DseTransMO_WorkShift_Allocate2.tbltrans_mo_workshift_allocateRow
        Dim PrimaryWU As Double
        Dim DV As Data.DataRowView
        Dim ObjWS As New Core_Module.ClsMasterWorkShift
        Dim Idx As Integer = 0
        Dim ListPrimaryWS As New List(Of Double)

        Try
            vDummyStr = ""
            TBLConflict.Rows.Clear()
            DTAWS1.Connection.ConnectionString = Core_Module.GetConstring
            DTAAllocate.Connection.ConnectionString = DTAWS1.Connection.ConnectionString

            'cari data work shift yang dituju
            DTAWS1.Fill(TBLWS1, vPrimaryWS)
            'cari data workshift lain yang bertabrakan dengan data workshift tersebut
            DTAWS1.FillBy1(TBLWS2, vPrimaryWS)
            'tandain hari apa dan work shift apa
            For DOW = 0 To 6
                rowWS1 = TBLWS1.Rows(DOW)
                If rowWS1.tipe_workdays = CInt(Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday) Then
                    TBLWS2.DefaultView.RowFilter = "[day of week] = " + DOW.ToString
                    For Each DV In TBLWS2.DefaultView
                        rowWS2 = DV.Row
                        If rowWS2.tipe_workdays = CInt(Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday) Then
                            If rowWS1.schedule_out < rowWS1.schedule_in Then
                                rowWS1.schedule_out += 1440
                            End If
                            If rowWS2.schedule_out < rowWS2.schedule_in Then
                                rowWS2.schedule_out += 1440
                            End If

                            If (rowWS1.schedule_in <= rowWS2.schedule_in + 1 And rowWS1.schedule_out >= rowWS2.schedule_in + 1) Or _
                            (rowWS1.schedule_in <= rowWS2.schedule_out - 1 And rowWS1.schedule_out >= rowWS2.schedule_out - 1) Then
                                rowConflict = TBLConflict.NewDataTable1Row
                                rowConflict.Day_Of_Week = rowWS1.day_of_week
                                rowConflict.Primary_WS = rowWS2.primary_transaksi
                                rowConflict.Conflict = 0

                                TBLConflict.Rows.Add(rowConflict)
                            End If
                        End If
                    Next
                End If
            Next

            'load semua item allocate tgl dari sampai pada tabel dummy
            TBLAllocate.Rows.Clear()
            TBLAllocateD.Rows.Clear()
            DTAAllocate.Fill(TBLAllocateD, vEndDate, vStartDate, vPrimaryWS)
            'pindahkan ke tabel hanya yang primary wu-nya di dalam list
            For Each PrimaryWU In vListPrimaryWU
                TBLAllocateD.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                For Each DV In TBLAllocateD.DefaultView
                    rowAllocateD = DV.Row
                    rowAllocate = TBLAllocate.Newtbltrans_mo_workshift_allocateRow
                    For ij As Integer = 0 To TBLAllocate.Columns.Count - 1
                        rowAllocate.Item(ij) = rowAllocateD.Item(ij)
                    Next
                    rowAllocate.day_of_week = rowAllocate._date.DayOfWeek
                    TBLAllocate.Rows.Add(rowAllocate)
                Next
            Next

            'untuk semua item tabrakan, loop tabel item allocate, cari yang cocok
            For Each rowConflict In TBLConflict.Rows
                TBLAllocate.DefaultView.RowFilter = "[day of week] = " + CInt(rowConflict.Day_Of_Week).ToString + " and [primary workshift] = " + rowConflict.Primary_WS.ToString
                If TBLAllocate.DefaultView.Count > 0 Then
                    rowConflict.Conflict = 1
                    If ListPrimaryWS.Contains(rowConflict.Primary_WS) = False Then
                        Idx += 1
                        ListPrimaryWS.Add(rowConflict.Primary_WS)
                        If vDummyStr = "" Then
                            vDummyStr = Idx.ToString + ". " + ObjWS.GetCodeByPrimary(rowConflict.Primary_WS)
                        Else
                            vDummyStr = vDummyStr + vbCrLf + Idx.ToString + ". " + ObjWS.GetCodeByPrimary(rowConflict.Primary_WS)
                        End If
                    End If
                End If
            Next

            If vDummyStr <> "" Then
                Return False
            End If

            'jika sampe sini, maka ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            TBLWS1.Dispose()
            DTAWS1.Dispose()
            TBLWS2.Dispose()
            TBLConflict.Dispose()
            DTAAllocate.Dispose()
            TBLAllocate.Dispose()
            TBLAllocateD.Dispose()
            ObjWS = Nothing
            ListPrimaryWS = Nothing
        End Try
    End Function

    Public Function GetJmlSch(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryWU As Double, _
    ByVal DTA As DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_item_wo_item_schTableAdapter) As Double
        Dim RetVal1 As Nullable(Of Long)
        Dim RetVal2 As Nullable(Of Long)
        Dim RetVal As Integer = 0

        Try
            RetVal1 = DTA.ScalarQuery_GetJml1(vEndDate.Date + TimeSpan.FromMinutes(1439), vStartDate.Date, vPrimaryWU)
            RetVal2 = DTA.ScalarQuery_GetJml2(vEndDate.Date + TimeSpan.FromMinutes(1439), vStartDate.Date, vPrimaryWU)
            If RetVal1.HasValue = True Then
                RetVal += RetVal1.Value
            End If
            If RetVal2.HasValue = True Then
                RetVal += RetVal2.Value
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlOT(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryWU As Double, ByVal DTA As DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_overtime_itemTableAdapter) As Double
        Dim RetVal As Nullable(Of Long)

        Try
            RetVal = DTA.ScalarQuery_GetJml(vPrimaryWU, vEndDate.Date + TimeSpan.FromMinutes(1439), vStartDate)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetJmlBundling(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryWU As Double, ByVal DTA As DseTransMO_WorkShift_Allocate2TableAdapters.tbltransaksi_stockbundlingTableAdapter) As Double
        Dim RetVal As Nullable(Of Long)

        Try
            RetVal = DTA.ScalarQuery_GetJml(CInt(EnumTipeBundling.Produksi1), vEndDate.Date + TimeSpan.FromMinutes(1439), vStartDate, vPrimaryWU)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetFirmPlan() As Date
        Dim Dummy As String
        Dim CurrentTime As Date = Today

        Try
            CurrentTime = Me.ObjMySQL.GetServerTime
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1)
            If Dummy = "" Then
                Dummy = Format(Me.GetDefaultSchHor_FirmPlanned(CurrentTime), "dd MMM yyyy HH:mm")
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_FirmPlanned1, 7)
            Else
                Dummy = Format(CurrentTime + TimeSpan.FromDays(Dummy), "dd MMM yyyy HH:mm")
            End If

            Return CDate(Dummy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Function GetDefaultSchHor_FirmPlanned(ByVal vCurrentTime As Date) As Date
        Try
            Return (vCurrentTime + TimeSpan.FromDays(7)).Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTBL_WU(ByVal TBL_WU As DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable, ByVal vPrimaryWUG As Double) As Boolean
        Dim DTA As New DseTransMO_WorkShift_AllocateTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL_WU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTBL_WU_byprimarywugroup(ByVal TBL_WU As DseTransMO_WorkShift_Allocate.tblmaster_workunitDataTable, ByVal vPrimaryWUG As Double) As Boolean
        Dim DTA As New DseTransMO_WorkShift_AllocateTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_primarywug(TBL_WU, vPrimaryWUG)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTBL_WS(ByVal TBL_WS As DseTransMO_WorkShift_Allocate.tblmaster_workshiftDataTable) As Boolean
        Dim DTA As New DseTransMO_WorkShift_AllocateTableAdapters.tblmaster_workshiftTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL_WS, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.HR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahbolehHapusAllocate(ByVal i As FrmTransMO_WorkShift_Allocate) As Boolean
        Try
            If i.dgvItem.Rows.Count <= 0 Then
                MsgBox("Please filter data allocation.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If i.dtpFrom.Value.Date > i.dtpTo.Value.Date Then
                MsgBox("Period from should not be greater than period to.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public msgHapus As String
    Public tglmsgx As String
    Public tglmsgy As Date
    Public list_msg As New List(Of Double)
    Public count As Integer = 3

    Public vlist_codewu As New List(Of String)
    Public vmsg_unallocate As String = ""
    Public vjmlbatas_msgunallocate As Integer = 1

    Public Function DeleteAllocate(ByVal vTglDari As Date, ByVal vTglSampai As Date, ByVal vPrimaryWU As Double, ByVal vPrimaryWS As Double, ByVal munculkanpesan As Boolean, ByVal filter_all As Boolean) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Dim dummy_date As Date
        Dim retval As String = ""
        Dim adaspl As Boolean = False
        Dim no_ As Integer = 1
        Dim tanggalMsgbox1 As String = ""
        Dim vcode_wu As String = ""
        Dim jml_pjgmsg As Integer = 0
        Dim objwu As New Core_Module.ClsMasterWorkUnit
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            vcode_wu = objwu.GetCode(vPrimaryWU)
            If filter_all = False Then
                Me.isitbl_bundling(True, vTglDari, vTglSampai)
            End If
            For i As Integer = 0 To DateDiff(DateInterval.Day, vTglDari.Date, vTglSampai.Date)
                dummy_date = vTglDari.Date.AddDays(i)
                'jika unallocate per workunit
                If filter_all = False Then
                    'cek di spl dan bundling
                    If Me.apakahsudahada_displ(vPrimaryWU, vPrimaryWS, dummy_date.Date, False) = False And Me.apakahsudahada_dibundling(vPrimaryWU, _
                     dummy_date, False) = False Then
                        dta.DeleteQuery_allocatepertgl(dummy_date.Date, vPrimaryWU, vPrimaryWS)
                    Else
                        MsgBox("Can't unallocate this data referred with other transaction.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    'jika unallocate all 
                    If Me.apakahsudahada_displ(vPrimaryWU, vPrimaryWS, dummy_date.Date, True) = False And Me.apakahsudahada_dibundling(vPrimaryWU, _
                     dummy_date.Date, True) = False Then
                        dta.DeleteQuery_allocatepertgl(dummy_date.Date, vPrimaryWU, vPrimaryWS)
                    Else
                        'jika tidak benar

                        If vjmlbatas_msgunallocate < 2 Then
                            If vlist_codewu.Contains(vcode_wu) = True Then
                                If jml_pjgmsg < 5 Then
                                    vmsg_unallocate &= ", " & Format(dummy_date.Date, "dd MMM yyyy") & vbNewLine
                                    jml_pjgmsg += 1
                                Else
                                    If vmsg_unallocate.Substring(vmsg_unallocate.Length - 1, 1) <> "," Then
                                        vmsg_unallocate &= " etc.,"
                                    End If
                                End If
                            Else
                                If vjmlbatas_msgunallocate < 2 Then
                                    vmsg_unallocate &= vbNewLine & vcode_wu & " : " & Format(dummy_date.Date, "dd MMM yyyy")
                                    vlist_codewu.Add(vcode_wu)
                                    vjmlbatas_msgunallocate += 1
                                End If
                            End If
                        Else
                            If vmsg_unallocate.Length > 0 Then
                                If vmsg_unallocate.Substring(vmsg_unallocate.Length - 1, 1) <> "." Then
                                    vmsg_unallocate &= "many more.."
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Dim tbl_spl As New DseTransMO_WorkShift_Allocate.tbltrans_spl_item_logDataTable
    Dim row_spl As DseTransMO_WorkShift_Allocate.tbltrans_spl_item_logRow
    Public Function isitbl_spl() As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_spl_item_logTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl_spl)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function InsertAllocate(ByVal vTglDari As Date, ByVal vTglSampai As Date, ByVal vUsercode As String, ByVal vPrimaryWS As String, _
    ByVal vPrimaryWU As Double, ByVal tblInsertDB As DseTransMO_WorkShift_Allocate.tblinsertDBDataTable) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Dim rowInsertDB As DseTransMO_WorkShift_Allocate.tblinsertDBRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            tblInsertDB.DefaultView.RowFilter = Nothing
            tblInsertDB.DefaultView.RowFilter = "[date] >=#" & Format(vTglDari.Date, "dd MMM yyy") & "#" & "AND [date]<= #" & Format(vTglSampai.Date, "dd MMM yyyy") & "#" _
            & "AND [primary work unit]='" & vPrimaryWU.ToString & "'"
            If tblInsertDB.DefaultView.Count > 0 Then
                For count As Integer = 0 To tblInsertDB.DefaultView.Count - 1
                    rowInsertDB = tblInsertDB.DefaultView(count).Row
                    With rowInsertDB
                        If apakahschedule_allocatebenar(vPrimaryWU, ._date, vPrimaryWS, False) = True Then
                            dta.InsertQuery_Allocate(._date, .primary_work_unit, vPrimaryWS, vUsercode, Now)
                        Else
                            Exit Function
                            Return False
                        End If
                    End With
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function apakahsudahada_displ(ByVal vprimarywu As Double, ByVal vprimaryws As Double, ByVal vtgl As Date, Optional ByVal difill As Boolean = False) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim retval As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If difill = False Then
                retval = dta.ScalarQuery_getjmlinspl(vtgl, vprimarywu, vprimaryws)
                If retval.HasValue = True Then
                    If retval.Value > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            Else
                tbl_spl.DefaultView.RowFilter = Nothing
                tbl_spl.DefaultView.RowFilter = "[primary wokunit]= " & vprimarywu & "AND [primary workshift]= " & vprimaryws & "AND [date]= #" & vtgl & "#"
                If tbl_spl.DefaultView.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    'Public Function getjmlws_by(ByVal vprimary_wu As Double, ByVal vtgl As Date) As Integer
    '    Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
    '    Dim retval As Nullable(Of Long)
    '    Try
    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        retval = dta.ScalarQuery_getjmlws_by(vprimary_wu, vtgl)
    '        If retval.HasValue = True Then
    '            If retval.Value > 0 Then
    '                Return retval
    '            Else
    '                Return 0
    '            End If
    '        End If
    '        Return 0
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        dta.Dispose()
    '    End Try
    'End Function



    'Dim tbl_schedule As New DseTransMO_WorkShift_Allocate.tblmaster_workshiftDataTable
    'Public Function get_schedule(ByRef in_ As Integer, ByRef out_ As Integer, ByVal vhari As Integer, ByVal vprimary_workshift As Double) As Boolean
    '    Try
    '        Me.isitbl_schedule(vprimary_workshift)
    '        If tbl_schedule.Rows.Count > 0 Then
    '            tbl_schedule.DefaultView.RowFilter = Nothing
    '            tbl_schedule.DefaultView.RowFilter = "[day of week]=" & vhari
    '            For i As Integer = 0 To tbl_schedule.DefaultView.Count - 1
    '                'row_schedule = tbl_schedule.DefaultView(0).Row
    '                'in_ = row_schedule.schedule_in
    '                'out_ = row_schedule.schedule_out
    '            Next
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Function isitbl_schedule(ByVal vprimary As Double) As Boolean
    '    Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tblmaster_workshift_itemTableAdapter
    '    Try
    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        'dta.FillBy_primaryws(tbl_schedule, vprimary)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        dta.Dispose()
    '    End Try
    'End Function

    Public Function isitbl_allocate() As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl_allocate1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function apakahsudahada_ws1(ByVal vtgl As Date, ByVal vprimary_wu As Double) As Boolean
        Try
            tbl_allocate1.DefaultView.RowFilter = Nothing
            tbl_allocate1.DefaultView.RowFilter = "[primary work unit]= " & vprimary_wu.ToString & "AND date = " & "#" & vtgl & "#"
            If tbl_allocate1.DefaultView.Count > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function apakahsudahada_ws(ByVal vtgl As Date, ByVal vprimary_wu As Double) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Dim retval As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_getjmlby_tgldanwu(vprimary_wu, vtgl)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    '
    Public Function get_schedule(ByVal vprimary_ws As Double, ByRef vin As Integer, ByRef vout As Double, ByVal vhari As Double) As Boolean
        Try
            If tbl_schedule.Rows.Count > 0 Then
                tbl_schedule.DefaultView.RowFilter = Nothing
                tbl_schedule.DefaultView.RowFilter = "[primary transaksi]=" & vprimary_ws & "AND [day of week]= " & vhari
                For i As Integer = 0 To tbl_schedule.DefaultView.Count - 1
                    row_schedule = tbl_schedule.DefaultView(0).Row
                    With row_schedule
                        vin = .schedule_in
                        vout = .schedule_out
                    End With
                    Return True
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Dim tbl_schedule As New DseTransMO_WorkShift_Allocate.tblmaster_workshift_itemDataTable
    Dim row_schedule As DseTransMO_WorkShift_Allocate.tblmaster_workshift_itemRow
    Public Function isitbl_schedule() As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tblmaster_workshift_itemTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl_schedule)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Dim tbl_allocate1 As New DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocate1DataTable
    Dim row_allocate1 As DseTransMO_WorkShift_Allocate.tbltrans_mo_workshift_allocate1Row
    Public Function isitbl_allocateby(ByVal vprimary_wu As Double, ByVal vtgl As Date) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_primarywu_anddate(tbl_allocate1, vtgl, vprimary_wu)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function apakahschedule_allocatebenar(ByVal vprimary_wu As Double, ByVal vtgl As Date, ByVal vprimary_ws As Double, ByVal difill As Boolean) As Boolean
        Dim in_, out_, in_pembanding, out_pembanding As Double
        Try
            tbl_allocate1.Rows.Clear()
            'jika allocater per work unit
            If difill = False Then
                'cek di allocatenya sendiri
                If Me.apakahsudahada_ws(vtgl, vprimary_wu) = True Then
                    Me.isitbl_allocateby(vprimary_wu, vtgl)
                    Me.isitbl_schedule()
                    For i As Integer = 0 To tbl_allocate1.Rows.Count - 1
                        row_allocate1 = tbl_allocate1.Rows(i)
                        Me.get_schedule(row_allocate1.primary_workshift, in_pembanding, out_pembanding, vtgl.DayOfWeek)
                        'bandingkan hanya dengan shift yg beda
                        If row_allocate1.primary_workshift <> vprimary_ws Then
                            Me.get_schedule(vprimary_ws, in_, out_, vtgl.DayOfWeek)
                            If in_ < out_pembanding Then
                                MsgBox("Can't allocate this schedule in already exist.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    Next
                End If
            Else
                'jika allocate semua work unit
                Me.isitbl_allocate()
                If Me.apakahsudahada_ws1(vtgl, vprimary_wu) = True Then
                    Me.isitbl_allocateby(vprimary_wu, vtgl)
                    Me.isitbl_schedule()
                    For i As Integer = 0 To tbl_allocate1.Rows.Count - 1
                        row_allocate1 = tbl_allocate1.Rows(i)
                        Me.get_schedule(row_allocate1.primary_workshift, in_pembanding, out_pembanding, vtgl.DayOfWeek)
                        If row_allocate1.primary_workshift <> vprimary_ws Then
                            Me.get_schedule(vprimary_ws, in_, out_, vtgl.DayOfWeek)
                            If in_ < out_pembanding Then
                                Return False
                            End If
                        End If
                    Next
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Dim tbl_bundling As New DseTransMO_WorkShift_Allocate.tbltransaksi_stockbundlingDataTable
    Dim row_bundling As DseTransMO_WorkShift_Allocate.tbltransaksi_stockbundlingRow

    Public Function isitbl_bundling(ByVal vpertglstartend As Boolean, ByVal vstart As Date, ByVal vend As Date) As Boolean
        Dim dta As New DseTransMO_WorkShift_AllocateTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vpertglstartend = False Then
                dta.Fill(tbl_bundling)
            Else
                Dim start As String = "#" & vstart & " 00:00:00#"
                Dim end_ As String = "#" & vend & " 23:59:59#"
                dta.FillBy_startendactual(tbl_bundling, start, end_)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function apakahsudahada_dibundling(ByVal vprimarywu As Double, ByVal vdate As Date, Optional ByVal difill As Boolean = False) As Boolean
        Try
            ''''cek unallocate per work unit
            '''If difill = False Then
            '''    Dim start As String = "#" & vtgldari & " 00:00:00#"
            '''    Dim end_ As String = "#" & vtglsampai & " 23:59:59#"

            '''    retval = dta.ScalarQuery_getjmlin_bundling(vprimarywu, start, end_)
            '''    If retval.HasValue = True Then
            '''        If retval.Value > 0 Then
            '''            Return True
            '''        Else
            '''            Return False
            '''        End If
            '''    End If

            '''Else
            '''    'cek per all unallocate
            '''    Dim start As String = vtgldari & " 00:00:00"
            '''    Dim end_ As String = vtglsampai & " 23:59:59"

            '''    tbl_bundling.DefaultView.RowFilter = Nothing
            '''    tbl_bundling.DefaultView.RowFilter = "[primary work unit]= " & vprimarywu & "AND [end time actual] >= #" & start & "# AND [start time actual] <= #" & end_ & "#"
            '''    If tbl_bundling.DefaultView.Count > 0 Then
            '''        Return True
            '''    Else
            '''        Return False
            '''    End If
            '''End If

            tbl_bundling.DefaultView.RowFilter = Nothing
            tbl_bundling.DefaultView.RowFilter = "[primary work unit]= " & vprimarywu & "AND [end time actual] >= #" & vdate & "# AND [start time actual] <= #" & vdate & "#"
            If tbl_bundling.DefaultView.Count > 0 Then
                Return True
            Else
                Return False
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

End Class
