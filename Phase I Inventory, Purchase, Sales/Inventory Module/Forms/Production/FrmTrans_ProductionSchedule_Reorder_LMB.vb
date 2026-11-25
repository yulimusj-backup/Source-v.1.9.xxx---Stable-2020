Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_Reorder_LMB
    Dim UserCode As String = "USER"
    Dim AnomaliesFound As Boolean = False
    Dim StartTime As Date = Now
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub btnStep1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep1.Click
        Try
            If MsgBox("Start allocating workshift?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim StartDate, EndDate As Date
        Dim PrimaryWS As Double = -1
        Dim DTAWU As New Core_Module.DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim TBLWU As New Core_Module.DseMasterWorkUnit.tblmaster_workunitDataTable
        Dim rowWU As Core_Module.DseMasterWorkUnit.tblmaster_workunitRow
        Dim DTAAllocate As New DseTransMO_WorkShift_Allocate2TableAdapters.tbltrans_mo_workshift_allocateTableAdapter

        Try
            ObjCD.UpdateThreading()

            'set nilai default
            StartDate = #1/1/2015#
            EndDate = #7/31/2017#
            PrimaryWS = 3

            DTAWU.Fill(TBLWU)
            Me.PrgBarStep1.Value = 0
            Me.PrgBarStep1.Maximum = TBLWU.Rows.Count
            For Each rowWU In TBLWU.Rows
                Me.PrgBarStep1.Value += 1
                Me.lblProgressStep1.Text = "Data " + Me.PrgBarStep1.Value.ToString + " of " + Me.PrgBarStep1.Maximum.ToString + " - " + rowWU.work_unit_code

                'hapuskan workshift
                DTAAllocate.DeleteQuery(rowWU.primary, PrimaryWS, EndDate, StartDate)
                'simpan ke db 
                Me.Simpan_UsingScript(PrimaryWS, rowWU.primary, StartDate, EndDate)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWU.Dispose()
            TBLWU.Dispose()
            DTAAllocate.Dispose()
        End Try
    End Sub

    Private Function Simpan_UsingScript(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim StartDateDummy As Date = #1/1/1950#
        Dim EndDateDummy As Date = #1/1/1950#

        Try
            'open connection
            Conn.Open()

            StartDateDummy = vStartDate.Date
            While StartDateDummy <= vEndDate
                If (StartDateDummy + TimeSpan.FromDays(30)).Date > vEndDate.Date Then
                    EndDateDummy = vEndDate.Date
                Else
                    EndDateDummy = (StartDateDummy + TimeSpan.FromDays(30)).Date
                End If
                sql = Me.GetInsertString(vPrimaryWS, vPrimaryWU, StartDateDummy, EndDateDummy)
                script.Query = sql
                'execute script
                script.Execute()

                StartDateDummy = (EndDateDummy + TimeSpan.FromDays(1)).Date
            End While
            'StartRowIndex += BatchQty

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString(ByVal vPrimaryWS As Double, ByVal vPrimaryWU As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As String
        Dim retval As String = ""
        'Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim Idx As Integer = 0
        Dim CurDate As Date = Today

        Try
            retval = "INSERT INTO `tbltrans_mo_workshift_allocate` (`date`,`primary work unit`,`primary workshift`,`created by`,`created date`) VALUES "
            While vStartDate.Date + TimeSpan.FromDays(Idx) <= vEndDate
                CurDate = vStartDate.Date + TimeSpan.FromDays(Idx)
                If CurDate.Date = vEndDate.Date Then
                    Delimiter = ";"
                End If
                retval = retval + vbCrLf + "('" + Format(CurDate, "yyyy-MM-dd") + "'," + vPrimaryWU.ToString + "," _
                   + vPrimaryWS.ToString + "," _
                   + "'" + Me.UserCode + "'," _
                   + "'" + Format(Now, "yyyy-MM-dd HH:mm:ss") + "'" _
                   + ")" + Delimiter

                Idx += 1
            End While

            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.PrgBarStep1.Value = 0
            Me.lblProgressStep1.Text = "Ready..."
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep2.Click
        Try
            If MsgBox("Start checking for anomalies?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim DTA As New DseTrans_Schedule2TableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL As New DseTrans_Schedule2.tbltrans_mo_item_wo_item_schDataTable
        Dim StartRef, EndRef As Date
        Dim ListDatRepair1 As New List(Of Double)

        Try
            ObjCD.UpdateThreading()

            AnomaliesFound = False
            StartRef = #1/1/2001#
            EndRef = #1/1/2050#

            DTA.FillBy_TglRef(TBL, StartRef, EndRef)

            If TBL.Rows.Count <> 0 Then
                Me.AnomaliesFound = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListDatRepair1 = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.AnomaliesFound = True Then
                MsgBox("Anomalies found, you must repair data manually.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Data checked.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UrutkanTanpaJeda()
        Dim DTA As New DseTrans_Schedule2TableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim TBL As New DseTrans_Schedule2.tbltrans_mo_item_wo_item_schDataTable
        Dim ObjSch As New ClsTrans_Schedule
        Dim TBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim DTAWU As New Core_Module.DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim TBLWU As New Core_Module.DseMasterWorkUnit.tblmaster_workunitDataTable
        Dim rowWU As Core_Module.DseMasterWorkUnit.tblmaster_workunitRow
        Dim ListWUTidakUpdate As New List(Of Double)
        Dim Dummy As String = ""
        Dim ObjWU As New Core_Module.ClsMasterWorkUnit

        Try
            Me.StartTime = Now

            DTAWU.Fill(TBLWU)
            Me.PrgBarStep3.Value = 0
            Me.PrgBarStep3.Maximum = TBLWU.Rows.Count
            For Each rowWU In TBLWU.Rows
                Me.PrgBarStep3.Value += 1
                Me.LblProgressStep3.Text = "Data " + Me.PrgBarStep3.Value.ToString + " of " + Me.PrgBarStep3.Maximum.ToString
                Me.UrutkanTanpaJeda(DTA, TBL, ObjSch, TBLSchBar, rowWU.primary, ListWUTidakUpdate)
            Next

            If ListWUTidakUpdate.Count > 0 Then
                For Each primarywu As Double In ListWUTidakUpdate
                    If Dummy = "" Then
                        Dummy = ObjWU.GetCode(primarywu)
                    Else
                        Dummy = Dummy + "; " + ObjWU.GetCode(primarywu)
                    End If
                Next

                MsgBox("There are unupdated work unit:" + Dummy, MsgBoxStyle.Information, "Information")
            End If

            Me.UpdateItemWO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjSch = Nothing
            TBLSchBar.Dispose()
            DTAWU.Dispose()
            TBLWU.Dispose()
            ListWUTidakUpdate = Nothing
            ObjWU = Nothing
        End Try
    End Sub

    Private Sub UpdateItemWO()
        Dim DTAItemWO As New DseTrans_Schedule2TableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TBLItemWO As New DseTrans_Schedule2.tbltrans_mo_item_woDataTable 
        Dim rowItemWO As DseTrans_Schedule2.tbltrans_mo_item_woRow
        Dim DTAItemSch As New DseTrans_Schedule2TableAdapters.tbltrans_mo_item_wo_item_sch1TableAdapter
        Dim TBLItemSch As New DseTrans_Schedule2.tbltrans_mo_item_wo_item_sch1DataTable 
        Dim rowItemSch As DseTrans_Schedule2.tbltrans_mo_item_wo_item_sch1Row
        Dim DTA As New DseTrans_Schedule2TableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim TBL As New DseTrans_Schedule2.tbltrans_mo_item_wo1DataTable
        Dim row As DseTrans_Schedule2.tbltrans_mo_item_wo1Row
        Dim StartDate, EndDate As Date

        Try
            DTA.Fill(TBL)
            DTAItemWO.Fill(TBLItemWO)
            DTAItemSch.Fill(TBLItemSch)

            Me.PrgBarStep3.Value = 0
            Me.PrgBarStep3.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.PrgBarStep3.Value += 1
                Me.LblProgressStep3.Text = "Updating item wo: " + Me.PrgBarStep3.Value.ToString + " of " + Me.PrgBarStep3.Maximum.ToString + " - " + row.MO_Number

                TBLItemWO.DefaultView.RowFilter = "[Primary MO] = " + row.Primary_MO.ToString
                TBLItemSch.DefaultView.RowFilter = "[primary mo] = " + row.Primary_MO.ToString

                For Each DV As Data.DataRowView In TBLItemWO.DefaultView
                    rowItemWO = DV.Row
                    StartDate = Date.MaxValue
                    EndDate = Date.MinValue

                    For Each DV1 As Data.DataRowView In TBLItemSch.DefaultView
                        rowItemSch = DV1.Row
                        If rowItemSch.identifier_jo = rowItemWO.Identifier_JO And rowItemSch.identifier_wo = rowItemWO.Identifier Then
                            If StartDate > rowItemSch.plan_start_time Then
                                StartDate = rowItemSch.plan_start_time
                            End If
                            If EndDate < rowItemSch.plan_end_time Then
                                EndDate = rowItemSch.plan_end_time
                            End If
                        End If
                    Next

                    If StartDate <> Date.MaxValue Then
                        DTAItemWO.UpdateQuery(StartDate, EndDate, rowItemWO.Primary)
                    End If
                Next

                While TBLItemWO.DefaultView.Count > 0
                    TBLItemWO.Rows.Remove(TBLItemWO.DefaultView(0).Row)
                End While
                While TBLItemSch.DefaultView.Count > 0
                    TBLItemSch.Rows.Remove(TBLItemSch.DefaultView(0).Row)
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemWO.Dispose()
            TBLItemWO.Dispose() 
            DTAItemSch.Dispose()
            TBLItemSch.Dispose() 
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub UrutkanTanpaJeda(ByVal DTA As DseTrans_Schedule2TableAdapters.tbltrans_mo_item_wo_item_schTableAdapter, _
    ByVal TBL As DseTrans_Schedule2.tbltrans_mo_item_wo_item_schDataTable, ByVal ObjSch As ClsTrans_Schedule, _
    ByVal TBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vPrimaryWU As Double, ByVal ListWUTidakUpdate As List(Of Double))
        Dim row, row1 As DseTrans_Schedule2.tbltrans_mo_item_wo_item_schRow
        Dim StartDate, EndDate, EndDateRef As Date
        Dim RefTime As Date = #1/1/2200#
        Dim RefStr As String = Me.LblProgressStep3.Text
        Dim counter As Long = 0

        Try
            StartDate = #1/1/1950#
            EndDate = #1/1/2200#
            EndDateRef = CDate(Format(Me.dtpStep3.Value.Date, "MM/dd/yyyy") + " 08:00")

            'isikan tbl wo item sch, sesuai primary wu
            TBL.Rows.Clear()
            DTA.Fill(TBL, vPrimaryWU)
            Me.LblProgressStep3.Text = RefStr + vbCrLf + "Loading item wo"
            If TBL.Rows.Count = 0 Then
                Exit Sub
            End If

            'isikan sch bar kosong, sesuai primary wu
            Me.LblProgressStep3.Text = RefStr + vbCrLf + "Loading sch bar"
            ObjSch.LoadScheduleBar_TanpaScheduledItem2(StartDate, EndDate, TBLSchBar, vPrimaryWU)

            'mulai dari yang terakhir, langsung urutkan tanpa jeda
            counter = 0
            For ij As Integer = TBL.Rows.Count - 1 To 0 Step -1
                counter += 1
                Me.LblProgressStep3.Text = RefStr + vbCrLf + "Recalculating - data " + counter.ToString + " of " + (TBL.Rows.Count - 1).ToString

                row = TBL.Rows(ij)

                If ij = TBL.Rows.Count - 1 Then
                    'If row.plan_end_time > EndDateRef Then
                    row.plan_end_time = EndDateRef
                    'End If
                    RefTime = #1/1/2200#
                Else
                    row1 = TBL.Rows(ij + 1)
                    RefTime = row1.plan_start_time
                End If

                If RefTime <> #1/1/2200# Then
                    If row.plan_end_time <> row1.plan_start_time Then
                        row.plan_end_time = ObjSch.GetStartTime(vPrimaryWU, row1.plan_start_time, 1, TBLSchBar)
                    End If
                End If
                row.plan_start_time = ObjSch.GetStartTime(vPrimaryWU, row.plan_end_time, row.plan_total_duration, TBLSchBar)

                If row.plan_start_time = #1/1/1950# Then
                    ListWUTidakUpdate.Add(vPrimaryWU)
                    Exit Sub
                End If
            Next

            'update ke db
            counter = 0
            For Each row In TBL.Rows
                counter += 1
                Me.LblProgressStep3.Text = RefStr + vbCrLf + "Saving to db - data " + counter.ToString + " of " + (TBL.Rows.Count - 1).ToString
                DTA.UpdateQuery1(row.plan_start_time, row.plan_end_time, row.primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep3.Click
        Try
            If MsgBox("Start re-order?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Timer1.Enabled = True
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            ObjCD.UpdateThreading()

            Me.UrutkanTanpaJeda()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.Timer1.Enabled = False
            Me.LblElapseTime.Text = ""
            Me.PrgBarStep3.Value = 0
            Me.LblProgressStep3.Text = "Ready..."
            MsgBox("Data updated in:" + ObjCD.DigitGrouping((Now - Me.StartTime).TotalMinutes, 1) + "minutes" + " :)", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_Reorder_LMB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.LblElapseTime.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.LblElapseTime.Text = "Elapsed time: " + ObjCD.DigitGrouping((Now - Me.StartTime).TotalMinutes, 1) + " minutes"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class