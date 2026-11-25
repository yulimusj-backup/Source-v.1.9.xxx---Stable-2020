Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_Reorder

    Public TblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim TblSchBarTanpaWO As New DseTrans_Schedule.tbl_sch_barDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjSch As New ClsTrans_Schedule
    Public PrimaryWU As Double = -1
    Public UserCode As String
    Public IdentifierWO As String = ""
    Public IdentifierJO As String = ""

    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ShowMO, ShowJO, ShowWO, ShowProcess, ShowCodeNo, ShowDesc, ShowRate, ShowQty, ShowUnit, ShowWH, ShowDayStart, ShowDayFinish, ShowProgress As String
    Dim ObjCD As New Core_Module.ClsCekData
    Dim RowSchBar1 As DseTrans_Schedule.tbl_sch_barRow
    Dim RowSchBar2 As DseTrans_Schedule.tbl_sch_barRow
    Dim delay As TimeSpan
    Dim rowSch As DseTrans_Schedule.tbl_sch_barRow

    Private Sub FrmTrans_ProductionSchedule_Reorder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.EnableControl(False)
            Me.Cursor = Cursors.WaitCursor
            Me.BackgroundWorker1.RunWorkerAsync()
            'Me.colIdentifierJO.Visible = True
            'Me.colIdentifierWO.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowItem()
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.colIdentifierJO.Index).Value = Me.IdentifierJO And _
                rowDGV.Cells(Me.colIdentifierWO.Index).Value = Me.IdentifierWO Then
                    rowDGV.Cells(Me.ColModify.Index).Value = Boolean.TrueString
                    Me.dgvItem.ClearSelection()
                    Me.dgvItem.Rows(rowDGV.Index).Selected = True

                    Me.dgvItem.FirstDisplayedScrollingRowIndex = rowDGV.Index

                    'While Me.dgvItem.Rows(rowDGV.Index).Displayed = False _
                    'Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    '    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    'End While
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColModify.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Dim listRow As New List(Of DataGridViewRow)
        Try
            Me.Pause()
            If IsValidModify(True, False) Then
                listRow = GetSelectedDgvRows(True)
                For Each row As DataGridViewRow In listRow
                    If row.Index <> 0 Then
                        Dim indexbaru As Integer = row.Index - 1

                        ''
                        Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & Me.PrimaryWU.ToString & " and [identifier jo] = '" & Me.dgvItem.Item(Me.colIdentifierJO.Index, indexbaru).Value & "' and [identifier wo] = '" & Me.dgvItem.Item(Me.colIdentifierWO.Index, indexbaru).Value & "'"
                        If Me.TblSchBar.DefaultView.Count > 0 Then
                            Me.RowSchBar1 = Me.TblSchBar.DefaultView(0).Row
                            Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & Me.PrimaryWU.ToString & " and [identifier jo] = '" & Me.dgvItem.Item(Me.colIdentifierJO.Index, indexbaru + 1).Value & "' and [identifier wo] = '" & Me.dgvItem.Item(Me.colIdentifierWO.Index, indexbaru + 1).Value & "'"
                            If Me.TblSchBar.DefaultView.Count > 0 Then
                                Me.RowSchBar2 = Me.TblSchBar.DefaultView(0).Row
                                Me.delay = Me.RowSchBar2.plan_start_time - Me.RowSchBar1.plan_end_time
                                Me.RowSchBar2.plan_start_time = Me.RowSchBar1.plan_start_time
                                Me.RowSchBar2.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar2.plan_start_time, Me.RowSchBar2.total_work_hour, Me.TblSchBarTanpaWO)
                                Me.RowSchBar1.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar2.plan_end_time, Me.delay.Minutes, Me.TblSchBarTanpaWO)
                                Me.RowSchBar1.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar1.plan_start_time, Me.RowSchBar1.total_work_hour, Me.TblSchBarTanpaWO)
                            End If
                        End If

                        Me.dgvItem.Rows.Remove(row)
                        Me.dgvItem.Rows.InsertRange(indexbaru, row)

                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartDate.Index).Value = Format(Me.RowSchBar2.plan_start_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar2.plan_start_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartHour.Index).Value = Format(Me.RowSchBar2.plan_start_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndDate.Index).Value = Format(Me.RowSchBar2.plan_end_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar2.plan_end_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndHour.Index).Value = Format(Me.RowSchBar2.plan_end_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColStartDate.Index).Value = Format(Me.RowSchBar1.plan_start_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColStartDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar1.plan_start_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColStartHour.Index).Value = Format(Me.RowSchBar1.plan_start_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColEndDate.Index).Value = Format(Me.RowSchBar1.plan_end_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColEndDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar1.plan_end_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru + 1).Cells(Me.ColEndHour.Index).Value = Format(Me.RowSchBar1.plan_end_time, "HH.mm")

                        Me.dgvItem.Focus()
                        Me.dgvItem.ClearSelection()
                        row.Selected = True
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Resume_()
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Dim listRow As New List(Of DataGridViewRow)
        Try
            Me.Pause()
            If IsValidModify(False, False) Then
                listRow = GetSelectedDgvRows(False)
                For Each row As DataGridViewRow In listRow
                    If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                        Dim indexbaru As Integer = row.Index + 1

                        ''
                        Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & Me.PrimaryWU.ToString & " and [identifier jo] = '" & Me.dgvItem.Item(Me.colIdentifierJO.Index, indexbaru - 1).Value & "' and [identifier wo] = '" & Me.dgvItem.Item(Me.colIdentifierWO.Index, indexbaru - 1).Value & "'"
                        If Me.TblSchBar.DefaultView.Count > 0 Then
                            Me.RowSchBar1 = Me.TblSchBar.DefaultView(0).Row
                            Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & Me.PrimaryWU.ToString & " and [identifier jo] = '" & Me.dgvItem.Item(Me.colIdentifierJO.Index, indexbaru).Value & "' and [identifier wo] = '" & Me.dgvItem.Item(Me.colIdentifierWO.Index, indexbaru).Value & "'"
                            If Me.TblSchBar.DefaultView.Count > 0 Then
                                Me.RowSchBar2 = Me.TblSchBar.DefaultView(0).Row
                                Me.delay = Me.RowSchBar2.plan_start_time - Me.RowSchBar1.plan_end_time
                                Me.RowSchBar2.plan_start_time = Me.RowSchBar1.plan_start_time
                                Me.RowSchBar2.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar2.plan_start_time, Me.RowSchBar2.total_work_hour, Me.TblSchBarTanpaWO)
                                Me.RowSchBar1.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar2.plan_end_time, Me.delay.Minutes, Me.TblSchBarTanpaWO)
                                Me.RowSchBar1.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.RowSchBar1.plan_start_time, Me.RowSchBar1.total_work_hour, Me.TblSchBarTanpaWO)
                            End If
                        End If

                        Me.dgvItem.Rows.Remove(row)
                        Me.dgvItem.Rows.InsertRange(indexbaru, row)

                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColStartDate.Index).Value = Format(Me.RowSchBar2.plan_start_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColStartDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar2.plan_start_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColStartHour.Index).Value = Format(Me.RowSchBar2.plan_start_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColEndDate.Index).Value = Format(Me.RowSchBar2.plan_end_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColEndDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar2.plan_end_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru - 1).Cells(Me.ColEndHour.Index).Value = Format(Me.RowSchBar2.plan_end_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartDate.Index).Value = Format(Me.RowSchBar1.plan_start_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar1.plan_start_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColStartHour.Index).Value = Format(Me.RowSchBar1.plan_start_time, "HH.mm")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndDate.Index).Value = Format(Me.RowSchBar1.plan_end_time, "dd MMM yyyy")
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndDay.Index).Value = Me.ObjCD.GetHari_English(Me.RowSchBar1.plan_end_time.Date.DayOfWeek)
                        Me.dgvItem.Rows(indexbaru).Cells(Me.ColEndHour.Index).Value = Format(Me.RowSchBar1.plan_end_time, "HH.mm")

                        Me.dgvItem.Focus()
                        Me.dgvItem.ClearSelection()
                        row.Selected = True
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Resume_()
        End Try
    End Sub

    Private Sub btnUpJO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpJO.Click
        Dim SelectedRows As New List(Of DataGridViewRow)
        Dim PrevFinish As Date = #1/1/1950#
        Dim PrevFinishRef As Date = #1/1/1950#
        Dim FirstReferenceRowIndex_Up As Integer
        Dim FirstInitial As Boolean
        Dim FirstSelectedJO As String = ""
        Dim FirstSelectedWO As String = ""
        Dim SelectedCout As Integer = 0
        Try
            Me.Pause()
            If IsValidModify(True, True) Then
                SelectedRows = GetSelectedDgvRows(True)
                FirstSelectedJO = SelectedRows(0).Cells(Me.colIdentifierJO.Index).Value.ToString
                FirstSelectedWO = SelectedRows(0).Cells(Me.colIdentifierWO.Index).Value.ToString
                SelectedCout = SelectedRows.Count

                ''reschedule selected items JO below => move to up
                ''catat delay:
                FirstInitial = True
                For Each Row As DataGridViewRow In SelectedRows
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                        Else
                            rowSch.delay = CDec(Me.ObjSch.GetWorkDuration(PrevFinish, rowSch.plan_start_time, Me.PrimaryWU, Me.TblSchBarTanpaWO))
                        End If
                        PrevFinish = rowSch.plan_end_time
                    End If
                Next
                ''reschedule, set start ke item paling atas (jo reference yang tidak terhalang jo lain)
                FirstReferenceRowIndex_Up = GetFirstReferenceRowIndex_Up()
                FirstInitial = True
                For Each Row As DataGridViewRow In SelectedRows
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                            rowSch.plan_start_time = CDate(Me.dgvItem(Me.ColStartDate.Index, FirstReferenceRowIndex_Up).Value.ToString).Date.AddMinutes(Me.ObjCD.Convert_ToMinutes(Me.dgvItem(Me.ColStartHour.Index, FirstReferenceRowIndex_Up).Value.ToString))
                        Else
                            rowSch.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, PrevFinish, rowSch.delay, Me.TblSchBarTanpaWO)
                        End If
                        rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, rowSch.plan_start_time, rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                        PrevFinish = rowSch.plan_end_time
                        PrevFinishRef = PrevFinish
                    End If
                Next

                ''reschedule all items JO above => move to below
                ''catat delay:
                FirstInitial = True 
                For i As Integer = FirstReferenceRowIndex_Up To Me.dgvItem.Rows.Count - 1
                    Dim Row As DataGridViewRow = Me.dgvItem.Rows(i)
                    If Row.Cells(Me.ColModify.Index).Value.ToString = Boolean.TrueString Then
                        Exit For
                    End If
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                        Else
                            rowSch.delay = CDec(Me.ObjSch.GetWorkDuration(PrevFinish, rowSch.plan_start_time, Me.PrimaryWU, Me.TblSchBarTanpaWO))
                        End If
                        PrevFinish = rowSch.plan_end_time
                    End If
                Next
                ''reschedule jo reference:
                FirstInitial = True
                For i As Integer = FirstReferenceRowIndex_Up To Me.dgvItem.Rows.Count - 1
                    Dim Row As DataGridViewRow = Me.dgvItem.Rows(i)
                    If Row.Cells(Me.ColModify.Index).Value.ToString = Boolean.TrueString Then
                        Exit For
                    End If
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                            rowSch.plan_start_time = PrevFinishRef
                        Else
                            rowSch.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, PrevFinish, rowSch.delay, Me.TblSchBarTanpaWO)
                        End If
                        rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, rowSch.plan_start_time, rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                        PrevFinish = rowSch.plan_end_time
                    End If
                Next

                Me.LoadDgv()
                For Each Row As DataGridViewRow In Me.dgvItem.Rows
                    If Row.Cells(Me.colIdentifierJO.Index).Value.ToString = FirstSelectedJO And _
                    Row.Cells(Me.colIdentifierWO.Index).Value.ToString = FirstSelectedWO Then
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(Row.Index).Selected = True
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Row.Index
                        For i As Integer = Row.Index To Row.Index + SelectedCout - 1
                            Dim Row_ As DataGridViewRow = Me.dgvItem.Rows(i)
                            Row_.Cells(Me.ColModify.Index).Value = Boolean.TrueString
                        Next
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Resume_()
        End Try
    End Sub

    Private Function GetFirstReferenceRowIndex_Up() As Integer
        Dim FirstReferenceRowIndex As Integer = -1
        Dim FirstSelectedRowIndex_Up As Integer = -1
        Dim IdentifierJO_Ref As String = ""
        Try
            FirstSelectedRowIndex_Up = GetFirstSelectedRowIndex_Up()
            IdentifierJO_Ref = Me.dgvItem(Me.colIdentifierJO.Index, FirstSelectedRowIndex_Up - 1).Value.ToString
            For i As Integer = FirstSelectedRowIndex_Up - 1 To 0 Step -1
                Dim Row As DataGridViewRow
                Row = Me.dgvItem.Rows(i)
                If Row.Cells(Me.colIdentifierJO.Index).Value.ToString = IdentifierJO_Ref Then
                    FirstReferenceRowIndex = i
                Else
                    Exit For
                End If
            Next
            Return FirstReferenceRowIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetFirstSelectedRowIndex_Up() As Integer
        Try
            For Each Row As DataGridViewRow In Me.dgvItem.Rows
                If Row.Cells(Me.ColModify.Index).Value.ToString = Boolean.TrueString Then
                    Return Row.Index
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function 

    Private Sub btnDownJO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownJO.Click
        Dim SelectedRows As New List(Of DataGridViewRow)
        Dim PrevFinish As Date = #1/1/1950#
        Dim PrevFinishRef As Date = #1/1/1950#
        Dim FirstReferenceRowIndex_Down As Integer
        Dim FirstInitial As Boolean
        Dim FirstSelectedJO As String = ""
        Dim FirstSelectedWO As String = ""
        Dim ReferenceCount As Integer = 0
        Dim SelectedCout As Integer = 0
        Try
            Me.Pause()
            If IsValidModify(False, True) Then
                SelectedRows = GetSelectedDgvRows(True)
                FirstSelectedJO = SelectedRows(0).Cells(Me.colIdentifierJO.Index).Value.ToString
                FirstSelectedWO = SelectedRows(0).Cells(Me.colIdentifierWO.Index).Value.ToString
                SelectedCout = SelectedRows.Count

                ''reschedule selected items JO below => move to above
                ''catat delay:
                FirstReferenceRowIndex_Down = GetFirstReferenceRowIndex_Down(ReferenceCount)
                FirstInitial = True
                For i As Integer = FirstReferenceRowIndex_Down To FirstReferenceRowIndex_Down + ReferenceCount - 1
                    Dim Row As DataGridViewRow = Me.dgvItem.Rows(i) 
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                        Else
                            rowSch.delay = CDec(Me.ObjSch.GetWorkDuration(PrevFinish, rowSch.plan_start_time, Me.PrimaryWU, Me.TblSchBarTanpaWO))
                        End If
                        PrevFinish = rowSch.plan_end_time
                    End If
                Next
                ''reschedule, set start ke item paling atas (jo reference yang tidak terhalang jo lain)
                FirstInitial = True
                For i As Integer = FirstReferenceRowIndex_Down To FirstReferenceRowIndex_Down + ReferenceCount - 1
                    Dim Row As DataGridViewRow = Me.dgvItem.Rows(i) 
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                            rowSch.plan_start_time = CDate(SelectedRows(0).Cells(Me.ColStartDate.Index).Value.ToString).Date.AddMinutes(Me.ObjCD.Convert_ToMinutes(SelectedRows(0).Cells(Me.ColStartHour.Index).Value.ToString))
                        Else
                            rowSch.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, PrevFinish, rowSch.delay, Me.TblSchBarTanpaWO)
                        End If
                        rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, rowSch.plan_start_time, rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                        PrevFinish = rowSch.plan_end_time
                        PrevFinishRef = PrevFinish
                    End If
                Next

                ''reschedule all items JO above => move to below
                ''catat delay:
                FirstInitial = True
                For Each Row As DataGridViewRow In SelectedRows
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                        Else
                            rowSch.delay = CDec(Me.ObjSch.GetWorkDuration(PrevFinish, rowSch.plan_start_time, Me.PrimaryWU, Me.TblSchBarTanpaWO))
                        End If
                        PrevFinish = rowSch.plan_end_time
                    End If
                Next
                ''reschedule jo reference:
                FirstInitial = True
                For Each Row As DataGridViewRow In SelectedRows
                    Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString + " and [identifier jo] = '" + Row.Cells(Me.colIdentifierJO.Index).Value.ToString + "' and [identifier wo] = '" + Row.Cells(Me.colIdentifierWO.Index).Value.ToString + "'"
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        rowSch = Me.TblSchBar.DefaultView(0).Row
                        If FirstInitial Then
                            FirstInitial = False
                            rowSch.plan_start_time = PrevFinishRef
                        Else
                            rowSch.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, PrevFinish, rowSch.delay, Me.TblSchBarTanpaWO)
                        End If
                        rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, rowSch.plan_start_time, rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                        PrevFinish = rowSch.plan_end_time 
                    End If
                Next

                Me.LoadDgv()
                For Each Row As DataGridViewRow In Me.dgvItem.Rows
                    If Row.Cells(Me.colIdentifierJO.Index).Value.ToString = FirstSelectedJO And _
                    Row.Cells(Me.colIdentifierWO.Index).Value.ToString = FirstSelectedWO Then
                        Me.dgvItem.ClearSelection()
                        Me.dgvItem.Rows(Row.Index).Selected = True
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Row.Index
                        For i As Integer = Row.Index To Row.Index + SelectedCout - 1
                            Dim Row_ As DataGridViewRow = Me.dgvItem.Rows(i)
                            Row_.Cells(Me.ColModify.Index).Value = Boolean.TrueString
                        Next
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Resume_()
        End Try
    End Sub

    Private Function GetFirstReferenceRowIndex_Down(ByRef vReferenceCount As Integer) As Integer
        Dim IdentifierJO_Ref As String = ""
        Dim SelectedFound As Boolean = False
        Dim FirstReferenceRowIndex_Down As Integer = Me.dgvItem.Rows.Count - 1
        Try
            vReferenceCount = 0
            For Each Row As DataGridViewRow In Me.dgvItem.Rows
                If Not SelectedFound Then
                    If Row.Cells(Me.ColModify.Index).Value.ToString = Boolean.TrueString Then
                        SelectedFound = True
                    End If
                Else
                    If Row.Cells(Me.ColModify.Index).Value.ToString = Boolean.FalseString Then
                        FirstReferenceRowIndex_Down = Row.Index
                        IdentifierJO_Ref = Row.Cells(Me.colIdentifierJO.Index).Value.ToString
                        Exit For
                    End If
                End If
            Next
            For i As Integer = FirstReferenceRowIndex_Down To Me.dgvItem.Rows.Count - 1
                Dim Row As DataGridViewRow = Me.dgvItem.Rows(i)
                If Row.Cells(Me.colIdentifierJO.Index).Value.ToString = IdentifierJO_Ref Then
                    vReferenceCount += 1
                Else
                    Exit For
                End If
            Next
            Return FirstReferenceRowIndex_Down
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanDataCodeNoDescUnitCode()
        Dim PrimaryMOMin As Double = 9999999999
        Dim PrimaryMOMax As Double = 0
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TBL As New DseTrans_Schedule.tbltrans_mo_item_woDataTable
        Dim row As DseTrans_Schedule.tbltrans_mo_item_woRow
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim ListPrimaryUnit As New List(Of Double)
        Dim DTAUnit As New DseTrans_ScheduleTableAdapters.tblmaster_master_unitTableAdapter
        Dim TBLUnit As New DseTrans_Schedule.tblmaster_master_unitDataTable
        Dim rowUnit As DseTrans_Schedule.tblmaster_master_unitRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString

            TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + Me.PrimaryWU.ToString + " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            TblSchBar.DefaultView.Sort = "[plan start time]"

            For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                rowSchBar = DV.Row
                If rowSchBar.primary_mo > PrimaryMOMax Then
                    PrimaryMOMax = rowSchBar.primary_mo
                End If
                If PrimaryMOMin > rowSchBar.primary_mo Then
                    PrimaryMOMin = rowSchBar.primary_mo
                End If
            Next

            DTA.Fill(TBL, PrimaryMOMin, PrimaryMOMax)
            DTAUnit.Fill(TBLUnit)
            For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                rowSchBar = DV.Row
                For Each row In TBL.Rows
                    If rowSchBar.identifier_jo = row.Identifier_JO And rowSchBar.identifier_wo = row.Identifier_WO Then
                        rowSchBar.product_code = row.Product_Code
                        rowSchBar.product_desc = row.Description
                        Exit For
                    End If
                Next
                For Each rowUnit In TBLUnit.Rows
                    If rowSchBar.primary_unit = rowUnit.Primary Then
                        rowSchBar.unit_code = rowUnit.Unit_Code
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsValidModify(ByVal vMoveUp As Boolean, ByVal vJOLevel As Boolean) As Boolean
        Dim lastModify As Integer = -2
        Dim lastJONo As String = String.Empty
        Dim FirmPlanned As Date
        Dim ListOfLoggedSch As New List(Of String)
        Dim objLog As New ClsTrans_StockBundling
        Dim RowSelectedBefore As Integer = -1
        Dim LastSelectedRow As Integer = -1
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColModify.Index).Value = True Then
                    If vMoveUp And (row.Index = 0 Or (row.Index = 1 And vJOLevel)) Then
                        MsgBox("No more space above.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If lastModify <> -2 And row.Index - lastModify > 1 Then
                        MsgBox("You need to reorder a sequential schedule.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    lastModify = row.Index

                    If vJOLevel And lastJONo <> String.Empty And row.Cells(Me.ColJONo.Index).Value <> lastJONo Then
                        MsgBox("You need to reorder the schedule that same JO.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    lastJONo = row.Cells(Me.ColJONo.Index).Value

                    If Not vMoveUp And (row.Index = Me.dgvItem.Rows.Count - 1 Or (row.Index = Me.dgvItem.Rows.Count - 2 And vJOLevel)) Then
                        MsgBox("No more space below.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            If lastModify = -2 Then
                MsgBox("Select at least 1 item to modify.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''batasan firm planned dan schedule yang sudah ada log
            FirmPlanned = Me.ObjSch.GetFirmPlan
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                If RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    If RowSelectedBefore = -1 Then RowSelectedBefore = RowDgv.Index - 1
                    LastSelectedRow = RowDgv.Index + 1
                    If CDate(RowDgv.Cells(Me.ColStartDate.Index).Value) <= FirmPlanned Then
                        MsgBox("The selected schedule is firm planned, you can't move the schedule.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If objLog.GetJml_ByJO_WO_PrimaryWorkUnit(RowDgv.Cells(Me.colIdentifierJO.Index).Value, RowDgv.Cells(Me.colIdentifierWO.Index).Value, Me.PrimaryWU) > 0 Then
                        If ListOfLoggedSch.Contains("JO: " & RowDgv.Cells(Me.ColJONo.Index).Value & " WO: " & RowDgv.Cells(Me.ColWONo.Index).Value) = False Then ListOfLoggedSch.Add("JO: " & RowDgv.Cells(Me.ColJONo.Index).Value & " WO: " & RowDgv.Cells(Me.ColWONo.Index).Value)
                    End If
                End If
            Next
            If ListOfLoggedSch.Count > 0 Then
                Dim LoggedSch As String = ""
                For Each s As String In ListOfLoggedSch
                    If LoggedSch = "" Then
                        LoggedSch = s
                    Else
                        LoggedSch = vbNewLine & s
                    End If
                Next
                MsgBox("The selected schedule is refered by 1 or more production log, you can't move the schedule." & vbNewLine & LoggedSch, MsgBoxStyle.Information, "Information")
                Return False
            End If
            If vMoveUp Then
                Dim RowDgv As DataGridViewRow = Me.dgvItem.Rows(RowSelectedBefore)
                If CDate(RowDgv.Cells(Me.ColStartDate.Index).Value) <= FirmPlanned Then
                    MsgBox("The schedule before first selected schedule is firm planned, you can't insert the schedule.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If objLog.GetJml_ByJO_WO_PrimaryWorkUnit(RowDgv.Cells(Me.colIdentifierJO.Index).Value, RowDgv.Cells(Me.colIdentifierWO.Index).Value, Me.PrimaryWU) > 0 Then
                    MsgBox("The schedule before first selected schedule is refered by 1 or more production log, you can't insert the schedule.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                Dim RowDgv As DataGridViewRow = Me.dgvItem.Rows(LastSelectedRow)
                If objLog.GetJml_ByJO_WO_PrimaryWorkUnit(RowDgv.Cells(Me.colIdentifierJO.Index).Value, RowDgv.Cells(Me.colIdentifierWO.Index).Value, Me.PrimaryWU) > 0 Then
                    MsgBox("The schedule after last selected schedule is refered by 1 or more production log, you can't insert the schedule.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function 

    Function GetSelectedDgvRows(ByVal vMoveUp As Boolean) As List(Of DataGridViewRow)
        GetSelectedDgvRows = New List(Of DataGridViewRow)
        Try
            If vMoveUp Then
                For Each row As DataGridViewRow In Me.dgvItem.Rows
                    If row.Cells(Me.ColModify.Index).Value = True Then
                        GetSelectedDgvRows.Add(row)
                    End If
                Next
            Else
                Dim row As DataGridViewRow
                For i As Integer = Me.dgvItem.Rows.Count - 1 To 0 Step -1
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColModify.Index).Value = True Then
                        GetSelectedDgvRows.Add(row)
                    End If
                Next
            End If
            Return GetSelectedDgvRows
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetSelectedDgvRows
        End Try
    End Function

    Sub LoadDgv()
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.IsikanDataCodeNoDescUnitCode()

            'load data awal
            TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + Me.PrimaryWU.ToString
            TblSchBar.DefaultView.Sort = "[plan start time]"
            Me.dgvItem.Rows.Clear()
            For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                rowSchBar = DV.Row
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColCodeNo.Index).Value = rowSchBar.product_code
                    .Cells(Me.ColDesc.Index).Value = rowSchBar.product_desc
                    .Cells(Me.ColEndDate.Index).Value = Format(rowSchBar.plan_end_time, "dd MMM yyyy")
                    .Cells(Me.ColEndDay.Index).Value = Me.ObjCD.GetHari_English(rowSchBar.plan_end_time.Date.DayOfWeek)
                    .Cells(Me.ColEndHour.Index).Value = Format(rowSchBar.plan_end_time, "HH.mm")
                    .Cells(Me.ColJONo.Index).Value = rowSchBar.jo_no
                    .Cells(Me.ColModify.Index).Value = Boolean.FalseString
                    .Cells(Me.ColMoNo.Index).Value = rowSchBar.mo_no
                    .Cells(Me.ColProcessCode.Index).Value = rowSchBar.process_code
                    .Cells(Me.ColProgress.Index).Value = Me.ObjCD.DigitGrouping(rowSchBar.progress, 2)
                    .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowSchBar.qty_produce, rowSchBar.decimal_point)
                    .Cells(Me.ColRate.Index).Value = Me.ObjCD.DigitGrouping(rowSchBar.rate, rowSchBar.decimal_point)
                    .Cells(Me.ColStartDate.Index).Value = Format(rowSchBar.plan_start_time, "dd MMM yyyy")
                    .Cells(Me.ColStartDay.Index).Value = Me.ObjCD.GetHari_English(rowSchBar.plan_start_time.Date.DayOfWeek)
                    .Cells(Me.ColStartHour.Index).Value = Format(rowSchBar.plan_start_time, "HH.mm")
                    .Cells(Me.ColUnit.Index).Value = rowSchBar.unit_code
                    .Cells(Me.ColWH.Index).Value = Me.ObjCD.Convert_ToString(rowSchBar.total_work_hour)
                    .Cells(Me.ColWONo.Index).Value = rowSchBar.wo_no
                    .Cells(Me.colIdentifierJO.Index).Value = rowSchBar.identifier_jo
                    .Cells(Me.colIdentifierWO.Index).Value = rowSchBar.identifier_wo
                    .Cells(Me.ColPlanStartTime.Index).Value = rowSchBar.plan_start_time
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.ClearSelection()
            If Me.dgvItem.Rows.Count = 1 Then
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Me.btnUpJO.Enabled = False
                Me.btnDownJO.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub Pause()
        Me.Cursor = Cursors.WaitCursor
        Me.btnUp.Enabled = False
        Me.btnDown.Enabled = False
        Me.btnUpJO.Enabled = False
        Me.btnDownJO.Enabled = False
        Me.BtnOK.Enabled = False
    End Sub

    Sub Resume_()
        Me.btnUp.Enabled = True
        Me.btnDown.Enabled = True
        Me.btnUpJO.Enabled = True
        Me.btnDownJO.Enabled = True
        Me.BtnOK.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub LlbColumnsSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbColumnsSetting.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_Reorder_ColSetting
        Try
            ij.UserCode = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.SetColumnVisibility()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetColumnVisibility()
        Try
            ShowMO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowMO, Me.UserCode)
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowJO, Me.UserCode)
            ShowWO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWO, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProcess, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDesc, Me.UserCode)
            ShowRate = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowRate, Me.UserCode)
            ShowQty = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowQty, Me.UserCode)
            ShowUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowUnit, Me.UserCode)
            ShowWH = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowWH, Me.UserCode)
            ShowDayStart = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayStart, Me.UserCode)
            ShowDayFinish = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowDayFinish, Me.UserCode)
            ShowProgress = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_Reorder_ShowProgress, Me.UserCode)
            If ShowMO = "1" Or ShowMO = "" Then
                Me.ColMoNo.Visible = True
            Else
                Me.ColMoNo.Visible = False
            End If
            If ShowJO = "1" Or ShowJO = "" Then
                Me.ColJONo.Visible = True
            Else
                Me.ColJONo.Visible = False
            End If
            If ShowWO = "1" Or ShowWO = "" Then
                Me.ColWONo.Visible = True
            Else
                Me.ColWONo.Visible = False
            End If
            If ShowProcess = "1" Or ShowProcess = "" Then
                Me.ColProcessCode.Visible = True
            Else
                Me.ColProcessCode.Visible = False
            End If
            If ShowCodeNo = "1" Or ShowCodeNo = "" Then
                Me.ColCodeNo.Visible = True
            Else
                Me.ColCodeNo.Visible = False
            End If
            If ShowDesc = "1" Or ShowDesc = "" Then
                Me.ColDesc.Visible = True
            Else
                Me.ColDesc.Visible = False
            End If
            If ShowRate = "1" Or ShowRate = "" Then
                Me.ColRate.Visible = True
            Else
                Me.ColRate.Visible = False
            End If
            If ShowQty = "1" Or ShowQty = "" Then
                Me.ColQty.Visible = True
            Else
                Me.ColQty.Visible = False
            End If
            If ShowUnit = "1" Or ShowUnit = "" Then
                Me.ColUnit.Visible = True
            Else
                Me.ColUnit.Visible = False
            End If
            If ShowWH = "1" Or ShowWH = "" Then
                Me.ColWH.Visible = True
            Else
                Me.ColWH.Visible = False
            End If
            If ShowDayStart = "1" Or ShowDayStart = "" Then
                Me.ColStartDay.Visible = True
                Me.ColStartDay.HeaderText = "Start"
                Me.ColStartDate.HeaderText = ""
            Else
                Me.ColStartDay.Visible = False
                Me.ColStartDay.HeaderText = ""
                Me.ColStartDate.HeaderText = "Start"
            End If
            If ShowDayFinish = "1" Or ShowDayFinish = "" Then
                Me.ColEndDay.Visible = True
                Me.ColEndDay.HeaderText = "Finish"
                Me.ColEndDate.HeaderText = ""
            Else
                Me.ColEndDay.Visible = False
                Me.ColEndDay.HeaderText = ""
                Me.ColEndDate.HeaderText = "Finish"
            End If
            If ShowProgress = "1" Or ShowProgress = "" Then
                Me.ColProgress.Visible = True
            Else
                Me.ColProgress.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                RowDgv.Cells(Me.ColModify.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim SelectedIdentifierJO As String = String.Empty
        Try
            If Me.IsValidToSelectAll(SelectedIdentifierJO) = True Then
                For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                    If RowDgv.Cells(Me.colIdentifierJO.Index).Value = SelectedIdentifierJO Then
                        RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsValidToSelectAll(ByRef vSelectedIdentifierJO As String) As Boolean
        Dim SelectedIdentifierJO As String = String.Empty
        Dim SelectedCount As Integer = 0
        Try

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                If RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    SelectedCount += 1
                    If SelectedIdentifierJO <> String.Empty And RowDgv.Cells(Me.colIdentifierJO.Index).Value <> SelectedIdentifierJO Then
                        MsgBox("Please select only one JO Number as reference.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    SelectedIdentifierJO = RowDgv.Cells(Me.colIdentifierJO.Index).Value
                End If
            Next
            If SelectedCount = 0 Then
                MsgBox("Please select at least one schedule as reference.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            vSelectedIdentifierJO = SelectedIdentifierJO
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.Label2.Text = "Loading schedule..."
            Me.LoadTblSch(Me.ObjSch.GetFirmPlan().AddDays(1).Date, #12/31/2200#)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadTblSch(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim MOMin, MOMax As Double
        Dim DTA As New DseTrans_ScheduleTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTrans_Schedule.DataTable2DataTable
        Dim row As DseTrans_Schedule.DataTable2Row
        Dim mn As Double = 0
        Dim jk As Double = 0
        Dim rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_workunitDataTable
        Dim rowWU As DseTrans_Schedule.tblmaster_workunitRow
        Dim DTAWU1 As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunit1TableAdapter
        Dim TBLWU1 As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim StartTime As Date = Now
        Dim Ketemu As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAWU.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWU1.Connection.ConnectionString = DTA.Connection.ConnectionString

            Me.TblSchBar.Rows.Clear()
            Me.ObjSch.LoadScheduleBar_WOItemOnly(vStartDate, vEndDate, Me.TblSchBar)
            Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(vStartDate, vEndDate, Me.TblSchBarTanpaWO, Me.PrimaryWU)

            Dim ij As Integer = 0
            While ij <= TblSchBar.Rows.Count - 1
                rowSch = TblSchBar.Rows(ij)
                If rowSch.tipe <> ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                    Me.TblSchBar.Rows.Remove(rowSch)
                Else
                    ij += 1
                End If
            End While

            'isikan nilai yang dibutuhkan
            If Me.TblSchBar.Rows.Count > 0 Then
                MOMin = Double.MaxValue
                MOMax = Double.MinValue
                For Each rowSch In Me.TblSchBar.Rows
                    If rowSch.primary_mo <> -1 Then
                        If rowSch.primary_mo < MOMin Then
                            MOMin = rowSch.primary_mo
                        End If
                        If rowSch.primary_mo > MOMax Then
                            MOMax = rowSch.primary_mo
                        End If
                    End If
                Next

                DTA.Fill(TBL, MOMax, MOMin, vStartDate)

                While mn <= Me.TblSchBar.Rows.Count - 1
                    Ketemu = False
                    rowSch = Me.TblSchBar.Rows(mn)
                    If rowSch.primary_mo <> -1 Then
                        If rowSch.wo_no = "" Then
                            For Each row In TBL.Rows
                                If row.Identifier_WO = rowSch.identifier_wo And rowSch.wo_no = "" Then
                                    rowSch.wo_no = row.WO_Number
                                    rowSch.jo_no = row.JO_Number
                                    rowSch.process_code = row.Process_Code
                                    rowSch.mo_no = row.MO_Number
                                    rowSch.unit_code = row.Unit_Code
                                    rowSch.is_closed = row.Is_Closed

                                    jk = mn + 1
                                    While jk <= Me.TblSchBar.Rows.Count - 1
                                        rowSch1 = Me.TblSchBar.Rows(jk)
                                        If rowSch1.identifier_wo = rowSch.identifier_wo Then
                                            rowSch1.wo_no = row.WO_Number
                                            rowSch1.jo_no = row.JO_Number
                                            rowSch1.process_code = row.Process_Code
                                            rowSch1.mo_no = row.MO_Number
                                            rowSch1.unit_code = row.Unit_Code
                                            rowSch1.is_closed = row.Is_Closed
                                        End If
                                        jk += 1
                                    End While

                                    TBL.Rows.Remove(row)
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    mn += 1
                    If Ketemu = False And rowSch.wo_no = "" Then
                        MsgBox("Logic error: additional data not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End While
                TBL.Rows.Clear()

                mn = 0
                DTAWU.Fill(TBLWU)
                While mn <= Me.TblSchBar.Rows.Count - 1
                    Ketemu = False
                    rowSch = Me.TblSchBar.Rows(mn)
                    If rowSch.identifier_wo <> "" Then
                        If rowSch.work_unit_code = "" Then
                            For Each rowWU In TBLWU.Rows
                                If rowWU.primary = rowSch.primary_work_unit And rowSch.work_unit_code = "" Then
                                    rowSch.work_unit_code = rowWU.work_unit_code

                                    jk = mn + 1
                                    While jk <= Me.TblSchBar.Rows.Count - 1
                                        rowSch1 = Me.TblSchBar.Rows(jk)
                                        If rowSch1.primary_work_unit = rowSch.primary_work_unit Then
                                            rowSch1.work_unit_code = rowWU.work_unit_code
                                        End If
                                        jk += 1
                                    End While

                                    Ketemu = True
                                    TBLWU.Rows.Remove(rowWU)
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    mn += 1
                    If Ketemu = False And rowSch.work_unit_code = "" Then
                        MsgBox("Logic error: work unit data not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End While
                TBLWU.Rows.Clear()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.LoadDgv()
            Me.ShowItem()
            Me.EnableControl(True)
            Me.SetColumnVisibility()
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub EnableControl(ByVal vEnabledControl As Boolean)
        Try
            If vEnabledControl Then
                Me.Panel1.Visible = False
                Me.ControlBox = True
                Me.LlbColumnsSetting.Enabled = True
                Me.LinkLabel1.Enabled = True
                Me.LinkLabel2.Enabled = True
                Me.dgvItem.ReadOnly = False
                Me.btnUp.Enabled = True
                Me.btnDown.Enabled = True
                Me.btnUpJO.Enabled = True
                Me.btnDownJO.Enabled = True
                Me.BtnOK.Enabled = True
                Me.BtnCancel.Enabled = True
            Else
                Me.Panel1.Visible = True
                Me.ControlBox = False
                Me.LlbColumnsSetting.Enabled = False
                Me.LinkLabel1.Enabled = False
                Me.LinkLabel2.Enabled = False
                Me.dgvItem.ReadOnly = True
                Me.btnUp.Enabled = False
                Me.btnDown.Enabled = False
                Me.btnUpJO.Enabled = False
                Me.btnDownJO.Enabled = False
                Me.BtnOK.Enabled = False
                Me.BtnCancel.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class