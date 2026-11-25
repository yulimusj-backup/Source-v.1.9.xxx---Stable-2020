Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_EditOT2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public IsEdited As Boolean = False

    Public Enum EnumGroup As Integer
        WorkUnitGroup
        WorkUnit
    End Enum

    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWUGroup As New Core_Module.ClsMasterWorkUnitGroup
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim ObjSch As New ClsTrans_Schedule

    Private Sub dgvFilterWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvFilterWU.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvFilterWU.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvFilterWU.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColFilterWU_Code.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColFilterWU_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvFilterWU.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code work_unit group dan level
                    If e.ColumnIndex = Me.ColFilterWU_Code.Index Or e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                        Me.buffer = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                        Me.buffer = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvFilterWU.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvFilterWU)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFilterWU.CellEndEdit
        Dim Level As EnumGroup = EnumGroup.WorkUnit
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell

        Try
            If IsNothing(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                    'kalo kode work_unit group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColFilterWU_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColFilterWU_Code.Index Then
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek ke master
                    CboDgv = Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex)
                    Level = CboDgv.Items.IndexOf(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)

                    If Level = EnumGroup.WorkUnitGroup Then
                        Me.EndEditWorkUnitGroup(e)
                    Else
                        Me.EndEditwork_unit(e)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditWorkUnitGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjWUGroup.Apakahsudahada(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode work unit group ada  --> load data nama + primary
                Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWUGroup.GetPrimary(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
            Else
                '2. kode work unit group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmMasterWorkUnitGroup_Cari
                ij.txtCari.Text = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryWorkUnitGroup <> -1 Then
                    '2.a. work unit group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnitGroup
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjWUGroup.GetCode(ij.PrimaryWorkUnitGroup)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(ij.PrimaryWorkUnitGroup)
                ElseIf ij.PrimaryWorkUnitGroup = -1 Then
                    '2.b. kode work unit group yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemWorkUnit(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWUGroup.GetPrimary(Me.buffer)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWUGroup.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditwork_unit(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjWU.ApakahSudahAda(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode work_unit ada  --> load data nama + primary
                Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWU.GetPrimary(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
            Else
                '2. kode work_unit yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                ij.txtCari.Text = Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryWorkUnit <> -1 Then
                    '2.a. work_unit yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjWU.GetCode(ij.PrimaryWorkUnit)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(ij.PrimaryWorkUnit)
                ElseIf ij.PrimaryWorkUnit = -1 Then
                    '2.b. kode work_unit yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemWorkUnit(Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvFilterWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value = Me.ObjWU.GetPrimary(Me.buffer)
                    Me.dgvFilterWU.Item(Me.ColFilterWU_Desc.Index, e.RowIndex).Value = Me.ObjWU.GetDesc(Me.dgvFilterWU.Item(Me.ColFilterWU_Primary.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemWorkUnit(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvFilterWU.Rows
                If rowDGV.Cells(Me.ColFilterWU_Code.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColFilterWU_Level.Index).Value = vLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvFilterWU_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvFilterWU.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvFilterWU_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFilterWU.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvFilterWU.Rows(Me.dgvFilterWU.Rows.Count - 1)
            rowDGV.Cells(Me.ColFilterWU_Code.Index).Value = ""
            rowDGV.Cells(Me.ColFilterWU_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColFilterWU_Level.Index).Value = "Work Unit Group"
            rowDGV.Cells(Me.ColFilterWU_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterWU_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvFilterWU.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvFilterWU)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim DTA As New DseTrans_OT_ItemTableAdapters.tblmaster_workunitTableAdapter
        Dim TBL As New DseTrans_OT_Item.tblmaster_workunitDataTable
        Dim row As DseTrans_OT_Item.tblmaster_workunitRow
        Dim TBL1 As New DseTrans_OT_Item.tbltrans_overtime_itemDataTable
        Dim row1 As DseTrans_OT_Item.tbltrans_overtime_itemRow
        Dim RefDate As Date = Me.dtpStartDate.Value.Date
        Dim StartOT, EndOT As Date
        Dim ListPrimaryOvertime As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Level As EnumGroup = EnumGroup.WorkUnit

        Try
            If Me.ApakahDataBenar = True Then
                If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                'simpan data overtime
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                For Each rowDGV In Me.dgvFilterWU.Rows
                    If rowDGV.Index <> Me.dgvFilterWU.RowCount - 1 Then
                        CboDgv = Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index)
                        Level = CboDgv.Items.IndexOf(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index).Value)
                        If Level = EnumGroup.WorkUnit Then
                            If ListPrimaryOvertime.Contains(rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value) = False Then
                                ListPrimaryOvertime.Add(rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value)
                            End If
                        Else
                            TBL.Rows.Clear()
                            DTA.Fill(TBL, rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value)
                            For Each row In TBL.Rows
                                If ListPrimaryOvertime.Contains(row.primary) = False Then
                                    ListPrimaryOvertime.Add(row.primary)
                                End If
                            Next
                        End If
                    End If
                Next

                RefDate = Me.dtpStartDate.Value.Date
                While RefDate.Date <= Me.dtpEndDate.Value.Date
                    If Me.dtpStartTime.Value < Me.dtpEndTime.Value Then
                        StartOT = RefDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                        EndOT = RefDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute)
                    Else
                        StartOT = RefDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                        EndOT = RefDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute) + TimeSpan.FromDays(1)
                    End If
                    For Each PrimaryWU As Double In ListPrimaryOvertime
                        row1 = TBL1.Newtbltrans_overtime_itemRow
                        row1.primary = -1
                        row1.primary_spl = -1
                        row1.primary_workunit = PrimaryWU
                        row1.start = StartOT
                        row1._end = EndOT
                        TBL1.Rows.Add(row1)
                    Next
                    RefDate = RefDate + TimeSpan.FromDays(1)
                End While

                Me.Cursor = Cursors.WaitCursor
                Dim objSchOT As New ClsTrans_Schedule_OT
                Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
                Dim vTBLSpaceBeforeOT As New DseTrans_Schedule.tbl_sch_barDataTable
                ObjSch.LoadScheduleBar_WOItemOnly(#1/1/1950#, #12/31/2200#, vTBLSchBar)
                ObjSch.LoadScheduleBar_TanpaScheduledItem(ObjSch.GetFirmPlan, #12/31/2200#, vTBLSpaceBeforeOT)
                If objSchOT.AnyClosedJO(vTBLSchBar, Me.dtpStartDate.Value.Date.AddHours(Me.dtpEndTime.Value.Hour).AddMinutes(Me.dtpEndTime.Value.Minute), ListPrimaryOvertime) Then
                    Exit Sub
                End If

                Dim message As String = ""
                If objSchOT.AnyProducedJO(vTBLSchBar, Me.dtpStartDate.Value.Date.AddHours(Me.dtpEndTime.Value.Hour).AddMinutes(Me.dtpEndTime.Value.Minute), ListPrimaryOvertime, message) Then
                    MsgBox("Cannot add overtime, there are log transaction for the following JO:" & vbNewLine & message & ".", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.SimpanData(TBL1) = True Then
                    For Each PrimaryWU As Double In ListPrimaryOvertime
                        objSchOT.RecalculateAfterAddSpace(PrimaryWU, Me.dtpStartDate.Value.Date.AddHours(Me.dtpEndTime.Value.Hour).AddMinutes(Me.dtpEndTime.Value.Minute), vTBLSpaceBeforeOT, vTBLSchBar)
                    Next
                    Me.IsEdited = True
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            End If
        Catch ex As Exception

            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            ListPrimaryOvertime = Nothing
            DTA.Dispose()
            TBL.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Function SimpanData(ByVal vTBL As DseTrans_OT_Item.tbltrans_overtime_itemDataTable) As Boolean
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
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vTBL)
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
    ByVal vTBL As DseTrans_OT_Item.tbltrans_overtime_itemDataTable) As String
        Dim retval As String = ""
        Dim row As DseTrans_OT_Item.tbltrans_overtime_itemRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_overtime_item` (`primary workunit`,`primary spl`,`start`,`end`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'If row.qty_produce + row.qty_produce_add > 0 Then
                    'hanya simpan yang ada qty producenya saja
                        'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(.primary_workunit).ToString + "," _
                        + CDbl(.primary_spl).ToString + "," _
                        + "'" + Format(CDate(.start), "yyyy-MM-dd HH:mm:ss") + "'," _
                        + "'" + Format(CDate(._end), "yyyy-MM-dd HH:mm:ss") + "'" _
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

    Private Function ApakahDataBenar() As Boolean
        Dim ListPrimaryOvertime As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Level As EnumGroup = EnumGroup.WorkUnit
        Dim DTA As New DseTrans_OT_ItemTableAdapters.tblmaster_workunitTableAdapter
        Dim TBL As New DseTrans_OT_Item.tblmaster_workunitDataTable
        Dim row As DseTrans_OT_Item.tblmaster_workunitRow
        Dim DTA1 As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBL1 As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim row1 As DseTrans_OT_Item.tbltrans_mo_workshift_allocateRow
        Dim DTA2 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL2 As New DseTrans_OT_Item.tblmaster_workshift_itemDataTable
        Dim row2 As DseTrans_OT_Item.tblmaster_workshift_itemRow
        Dim DTA3 As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim TBL3 As New DseTrans_OT_Item.tbltrans_overtime_itemDataTable
        Dim row3 As DseTrans_OT_Item.tbltrans_overtime_itemRow
        Dim DTA4 As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBL4 As New DseTrans_OT_Item.tblmaster_workshift_itemholidayDataTable
        Dim CekDate As Date = Today
        Dim TBLCek As New DseTrans_OT_Item.TblCekDataTable
        Dim rowCek, rowCek1 As DseTrans_OT_Item.TblCekRow
        Dim Ketemu As Boolean = False
        Dim WUCode As String = ""
        Dim WUDesc As String = ""
        Dim StartOT, EndOT, SchIn, SchOut As Date

        Try
            'ketika set overtime, pastikan:
            '-3. jika start time < end time -> berikan peringatan
            '-2. start date < firm plan
            '-1. start date < end date
            '0. minimal ada 1 work unit
            '1. work unit ybs untuk tanggal ybs harus sudah allocate work shift
            '2. overtime tidak boleh bersinggungan dengan workshift
            '3. tidak boleh conflict dengan overtime yang sudah ada

            '-3. jika start time < end time -> berikan peringatan
            If Me.dtpStartTime.Value > Me.dtpEndTime.Value Then
                If MsgBox("The start time < end time, this indicates worktime happens in 2 consecutive date, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            ElseIf Me.dtpStartTime.Value = Me.dtpEndTime.Value Then
                If MsgBox("The start time = end time, please fix.", MsgBoxStyle.Information, "Information") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            '-2. start date < firm plan
            If Me.dtpStartDate.Value.Date <= Me.ObjSch.GetFirmPlan Then
                MsgBox("Please specify start date > firm plan date (" + Format(Me.ObjSch.GetFirmPlan, "dd MMM yyyy") + ").", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '-1. start date < end date
            If Me.dtpStartDate.Value.Date > Me.dtpEndDate.Value.Date Then
                MsgBox("Please specify end date > start date.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '0. minimal ada 1 work unit
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString

            For Each rowDGV In Me.dgvFilterWU.Rows
                If rowDGV.Index <> Me.dgvFilterWU.RowCount - 1 Then
                    CboDgv = Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index)
                    Level = CboDgv.Items.IndexOf(Me.dgvFilterWU.Item(Me.ColFilterWU_Level.Index, rowDGV.Index).Value)
                    If Level = EnumGroup.WorkUnit Then
                        If ListPrimaryOvertime.Contains(rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value) = False Then
                            ListPrimaryOvertime.Add(rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value)
                        End If
                    Else
                        TBL.Rows.Clear()
                        DTA.Fill(TBL, rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value)
                        For Each row In TBL.Rows
                            If ListPrimaryOvertime.Contains(row.primary) = False Then
                                ListPrimaryOvertime.Add(row.primary)
                            End If
                        Next
                    End If
                End If
            Next
            If ListPrimaryOvertime.Count = 0 Then
                MsgBox("Please specify at least 1 work unit.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. work unit ybs untuk tanggal ybs harus sudah allocate work shift
            TBL1.Rows.Clear()
            DTA1.Fill(TBL1, Me.dtpEndDate.Value.Date, Me.dtpStartDate.Value.Date)
            For Each PrimaryWU As Double In ListPrimaryOvertime
                WUCode = Me.ObjWU.GetCode(PrimaryWU)
                WUDesc = Me.ObjWU.GetDesc(PrimaryWU)
                TBL1.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                TBL1.DefaultView.Sort = "Date"
                CekDate = Me.dtpStartDate.Value.Date

                While CekDate <= Me.dtpEndDate.Value.Date
                    Ketemu = False
                    For ij As Double = 0 To TBL1.DefaultView.Count - 1
                        row1 = TBL1.DefaultView(ij).Row
                        If row1._date = CekDate Then
                            Ketemu = True
                            Exit For
                        ElseIf row1._date > CekDate Then
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        'tambahkan baris cek
                        rowCek = TBLCek.NewTblCekRow
                        rowCek.Primary_Work_Unit = PrimaryWU
                        rowCek.Ref_Date = CekDate
                        rowCek.WU_Code = WUCode
                        rowCek.WU_Desc = WUDesc

                        TBLCek.Rows.Add(rowCek)
                    End If

                    CekDate = CekDate + TimeSpan.FromDays(1)
                End While
            Next
            If TBLCek.Rows.Count > 0 Then
                Dim FrmCek As New FrmTrans_ProductionSchedule_EditOT3
                TBLCek.DefaultView.Sort = "[WU Code], [Ref Date]"
                TBLCek.DefaultView.RowFilter = ""

                For Each DV As Data.DataRowView In TBLCek.DefaultView
                    rowCek = DV.Row
                    If IsNothing(rowCek1) = False Then
                        If rowCek1.Primary_Work_Unit = rowCek.Primary_Work_Unit And rowCek1.Ref_Date = rowCek.Ref_Date - TimeSpan.FromDays(1) Then
                            rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                            rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                        Else
                            FrmCek.dgvCek.Rows.Add()
                            rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                            rowDGV.Cells(FrmCek.ColDate.Index).Value = rowCek.Ref_Date
                            rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                            rowDGV.Cells(FrmCek.ColNo.Index).Value = rowDGV.Index + 1
                            rowDGV.Cells(FrmCek.ColWUCode.Index).Value = rowCek.WU_Code
                            rowDGV.Cells(FrmCek.ColWUDesc.Index).Value = rowCek.WU_Desc
                        End If
                    Else
                        FrmCek.dgvCek.Rows.Add()
                        rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                        rowDGV.Cells(FrmCek.ColDate.Index).Value = rowCek.Ref_Date
                        rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                        rowDGV.Cells(FrmCek.ColNo.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(FrmCek.ColWUCode.Index).Value = rowCek.WU_Code
                        rowDGV.Cells(FrmCek.ColWUDesc.Index).Value = rowCek.WU_Desc
                    End If
                    rowCek1 = rowCek
                Next
                FrmCek.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                FrmCek.ColDate1.DefaultCellStyle.Format = "dd MMM yyyy"
                FrmCek.ShowDialog()
                FrmCek.Dispose()
                Return False
            End If

            '2. overtime tidak boleh bersinggungan dengan workshift
            'isikan day of week, sch in, dan sch out
            TBLCek.Rows.Clear()
            For Each row1 In TBL1.Rows
                row1.day_of_week = row1._date.DayOfWeek
            Next
            For Each row1 In TBL1.Rows
                If row1.sch_in = -1 Then
                    DTA2.FillBy1(TBL2, row1._date.DayOfWeek, row1.primary_workshift)
                    row2 = TBL2.Rows(0)
                    Me.SetSch(TBL1, row1.primary_workshift, row2.schedule_in, row2.schedule_out, row1.day_of_week, row2.tipe_workdays)
                End If
            Next
            DTA4.Fill(TBL4, Me.dtpStartDate.Value.Date)
            For Each row1 In TBL1.Rows
                TBL4.DefaultView.RowFilter = "[primary transaksi] = " + row1.primary_workshift.ToString + " AND date = #" + Format(row1._date, "MM/dd/yyyy") + "#"
                If TBL4.DefaultView.Count > 0 Then
                    row1.tipe_workdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday
                End If
            Next
            For Each PrimaryWU As Double In ListPrimaryOvertime
                WUCode = Me.ObjWU.GetCode(PrimaryWU)
                WUDesc = Me.ObjWU.GetDesc(PrimaryWU)
                TBL1.DefaultView.RowFilter = "[primary work unit] = " + PrimaryWU.ToString
                TBL1.DefaultView.Sort = "Date"

                CekDate = Me.dtpStartDate.Value.Date
                While CekDate <= Me.dtpEndDate.Value.Date
                    If Me.dtpStartTime.Value < Me.dtpEndTime.Value Then
                        StartOT = CekDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                        EndOT = CekDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute)
                    Else
                        StartOT = CekDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                        EndOT = CekDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute) + TimeSpan.FromDays(1)
                    End If
                    Ketemu = False

                    For ij As Double = 0 To TBL1.DefaultView.Count - 1
                        row1 = TBL1.DefaultView(ij).Row
                        If row1._date = CekDate Then
                            'pastikan tidak bersinggungan
                            If row1.tipe_workdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday Then
                                If row1.sch_in < row1.sch_out Then
                                    SchIn = CekDate + TimeSpan.FromMinutes(row1.sch_in)
                                    SchOut = CekDate + TimeSpan.FromMinutes(row1.sch_out)
                                Else
                                    SchIn = CekDate + TimeSpan.FromMinutes(row1.sch_in)
                                    SchOut = CekDate + TimeSpan.FromMinutes(row1.sch_out) + TimeSpan.FromDays(1)
                                End If
                                If (StartOT >= SchOut And EndOT > SchOut) Or _
                                (EndOT <= SchIn And StartOT < SchIn) Then
                                    Ketemu = True
                                    Exit For
                                End If
                            Else
                                Ketemu = True
                                Exit For
                            End If
                        ElseIf row1._date > CekDate Then
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        'tambahkan baris cek
                        rowCek = TBLCek.NewTblCekRow
                        rowCek.Primary_Work_Unit = PrimaryWU
                        rowCek.Ref_Date = CekDate
                        rowCek.WU_Code = WUCode
                        rowCek.WU_Desc = WUDesc

                        TBLCek.Rows.Add(rowCek)
                    End If

                    CekDate = CekDate + TimeSpan.FromDays(1)
                End While
            Next
            If TBLCek.Rows.Count > 0 Then
                Dim FrmCek As New FrmTrans_ProductionSchedule_EditOT3
                FrmCek.Label1.Text = "The following item conflict with the reguler work time. Please fix."
                FrmCek.Text = "Conflicting with reguler work time"
                TBLCek.DefaultView.Sort = "[WU Code], [Ref Date]"
                TBLCek.DefaultView.RowFilter = ""

                For Each DV As Data.DataRowView In TBLCek.DefaultView
                    rowCek = DV.Row
                    If IsNothing(rowCek1) = False Then
                        If rowCek1.Primary_Work_Unit = rowCek.Primary_Work_Unit And rowCek1.Ref_Date = rowCek.Ref_Date - TimeSpan.FromDays(1) Then
                            rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                            rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                        Else
                            FrmCek.dgvCek.Rows.Add()
                            rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                            rowDGV.Cells(FrmCek.ColDate.Index).Value = rowCek.Ref_Date
                            rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                            rowDGV.Cells(FrmCek.ColNo.Index).Value = rowDGV.Index + 1
                            rowDGV.Cells(FrmCek.ColWUCode.Index).Value = rowCek.WU_Code
                            rowDGV.Cells(FrmCek.ColWUDesc.Index).Value = rowCek.WU_Desc
                        End If
                    Else
                        FrmCek.dgvCek.Rows.Add()
                        rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                        rowDGV.Cells(FrmCek.ColDate.Index).Value = rowCek.Ref_Date
                        rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                        rowDGV.Cells(FrmCek.ColNo.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(FrmCek.ColWUCode.Index).Value = rowCek.WU_Code
                        rowDGV.Cells(FrmCek.ColWUDesc.Index).Value = rowCek.WU_Desc
                    End If
                    rowCek1 = rowCek
                Next
                FrmCek.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                FrmCek.ColDate1.DefaultCellStyle.Format = "dd MMM yyyy"
                FrmCek.ShowDialog()
                FrmCek.Dispose()
                Return False
            End If

            '3. tidak boleh conflict dengan overtime yang sudah ada
            For Each PrimaryWU As Double In ListPrimaryOvertime
                DTA3.FillBy1(TBL3, PrimaryWU, Me.dtpStartDate.Value - TimeSpan.FromDays(1))
                WUCode = Me.ObjWU.GetCode(PrimaryWU)
                WUDesc = Me.ObjWU.GetDesc(PrimaryWU)
                TBL3.DefaultView.RowFilter = "[primary workunit] = " + PrimaryWU.ToString
                TBL3.DefaultView.Sort = "start"

                If TBL3.DefaultView.Count > 0 Then
                    CekDate = Me.dtpStartDate.Value.Date
                    While CekDate.Date <= Me.dtpEndDate.Value.Date
                        'pastikan tidak bersinggungan
                        If Me.dtpStartTime.Value < Me.dtpEndTime.Value Then
                            StartOT = CekDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                            EndOT = CekDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute)
                        Else
                            StartOT = CekDate + TimeSpan.FromMinutes(Me.dtpStartTime.Value.Hour * 60 + Me.dtpStartTime.Value.Minute)
                            EndOT = CekDate + TimeSpan.FromMinutes(Me.dtpEndTime.Value.Hour * 60 + Me.dtpEndTime.Value.Minute) + TimeSpan.FromDays(1)
                        End If
                        Ketemu = False

                        For ij As Double = 0 To TBL3.DefaultView.Count - 1
                            row3 = TBL3.DefaultView(ij).Row
                            If row3.start.Date <= CekDate.Date Then
                                SchIn = row3.start
                                SchOut = row3._end

                                If (StartOT >= SchOut And EndOT > SchOut) Or _
                                (EndOT <= SchIn And StartOT < SchIn) Then
                                Else
                                    Ketemu = True
                                    Exit For
                                End If
                            ElseIf row3.start > CekDate Then
                                Exit For
                            End If
                        Next
                        If Ketemu = True Then
                            'tambahkan baris cek
                            rowCek = TBLCek.NewTblCekRow
                            rowCek.Primary_Work_Unit = PrimaryWU
                            rowCek.Ref_Date = CekDate
                            rowCek.WU_Code = WUCode
                            rowCek.WU_Desc = WUDesc
                            If row3.start.Date = row3._end.Date Then
                                rowCek.Remarks = Format(row3.start, "dd MMM yyyy HH.mm") + " - " + Format(row3._end, "HH.mm")
                            Else
                                rowCek.Remarks = Format(row3.start, "dd MMM yyyy HH.mm") + " - " + Format(row3._end, "dd MMM yyyy HH.mm")
                            End If

                            TBLCek.Rows.Add(rowCek)
                        End If

                        CekDate = CekDate + TimeSpan.FromDays(1)
                    End While
                End If
            Next
            If TBLCek.Rows.Count > 0 Then
                Dim FrmCek As New FrmTrans_ProductionSchedule_EditOT3
                FrmCek.Label1.Text = "The following item conflict with another scheduled overtime. Please fix."
                FrmCek.Text = "Conflicting with scheduled overtime"
                TBLCek.DefaultView.Sort = "[WU Code], [Ref Date]"
                TBLCek.DefaultView.RowFilter = ""

                For Each DV As Data.DataRowView In TBLCek.DefaultView
                    rowCek = DV.Row
                    FrmCek.dgvCek.Rows.Add()
                    rowDGV = FrmCek.dgvCek.Rows(FrmCek.dgvCek.Rows.Count - 1)
                    rowDGV.Cells(FrmCek.ColDate.Index).Value = rowCek.Remarks
                    rowDGV.Cells(FrmCek.ColDate1.Index).Value = rowCek.Ref_Date
                    rowDGV.Cells(FrmCek.ColNo.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(FrmCek.ColWUCode.Index).Value = rowCek.WU_Code
                    rowDGV.Cells(FrmCek.ColWUDesc.Index).Value = rowCek.WU_Desc
                Next
                'FrmCek.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                'FrmCek.ColDate1.DefaultCellStyle.Format = "dd MMM yyyy"
                FrmCek.ColDate1.Visible = False
                FrmCek.ColDate.HeaderText = "Date"
                FrmCek.ShowDialog()
                FrmCek.Dispose()
                Return False
            End If

            Dim dtaBundling As New DseTrans_OT_ItemTableAdapters.tbltransaksi_stockbundlingTableAdapter
            Dim tblBundling As New DseTrans_OT_Item.tbltransaksi_stockbundlingDataTable
            Dim rowBundling As DseTrans_OT_Item.tbltransaksi_stockbundlingRow
            dtaBundling.ClearBeforeFill = False
            dtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            Me.dtpStartDate.Value = Me.dtpStartDate.Value.Date.AddHours(Me.dtpStartTime.Value.Hour).AddMinutes(Me.dtpStartTime.Value.Minute)
            Me.dtpEndDate.Value = Me.dtpEndDate.Value.Date.AddHours(Me.dtpEndTime.Value.Hour).AddMinutes(Me.dtpEndTime.Value.Minute)
            For Each rowDGV In Me.dgvFilterWU.Rows
                If rowDGV.Cells(Me.ColFilterWU_Level.Index).Value = "Work Unit Group" Then
                    dtaBundling.FillBy_PrimaryWorkunitGroup(tblBundling, Me.dtpEndDate.Value, Me.dtpStartDate.Value, rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value)
                Else
                    dtaBundling.Fill(tblBundling, Me.dtpEndDate.Value, rowDGV.Cells(Me.ColFilterWU_Primary.Index).Value, Me.dtpStartDate.Value)
                End If
            Next
            If tblBundling.Count > 0 Then
                Dim msg As String = ""
                For Each rowBundling In tblBundling
                    msg = msg & "Work unit " & rowBundling.work_unit_code & " on production log " & rowBundling.No_Transaksi & ", actual start " & Format(rowBundling.start_time_actual, "dd MMM yyyy HH.mm") & " to " & Format(rowBundling.end_time_actual, "dd MMM yyyy HH.ss") & vbNewLine
                Next
                MsgBox("The production log in the selected period already exists: " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryOvertime = Nothing
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            TBLCek.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            DTA3.Dispose()
            TBL3.Dispose()
            DTA4.Dispose()
            TBL4.Dispose()
        End Try
    End Function

    Private Sub SetSch(ByVal vTBL1 As DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable, ByVal vPrimaryWorkshift As Double, _
    ByVal vSchIn As Double, ByVal vSchOut As Double, ByVal vDayOfWeek As Integer, ByVal vTipeWorkDays As Integer)
        Dim row1 As DseTrans_OT_Item.tbltrans_mo_workshift_allocateRow

        Try
            For Each row1 In vTBL1.Rows
                If row1.day_of_week = vDayOfWeek And row1.primary_workshift = vPrimaryWorkshift Then
                    row1.sch_in = vSchIn
                    row1.sch_out = vSchOut
                    row1.tipe_workdays = vTipeWorkDays
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_EditOT2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim day_of_week As DayOfWeek = DayOfWeek.Friday
        Dim DTA As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBL As New DseTrans_OT_Item.tblmaster_workshift_itemDataTable
        Dim row As DseTrans_OT_Item.tblmaster_workshift_itemRow

        Try
            Me.dgvFilterWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.lblFirmPlanDate.Text = Format(Me.ObjSch.GetFirmPlan, "dd MMM yyyy")

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, Me.dtpStartDate.Value.Date.DayOfWeek)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.tipe_workdays = 0 Then
                    'hari kerja
                    'set overtime 2 jam, setelah keluar untuk hari kerja biasa
                    Me.dtpStartTime.Value = Me.dtpStartDate.Value.Date + TimeSpan.FromMinutes(row.schedule_out)
                    Me.dtpEndTime.Value = Me.dtpStartDate.Value.Date + TimeSpan.FromMinutes(row.schedule_out + 120)
                Else
                    'liburan
                    'set overtime 2 jam setelah jam masuk untuk liburan
                    Me.dtpStartTime.Value = Me.dtpStartDate.Value.Date + TimeSpan.FromMinutes(row.schedule_in)
                    Me.dtpEndTime.Value = Me.dtpStartDate.Value.Date + TimeSpan.FromMinutes(row.schedule_in + 120)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class