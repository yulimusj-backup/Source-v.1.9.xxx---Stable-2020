''@rosadinnn
''201508191421
Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmTrans_Attendance_Log
    Public backcolorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = "USER"
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjSrv As New ClsMasterAttendanceServer
    Dim Obj As New ClsTrans_Attendance_Log
    Dim RowDgv As DataGridViewRow

    Dim TblServer As New DseTrans_Attendance_Log.tblmaster_fingerprint_serverDataTable

    Private Sub FrmTrans_Attendance_Log_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.LoadServerList(True)
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.ChkDeleteBeforeImport.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadServerList(ByVal vActiveOnly As Boolean)
        Dim DtaServer As New DseTrans_Attendance_LogTableAdapters.tblmaster_fingerprint_serverTableAdapter
        Dim rowServer As DseTrans_Attendance_Log.tblmaster_fingerprint_serverRow
        Try
            DtaServer.Connection.ConnectionString = Core_Module.GetConstring
            If vActiveOnly = False Then
                DtaServer.Fill(Me.TblServer)
            Else
                DtaServer.FillBy_ActiveOnly(Me.TblServer)
            End If
            Me.ChkListServer.Items.Clear()
            rowServer = Me.TblServer.Newtblmaster_fingerprint_serverRow
            rowServer.primary = -2
            rowServer.Server = "From File"
            Me.TblServer.Rows.Add(rowServer)
            For Each rowSrv As DseTrans_Attendance_Log.tblmaster_fingerprint_serverRow In Me.TblServer
                Me.ChkListServer.Items.Add(rowSrv.Server, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaServer.Dispose()
        End Try
    End Sub

    Private Sub BtnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLoad.Click
        Dim ListPrimaryActiveServer As New List(Of Double)
        Dim ListActiveTerminalID As New List(Of String)
        Dim ThereIsActiveTerminal As Boolean = False
        Dim ServerName As String = ""
        Dim ServerLocation As String = ""

        Dim DtaLog As New DseTrans_Attendance_LogTableAdapters.NGAC_AUTHLOGTableAdapter
        Dim TblLog As New DseTrans_Attendance_Log.NGAC_AUTHLOGDataTable
        Dim RowLog As DseTrans_Attendance_Log.NGAC_AUTHLOGRow

        Dim DtaLogAcc As New DseTrans_Attendance_LogTableAdapters.DataTable2TableAdapter
        Dim TblLogAcc As New DseTrans_Attendance_Log.DataTable2DataTable
        Dim RowLogAcc As DseTrans_Attendance_Log.DataTable2Row

        Dim DtaEmployee As New DseTrans_Attendance_LogTableAdapters.tblmaster_karyawanTableAdapter
        Dim TblEmployee As New DseTrans_Attendance_Log.tblmaster_karyawanDataTable
        Dim RowEmployee As DseTrans_Attendance_Log.tblmaster_karyawanRow
        Dim EmployeeCode, EmployeeName As String
        Dim EmployeePrimary As Double
        Dim SqlServerConnectionState_ As Boolean = False

        Dim SelectedCount_ As Integer = 0
        Dim RowSrv As DseTrans_Attendance_Log.tblmaster_fingerprint_serverRow

        Try
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("From should not be greater than to.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            For Each items_ As Object In Me.ChkListServer.Items
                If Me.ChkListServer.GetItemCheckState(Me.ChkListServer.Items.IndexOf(items_)) = CheckState.Checked Then
                    SelectedCount_ += 1
                    RowSrv = Me.TblServer.Rows(Me.ChkListServer.Items.IndexOf(items_))
                    ListPrimaryActiveServer.Add(RowSrv.primary)
                End If
            Next
            If SelectedCount_ < 1 Then
                MsgBox("Please select at least 1 server.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvItem.Rows.Clear()
            DtaEmployee.Connection.ConnectionString = Core_Module.GetConstring
            DtaEmployee.Fill(TblEmployee)

            If ListPrimaryActiveServer.Count > 0 Then
                For Each PrimarySrv As Double In ListPrimaryActiveServer
                    ServerName = Me.ObjSrv.GetServerName(PrimarySrv)
                    ServerLocation = Me.ObjSrv.GetServerLocation(PrimarySrv)
                    ListActiveTerminalID.Clear()
                    ListActiveTerminalID = ObjSrv.GetListActiveTerminalID(PrimarySrv)
                    If ListActiveTerminalID.Count > 0 Then
                        ThereIsActiveTerminal = True
                    End If

                    If ObjSrv.GetMachineType(PrimarySrv) = ClsMasterAttendanceServer.EnumMachineType.NITGEN Then
                        ''test connection
                        Dim Connection_ As New SqlClient.SqlConnection
                        Connection_.ConnectionString = ObjSrv.GetAcmConnectionString(PrimarySrv)
                        Try
                            Connection_.Open()
                            If Connection_.State = ConnectionState.Open Then
                                SqlServerConnectionState_ = True
                                Connection_.Close()
                            End If
                        Catch ex0 As SqlClient.SqlException
                            SqlServerConnectionState_ = False
                            MsgBox("Failed show log from server : " & Me.ObjSrv.GetServerName(PrimarySrv), MsgBoxStyle.Information, "Information")
                        End Try

                        If SqlServerConnectionState_ = True Then
                            DtaLog.Connection.ConnectionString = ObjSrv.GetAcmConnectionString(PrimarySrv)
                            If DtaLog.FillBy_Date(TblLog, CDate(Format(Me.DtpFrom.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(Me.DtpTo.Value, "yyyy-MM-dd 23:59:00"))) > 0 Then
                                Me.ColLocation.Visible = True
                                Me.ColTerminalID.Visible = True
                                For Each RowLog In TblLog
                                    If ListActiveTerminalID.Contains(RowLog.TerminalID) = True Then
                                        Me.DgvItem.Rows.Add("")
                                        Me.RowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                                        With Me.RowDgv
                                            TblEmployee.DefaultView.RowFilter = "[fingerprint id] = '" & RowLog.UserID & "'"
                                            If TblEmployee.DefaultView.Count > 0 Then
                                                RowEmployee = TblEmployee.DefaultView(0).Row
                                                EmployeeCode = RowEmployee.Code
                                                EmployeeName = RowEmployee.Full_Name
                                                EmployeePrimary = RowEmployee.Primary
                                            Else
                                                EmployeeCode = ""
                                                EmployeeName = ""
                                                EmployeePrimary = -1
                                            End If
                                            .Cells(Me.ColNo.Index).Value = ""
                                            .Cells(Me.ColEmployeeCode.Index).Value = EmployeeCode
                                            .Cells(Me.ColEmployeeName.Index).Value = EmployeeName
                                            .Cells(Me.ColFingerprintID.Index).Value = RowLog.UserID
                                            .Cells(Me.ColServerName.Index).Value = ServerName
                                            .Cells(Me.ColLocation.Index).Value = ServerLocation
                                            .Cells(Me.ColTerminalID.Index).Value = RowLog.TerminalID
                                            .Cells(Me.ColTime.Index).Value = Format(RowLog.TransactionTime, "dd-MMM-yyyy HH:mm:ss")
                                            .Cells(Me.ColPrimaryEmployee.Index).Value = EmployeePrimary
                                            .Cells(Me.ColPrimaryServer.Index).Value = PrimarySrv
                                        End With
                                    End If
                                Next
                            End If
                        End If

                    Else
                        ''test connection
                        Dim Connection_ As New OleDb.OleDbConnection
                        Connection_.ConnectionString = ObjSrv.GetMagicConnectionString(PrimarySrv)
                        Try
                            Connection_.Open()
                            If Connection_.State = ConnectionState.Open Then
                                SqlServerConnectionState_ = True
                                Connection_.Close()
                            End If
                        Catch exAcc As OleDb.OleDbException
                            SqlServerConnectionState_ = False
                            MsgBox("Failed show log from server : " & Me.ObjSrv.GetServerName(PrimarySrv), MsgBoxStyle.Information, "Information")
                        End Try

                        If SqlServerConnectionState_ = True Then
                            DtaLogAcc.Connection.ConnectionString = ObjSrv.GetMagicConnectionString(PrimarySrv)
                            If DtaLogAcc.Fill(TblLogAcc, CDate(Format(Me.DtpFrom.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(Me.DtpTo.Value, "yyyy-MM-dd 23:59:59"))) > 0 Then
                                For Each RowLogAcc In TblLogAcc
                                    'If ListActiveTerminalID.Contains(RowLog.TerminalID) = True Then
                                    Me.DgvItem.Rows.Add("")
                                    Me.RowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                                    With Me.RowDgv
                                        TblEmployee.DefaultView.RowFilter = "[fingerprint id] = '" & RowLogAcc.USERID & "'"
                                        If TblEmployee.DefaultView.Count > 0 Then
                                            RowEmployee = TblEmployee.DefaultView(0).Row
                                            EmployeeCode = RowEmployee.Code
                                            EmployeeName = RowEmployee.Full_Name
                                            EmployeePrimary = RowEmployee.Primary
                                        Else
                                            EmployeeCode = ""
                                            EmployeeName = ""
                                            EmployeePrimary = -1
                                        End If
                                        .Cells(Me.ColNo.Index).Value = ""
                                        .Cells(Me.ColEmployeeCode.Index).Value = EmployeeCode
                                        .Cells(Me.ColEmployeeName.Index).Value = EmployeeName
                                        .Cells(Me.ColFingerprintID.Index).Value = RowLogAcc.USERID
                                        .Cells(Me.ColServerName.Index).Value = ServerName
                                        .Cells(Me.ColLocation.Index).Value = ServerLocation
                                        .Cells(Me.ColTerminalID.Index).Value = RowLogAcc.SENSORID
                                        .Cells(Me.ColTime.Index).Value = Format(RowLogAcc.CHECKTIME, "dd-MMM-yyyy HH:mm:ss")
                                        .Cells(Me.ColPrimaryEmployee.Index).Value = EmployeePrimary
                                        .Cells(Me.ColPrimaryServer.Index).Value = PrimarySrv
                                    End With
                                    'End If
                                Next
                            End If
                        End If
                    End If
                Next
                If ThereIsActiveTerminal = False Then
                    MsgBox("There is no active machine.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Else
                MsgBox("There is no active server settings.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.ObjCd.SetNumber(Me.DgvItem)
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            If Me.DgvItem.Rows.Count = 0 Then
                MsgBox("There is no log on this period.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Label3.Text = "Showing : log from attendance server."
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnImport.Click
        Dim DeletedCount As Integer = 0
        Dim InsertedCount As Integer = 0
        Dim MinDate As Date = Date.MaxValue
        Dim MaxDate As Date = Date.MinValue
        Dim MsgRst As MsgBoxResult
        Dim ListPrimarySrv As New List(Of Double)
        Try
            If Me.DgvItem.Rows.Count = 0 Or Me.Label3.Text = "Showing : log data from SMESS database." Then
                MsgBox("Show log from server first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            For Each RowDgv In Me.DgvItem.Rows
                If Format(CDate(RowDgv.Cells(Me.ColTime.Index).Value), "dd MMM yyyy") > MaxDate Then
                    MaxDate = CDate(RowDgv.Cells(Me.ColTime.Index).Value)
                End If
                If Format(CDate(RowDgv.Cells(Me.ColTime.Index).Value), "dd MMM yyyy") < MinDate Then
                    MinDate = CDate(RowDgv.Cells(Me.ColTime.Index).Value)
                End If
                If ListPrimarySrv.Contains(CDbl(RowDgv.Cells(Me.ColPrimaryServer.Index).Value)) = False Then
                    ListPrimarySrv.Add(CDbl(RowDgv.Cells(Me.ColPrimaryServer.Index).Value))
                End If
            Next
            If Me.ChkDeleteBeforeImport.Checked = True Then
                MsgRst = MsgBox("Are you sure to delete exist log from selected server and import new from these server?", MsgBoxStyle.YesNo, "Confirmation")
            Else
                MsgRst = MsgBox("Are you sure to import log from server?", MsgBoxStyle.YesNo, "Confirmation")
            End If
            If MsgRst = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                If Me.ChkDeleteBeforeImport.Checked = True Then
                    DeletedCount = Me.Obj.DeleteLog(CDate(Format(MinDate, "dd MMM yyyy 00:00:00")), CDate(Format(MaxDate, "dd MMM yyyy 23:59:00")), ListPrimarySrv)
                End If
                'InsertedCount = Me.Obj.InsertLog(Me, Me.DgvItem, Me.UserCode)
                InsertedCount = Me.Obj.InsertItemUsingScript(Me, Me.DgvItem, Me.UserCode)
                If Me.ChkDeleteBeforeImport.Checked = True Then
                    MsgBox(DeletedCount & " log deleted and " & InsertedCount & " log imported.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox(InsertedCount & " log imported.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        Dim Result As Integer = 0
        Dim MsgBoxResult_ As MsgBoxResult
        Dim ListPrimarySrv As New List(Of Double)
        Dim SelectedCount_ As Integer = 0
        Dim RowSrv As DseTrans_Attendance_Log.tblmaster_fingerprint_serverRow
        Try
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("From should not be greater than to.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each items_ As Object In Me.ChkListServer.Items
                If Me.ChkListServer.GetItemCheckState(Me.ChkListServer.Items.IndexOf(items_)) = CheckState.Checked Then
                    SelectedCount_ += 1
                    RowSrv = Me.TblServer.Rows(Me.ChkListServer.Items.IndexOf(items_))
                    ListPrimarySrv.Add(RowSrv.primary)
                End If
            Next
            If SelectedCount_ < 1 Then
                MsgBox("Please select at least 1 server.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.DtpFrom.Value.Date = Me.DtpTo.Value.Date Then
                MsgBoxResult_ = MsgBox("Are you sure to delete log at " & Format(Me.DtpFrom.Value, "dd MMM yyyy") & " from selected server?", MsgBoxStyle.YesNo, "Confirmation")
            Else
                MsgBoxResult_ = MsgBox("Are you sure to delete log from " & Format(Me.DtpFrom.Value, "dd MMM yyyy") & " to " & Format(Me.DtpTo.Value, "dd MMM yyyy") & " from selected server?", MsgBoxStyle.YesNo, "Confirmation")
            End If
            If MsgBoxResult_ = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                Result = Me.Obj.DeleteLog(CDate(Format(Me.DtpFrom.Value, "dd MMM yyyy 00:00:00")), CDate(Format(Me.DtpTo.Value, "dd MMM yyyy 23:59:00")), ListPrimarySrv)
                If Result > 0 Then
                    Me.BtnShowLog.PerformClick()
                    MsgBox(Result & " log deleted.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("No log deleted.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub BtnShowLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShowLog.Click
        Dim DtaLogMySQL As New DseTrans_Attendance_LogTableAdapters.DataTable1TableAdapter
        Dim TblMySQL As New DseTrans_Attendance_Log.DataTable1DataTable
        Dim RowLogMySQL As DseTrans_Attendance_Log.DataTable1Row
        Dim ListPrimaryServer As New List(Of Double)
        Dim NamaServer As String = ""
        Dim SelectedCount_ As Integer = 0
        Dim RowSrv As DseTrans_Attendance_Log.tblmaster_fingerprint_serverRow

        Try
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("From date should not be greater than To date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvItem.Rows.Clear()

            DtaLogMySQL.Connection.ConnectionString = Core_Module.GetConstring
            DtaLogMySQL.FillBy_Date(TblMySQL, CDate(Format(Me.DtpFrom.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(Me.DtpTo.Value, "yyyy-MM-dd 23:59:59")))

            'buang baris yang tidak dicek list servernya
            For Each items_ As Object In Me.ChkListServer.Items
                If Me.ChkListServer.GetItemCheckState(Me.ChkListServer.Items.IndexOf(items_)) = CheckState.Checked Then
                    SelectedCount_ += 1
                    RowSrv = Me.TblServer.Rows(Me.ChkListServer.Items.IndexOf(items_))
                    ListPrimaryServer.Add(RowSrv.primary)
                End If
            Next
            For ij As Double = TblMySQL.Rows.Count - 1 To 0 Step -1
                RowLogMySQL = TblMySQL.Rows(ij)
                If ListPrimaryServer.Contains(RowLogMySQL.primary_attendance_server) = False Then
                    TblMySQL.Rows.Remove(RowLogMySQL)
                End If
            Next

            If TblMySQL.Count > 0 Then
                For Each RowLogMySQL In TblMySQL
                    Me.DgvItem.Rows.Add("")
                    Me.RowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                    With Me.RowDgv
                        If RowLogMySQL.server_name <> "" Or RowLogMySQL.primary_attendance_server = -2 Then
                            If RowLogMySQL.primary_attendance_server = -2 Then
                                RowLogMySQL.server_name = "From File"
                            End If
                            .Cells(Me.ColNo.Index).Value = ""
                            .Cells(Me.ColEmployeeCode.Index).Value = RowLogMySQL.Code
                            .Cells(Me.ColEmployeeName.Index).Value = RowLogMySQL.Full_Name
                            .Cells(Me.ColFingerprintID.Index).Value = RowLogMySQL.fingerprint_id
                            .Cells(Me.ColServerName.Index).Value = RowLogMySQL.server_name
                            .Cells(Me.ColLocation.Index).Value = RowLogMySQL.location
                            .Cells(Me.ColTerminalID.Index).Value = RowLogMySQL.terminal_id
                            .Cells(Me.ColTime.Index).Value = Format(RowLogMySQL.transaction_time, "dd-MMM-yyyy HH:mm:ss")
                            .Cells(Me.ColPrimaryEmployee.Index).Value = RowLogMySQL.primary_employee
                            .Cells(Me.ColPrimaryServer.Index).Value = RowLogMySQL.primary_attendance_server
                        End If
                    End With
                Next
            Else
                MsgBox("There is no log on this period.", MsgBoxStyle.Information, "Information")
            End If

            Me.ObjCd.SetNumber(Me.DgvItem)
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Label3.Text = "Showing : log data from SMESS database."
            Me.Cursor = Cursors.Default
            ListPrimaryServer = Nothing
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnIndexSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnIndexSettings.Click
        Dim FrmTrans_Attendance_Log_IndexSettings As New FrmTrans_Attendance_Log_IndexSettings
        Try
            FrmTrans_Attendance_Log_IndexSettings.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub

    Private Sub BtnShowFromFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShowFromFile.Click
        Dim FileName As String = String.Empty
        Dim DR As DataGridViewRow
        Dim Att As New ClsAttendanceFileParser
        Dim AttList As New List(Of ClsAttendanceFileParser)
        Dim dtaEmployee As New Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim tblEmployee As New Core_Module.DseMasterKaryawan.tblmaster_karyawanDataTable
        Dim rowEmployee As Core_Module.DseMasterKaryawan.tblmaster_karyawanRow
        Try
            If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK And System.IO.File.Exists(Me.OpenFileDialog1.FileName) Then
                FileName = Me.OpenFileDialog1.FileName
                'Dim infoReader As System.IO.FileInfo
                'infoReader = My.Computer.FileSystem.GetFileInfo(FileName)
            Else
                Exit Sub
            End If
            If System.IO.File.Exists(FileName) = False Then
                MsgBox("File doesn't exist.", MsgBoxStyle.Information, "Info")
                Exit Sub
            End If
            Att.SetIndexAndLength(My.Settings.YEAR, My.Settings.YEAR_LENGTH, My.Settings.MONTH, My.Settings.MONTH_LENGTH, My.Settings.DATE_, My.Settings.DATE_LENGTH, My.Settings.HOUR, My.Settings.HOUR_LENGTH, My.Settings.MINUTE, My.Settings.MINUTE_LENGTH, My.Settings.USER_ID, My.Settings.USER_ID_LENGTH)
            Me.Cursor = Cursors.WaitCursor
            AttList = Att.GetLog(FileName, My.Settings.REMOVE_ZERO_BEFORE_UID)
            If AttList.Count > 0 Then
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvItem.Rows.Clear()
                Me.ColLocation.Visible = False
                Me.ColTerminalID.Visible = False
                dtaEmployee.Connection.ConnectionString = Core_Module.GetConstring
                dtaEmployee.Fill(tblEmployee)
                For Each Att In AttList
                    DR = Me.DgvItem.Rows(Me.DgvItem.Rows.Add())
                    With DR
                        .Cells(Me.ColNo.Index).Value = Me.DgvItem.Rows.Count
                        .Cells(Me.ColFingerprintID.Index).Value = Att.UserID
                        .Cells(Me.ColTime.Index).Value = Format(Att.LogTime, "dd-MMM-yyyy HH:mm:ss")
                        tblEmployee.DefaultView.RowFilter = "[fingerprint id] = '" & Att.UserID & "'"
                        If tblEmployee.DefaultView.Count > 0 Then
                            rowEmployee = tblEmployee.DefaultView(0).Row
                            .Cells(Me.ColEmployeeCode.Index).Value = rowEmployee.Code
                            .Cells(Me.ColEmployeeName.Index).Value = rowEmployee.Full_Name
                            .Cells(Me.ColPrimaryEmployee.Index).Value = rowEmployee.Primary
                        Else
                            .Cells(Me.ColPrimaryEmployee.Index).Value = -1
                        End If
                        .Cells(Me.ColServerName.Index).Value = "From File"
                        .Cells(Me.ColLocation.Index).Value = "-"
                        .Cells(Me.ColTerminalID.Index).Value = "-2"
                        .Cells(Me.ColPrimaryServer.Index).Value = "-2"
                    End With
                Next
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            Me.Label3.Text = "Showing : log from file."
        End Try
    End Sub
End Class