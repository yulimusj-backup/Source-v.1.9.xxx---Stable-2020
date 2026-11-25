Option Explicit On
Option Strict Off

'Imports MySql.Data.MySqlClient
'Imports MySql.Data

Public Class ClsMasterBackUpDB
    Dim ObjCD As New Core_Module.ClsCekData

    Public WithEvents BackUpScheduler As New Timer
    Public RefTimeScheduler As Date = Now
    'Dim WithEvents BW As New System.ComponentModel.BackgroundWorker

    Dim mDummyUserID As String = ""
    Dim mDefaultInterval As Integer = 1800000

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDescription As String = ""
    Dim mLocation1 As String = ""
    Dim mLocation2 As String = ""
    Dim mInterval As EnumInterval = EnumInterval.Daily
    Dim mDayOfWeek As DayOfWeek = DayOfWeek.Sunday
    Dim mBackUpTime As Double = 0
    Dim mPrimaryUser As Double = -1
    Dim mShutdownAfter As Integer = 0
    Dim mRemindOnClose As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumInterval As Integer
        Daily
        Weekly
    End Enum

    Public Property RemindOnClose() As Integer
        Get
            Return Me.mRemindOnClose
        End Get
        Set(ByVal value As Integer)
            mRemindOnClose = value
        End Set
    End Property

    Public Property ShutdownAfter() As Integer
        Get
            Return Me.mShutdownAfter
        End Get
        Set(ByVal value As Integer)
            mShutdownAfter = value
        End Set
    End Property

    Public Property PrimaryUser() As Double
        Get
            Return Me.mPrimaryUser
        End Get
        Set(ByVal value As Double)
            mPrimaryUser = value
        End Set
    End Property

    Public Property BackUpTime() As Double
        Get
            Return Me.mBackUpTime
        End Get
        Set(ByVal value As Double)
            mBackUpTime = value
        End Set
    End Property

    Public Property DayOfWeek() As DayOfWeek
        Get
            Return Me.mDayOfWeek
        End Get
        Set(ByVal value As DayOfWeek)
            mDayOfWeek = value
        End Set
    End Property

    Public Property Interval() As EnumInterval
        Get
            Return Me.mInterval
        End Get
        Set(ByVal value As EnumInterval)
            mInterval = value
        End Set
    End Property

    Public Property Location1() As String
        Get
            Return Me.mLocation1
        End Get
        Set(ByVal value As String)
            mLocation1 = value
        End Set
    End Property

    Public Property Location2() As String
        Get
            Return Me.mLocation2
        End Get
        Set(ByVal value As String)
            mLocation2 = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDescription, Me.mLocation1, Me.mLocation2, Me.mInterval, Me.mDayOfWeek, _
            Me.mBackUpTime, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryUser, Me.mShutdownAfter, Me.mRemindOnClose) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDescription, Me.mLocation1, Me.mLocation2, Me.mInterval, Me.mDayOfWeek, Me.mBackUpTime, _
            vUserCode, Now, Now, Me.mPrimaryUser, Me.mShutdownAfter, Me.mRemindOnClose, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(CDec(vPrimary))
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vCode).Value > 0 Then
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
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahRemindOnClose(ByVal vUserID As String) As Boolean
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryUser = ObjUser.GetPrimary(vUserID)
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If PrimaryUser = -1 Then
                Return False
            Else
                Dummy = DTA.ScalarQuery_GetJml2(PrimaryUser)
                If Dummy.HasValue = False Then
                    Return False
                Else
                    If Dummy.Value > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterBackUpDB.tblmaster_backup_db1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_db1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
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
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter
        Dim TBL As New DseMasterBackUpDB.tblmaster_backup_dbDataTable
        Dim row As DseMasterBackUpDB.tblmaster_backup_dbRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterBackUpDB.tblmaster_backup_dbRow)
                With row
                    Me.mPrimary = row.primary
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mDescription = .desc
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date

                    Me.mCode = .code
                    Me.mTimestamp = .timestamp
                    Me.mLocation1 = .location1
                    Me.mLocation2 = .location2
                    Me.mBackUpTime = .back_up_time
                    Me.mDayOfWeek = .day_of_week
                    Me.mInterval = .interval
                    Me.mPrimaryUser = .primary_user
                    Me.mShutdownAfter = .shut_down_after
                    Me.mRemindOnClose = .back_up_on_close
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetCode(CDec(vPrimary))
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub StartBackUpScheduler(ByVal vUserID As String)
        'Dim strConn As String = Core_Module.GetConstring
        'Dim connection As New MySqlConnection(strConn)
        'connection.Open()
        Dim ObjUser As New Core_Module.ClsUser
        Dim vPrimaryUserID As Double = -1
        Dim DTA As New DseMasterBackUpDBTableAdapters.tblmaster_backup_dbTableAdapter
        Dim TBL As New DseMasterBackUpDB.tblmaster_backup_dbDataTable
        Dim row As DseMasterBackUpDB.tblmaster_backup_dbRow
        Dim NextBackUpDate As Date = Date.MaxValue
        Dim NextBackUpDate_Dummy As Date = Now
        Dim vIntervalDummy As EnumInterval = EnumInterval.Daily
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TS As TimeSpan
        Dim vPrimaryReference As Double = -1

        Try
            Me.BackUpScheduler.Enabled = False
            'Me.BackUpScheduler.Interval = 2000
            'Me.BackUpScheduler.Enabled = True
            'Exit Sub

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mDummyUserID = vUserID

            vPrimaryUserID = ObjUser.GetPrimary(vUserID)
            If vPrimaryUserID <> -1 Then
                DTA.FillBy_PrimaryUser(TBL, vPrimaryUserID)
                If TBL.Rows.Count > 0 Then
                    'cari next back up schedule
                    For Each row In TBL.Rows
                        vIntervalDummy = row.interval
                        If vIntervalDummy = EnumInterval.Daily Then
                            NextBackUpDate_Dummy = CDate(Format(Today.Date, "MM/dd/yyyy") + " " + ObjCD.Convert_ToString(row.back_up_time).Replace(".", ":"))
                            If NextBackUpDate_Dummy < NextBackUpDate And NextBackUpDate_Dummy > Now Then
                                NextBackUpDate = NextBackUpDate_Dummy
                                vPrimaryReference = row.primary
                            End If
                        Else
                            If Today.DayOfWeek = row.day_of_week Then
                                NextBackUpDate_Dummy = CDate(Format(Today.Date, "MM/dd/yyyy") + " " + ObjCD.Convert_ToString(row.back_up_time).Replace(".", ":"))
                                If NextBackUpDate_Dummy < NextBackUpDate And NextBackUpDate_Dummy > Now Then
                                    NextBackUpDate = NextBackUpDate_Dummy
                                    vPrimaryReference = row.primary
                                End If
                            End If
                        End If
                    Next

                    'salin data sesuai next back up schedule
                    If vPrimaryReference = -1 Or NextBackUpDate = Date.MaxValue Then
                        Me.BackUpScheduler.Interval = Me.mDefaultInterval
                        Me.RefTimeScheduler = Now
                    Else
                        'set interval 
                        If NextBackUpDate < Now Then
                            Me.BackUpScheduler.Interval = Me.mDefaultInterval
                            Me.RefTimeScheduler = Now
                        Else
                            TS = NextBackUpDate - Now
                            If TS.TotalMilliseconds > Me.mDefaultInterval Then
                                Me.BackUpScheduler.Interval = Me.mDefaultInterval
                                Me.RefTimeScheduler = Now
                            Else
                                Me.BackUpScheduler.Interval = (NextBackUpDate - Now).TotalMilliseconds
                                Me.RefTimeScheduler = Now
                                'If TS.TotalMinutes > 1 Then
                                '    Me.BackUpScheduler.Interval = TS.TotalMinutes * 60000
                                '    Me.RefTimeScheduler = Now
                                'Else
                                '    Me.BackUpScheduler.Interval = TS.TotalMilliseconds
                                '    Me.RefTimeScheduler = Now
                                'End If
                            End If
                        End If
                    End If

                    If vPrimaryReference <> -1 Then
                        For Each row In TBL.Rows
                            If row.primary = vPrimaryReference Then
                                With row
                                    Me.mPrimary = row.primary
                                    Me.mCreatedBy = .created_by
                                    Me.mCreatedDate = .created_date
                                    Me.mDescription = .desc
                                    Me.mModifiedBy = .modified_by
                                    Me.mModifiedDate = .modified_date

                                    Me.mCode = .code
                                    Me.mTimestamp = .timestamp
                                    Me.mLocation1 = .location1
                                    Me.mLocation2 = .location2
                                    Me.mBackUpTime = .back_up_time
                                    Me.mDayOfWeek = .day_of_week
                                    Me.mInterval = .interval
                                    Me.mPrimaryUser = .primary_user
                                    Me.mShutdownAfter = .shut_down_after
                                    Me.mRemindOnClose = .back_up_on_close
                                End With
                                Exit For
                            End If
                        Next
                    End If

                    'start Timer
                    If Me.BackUpScheduler.Enabled = False Then
                        Me.BackUpScheduler.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
            'start Timer
            If Me.BackUpScheduler.Enabled = False Then
                Me.BackUpScheduler.Enabled = True
                Me.RefTimeScheduler = Now
            End If
        End Try
    End Sub

    Private Sub BackUpScheduler_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackUpScheduler.Tick
        Try
            'Me.BackUpScheduler.Enabled = False
            'Me.BW.RunWorkerAsync()

            If Me.BackUpScheduler.Interval = Me.mDefaultInterval Then
                Me.BackUpScheduler.Enabled = False
                Me.StartBackUpScheduler(Me.mDummyUserID)
            Else
                Me.BackUpScheduler.Enabled = False
                Me.BackUp()
                'Threading.Thread.Sleep(121000)

                'If Me.BW.IsBusy = False Then
                '    Me.BackUp()
                '    'Me.BW.RunWorkerAsync()
                '    'wait 1 minute
                '    Threading.Thread.Sleep(61000)
                'Else
                '    'wait 1 minutes while auto back up running
                '    'While Me.BW.IsBusy = True
                '    '    Threading.Thread.Sleep(61000)
                '    'End While
                '    'Me.BW.RunWorkerAsync()
                '    Me.BackUp()
                '    'wait 1 minute
                '    Threading.Thread.Sleep(61000)
                'End If

                'aim for next back up, setelah suspend 61 detik + waktu selama lagi back up, supaya tidak ada back up sekarang yang menggunakan nama back up selanjutnya
                'Me.StartBackUpScheduler(Me.mDummyUserID)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New()
        Try
            Me.BackUpScheduler.Enabled = False
            Me.BackUpScheduler.Interval = Me.mDefaultInterval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub BW_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BW.DoWork
    '    Me.BackUp()
    'End Sub

    Private Sub BackUp()
        Try
            Dim ObjBackUp As New Core_Module.ClsMySQL
            Dim Shutdown_Bool As Boolean

            If Me.mShutdownAfter = 1 Then
                Shutdown_Bool = True
            Else
                Shutdown_Bool = False
            End If
            If Me.mLocation2 <> "" Then
                ObjBackUp.BackUpDataBase(Core_Module.GetConstring, Me.mLocation1 + "\" + Me.mDescription + ".smess", Core_Module.GetDBServer, _
                Core_Module.GetDBName, Core_Module.GetDBUserID, Core_Module.GetDBPassword, Core_Module.GetDBPort, Me.mLocation2 + "\" + Me.mDescription + ".smess" _
                , Me.mDescription, True, Shutdown_Bool, True)
            Else
                ObjBackUp.BackUpDataBase(Core_Module.GetConstring, Me.mLocation1 + "\" + Me.mDescription + ".smess", Core_Module.GetDBServer, _
                Core_Module.GetDBName, Core_Module.GetDBUserID, Core_Module.GetDBPassword, Core_Module.GetDBPort, "" _
                , Me.mDescription, True, Shutdown_Bool)
            End If

            While ObjBackUp.OnProgressBackUp = True
                Threading.Thread.Sleep(60000 * 1)
            End While
            Me.StartBackUpScheduler(Me.mDummyUserID)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
