Option Explicit On
Option Strict Off

Public Class ClsMasterSetupTime
    Dim ObjCD As New Core_Module.ClsCekData
    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDescription As String = ""
    Dim mDefaultTime As Decimal = 0
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now


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

    Public Property DefaultTime() As Decimal
        Get
            Return Me.mDefaultTime
        End Get
        Set(ByVal value As Decimal)
            mDefaultTime = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
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

        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.InsertQuery(Me.mCode, Me.mDescription, _
            Me.mDefaultTime, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mCode)
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
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

    Public Function CariDetail(ByVal vprimary As Double) As Boolean
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter
        Dim tbl As New DseMasterSetupTime.tblmaster_setuptimeDataTable
        Dim row As DseMasterSetupTime.tblmaster_setuptimeRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(tbl, vprimary)
            If tbl.Rows.Count > 0 Then
                row = tbl.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDescription = .Description
                    Me.mRemarks = .Remarks
                    Me.mCode = .Code
                    Me.mDefaultTime = .Default_Time
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            tbl.Dispose()
        End Try

    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterSetupTime.tblmaster_setuptime_cariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptime_cariTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim row As DseMasterSetupTime.tblmaster_setuptime_cariRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_Code(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_Desc(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_Remarks(TBL, Criteria + "%")
            End If

            For Each row In TBL.Rows
                With row
                    .Default_Time_String = Me.ObjCD.Convert_ToString(.Default_Time)
                End With
            Next


            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()

            DGV.Columns(0).Visible = False
            DGV.Columns(3).Visible = False

            If DGV.Columns.Count > 0 Then
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If

            DGV.Columns(DGV.ColumnCount - 1).HeaderText = "Default Time"
            DGV.Columns(DGV.ColumnCount - 1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DGV.Columns(DGV.ColumnCount - 1).DisplayIndex = 3

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


    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue AndAlso DTA.ScalarQuery_GetPrimary(vCode).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try

    End Function

    Public Function GetDuration(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDuration(vPrimary).HasValue AndAlso DTA.ScalarQuery_GetDuration(vPrimary).Value > 0 Then
                Return DTA.ScalarQuery_GetDuration(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Not String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) Then
                Return DTA.ScalarQuery_GetCode(vPrimary).ToString
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescription(vPrimary)) Then
                Return DTA.ScalarQuery_GetDescription(vPrimary).ToString
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
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function



    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.UpdateQuery(Me.mCode, Me.mDescription, _
            Me.mDefaultTime, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary) > 0 Then
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
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
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.DeleteQuery(vPrimary)
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
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterSetupTimeTableAdapters.Work_UnitTableAdapter

        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimary).HasValue AndAlso DTA.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                Return False
            End If




            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMasterSetupTimeTableAdapters.tblmaster_setuptimeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlByCode(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByCode(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%").Value
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

End Class
