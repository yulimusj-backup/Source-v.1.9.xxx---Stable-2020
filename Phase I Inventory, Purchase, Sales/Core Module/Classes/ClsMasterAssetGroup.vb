Option Explicit On
Option Strict Off

Public Class ClsMasterAssetGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mGroupCode As String = ""
    Dim mDescription As String = ""
    Dim ObjCekdata As New ClsCekData
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property GroupCode() As String
        Get
            Return Me.mGroupCode
        End Get
        Set(ByVal value As String)
            Me.mGroupCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDescription
        End Get
        Set(ByVal value As String)
            Me.mDescription = value
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

    Public Overridable Function ApakahDataBenar() As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.GetJmlPrimaryByCode(Me.GroupCode) < 0 Then
                'salah
                Return False
            End If

            If Me.mGroupCode = "" Then
                'salah
                Return False
            End If

            If Me.mDescription = "" Then
                'salah
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim ObjAset As New ClsMasterAsset
        Try
            If ObjAset.GetJmlByGroup(vPrimary) > 0 Then
                MsgBox("This asset group data can't delete because this data is refering to master asset", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAset = Nothing
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.GroupCode, Me.mDescription, Now, vUserCode, Now, vUserCode, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mGroupCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(CDec(Me.mPrimary)).Value
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

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mGroupCode, Me.mDescription, Now, vUserCode, Now, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestampByPrimary(CDec(Me.mPrimary)).Value
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

    Public Function ApakahSudahAda(ByVal vGroupCode As String) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlPrimaryByCode(vGroupCode).HasValue Then
                If DTA.ScalarQuery_GetJmlPrimaryByCode(vGroupCode).Value > 0 Then
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

    Public Function GetJmlPrimaryByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlPrimaryByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlPrimaryByCode(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetGroupCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCodeByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCodeByPrimary(vPrimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestampByPrimary(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vGroupCode As String) As Double
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByCode(vGroupCode).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByCode(vGroupCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Dim TBL As New DseMasterAssetGroup.tblmaster_asset_groupDataTable
        Dim row As DseMasterAssetGroup.tblmaster_asset_groupRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterAssetGroup.tblmaster_asset_groupRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_date
                    Me.mDescription = .Description
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mGroupCode = .Group_Code
                    Me.mTimestamp = .Timestamp
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

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterAssetGroup.tblmaster_asset_groupDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter
        Dim row As DseMasterAssetGroup.tblmaster_asset_groupRow
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_LikeCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_LikeDesc(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataBindings.Clear()
            DGV.DataSource = Nothing
            DGV.Rows.Clear()
            For Each row In TBL.Rows
                DGV.Rows.Add(row.Primary, row.Group_Code, row.Description)
            Next
            'If DGV.ColumnCount > 0 Then
            '    DGV.Columns(0).Visible = False
            '    DGV.Columns(1).HeaderText = "Group code"
            '    DGV.Columns(0).Frozen = True
            '    DGV.Columns(1).Frozen = True
            'End If


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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
      EnumSearchCriteria) As Double
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlLikeCode(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeCode(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlLikeDesc(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeDesc(vKriteria + "%").Value
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

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterAssetGroupTableAdapters.tblmaster_asset_groupTableAdapter

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
End Class
