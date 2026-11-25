Option Explicit On
Option Strict Off

Imports system.data
Imports system.data.sqlclient

Public Class ClsMasterMachineData
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mMachineDataCode As String = ""
    Dim mDescription As String = ""
    Dim mPrimaryMachineGroup As Double = -1
    Dim mRemarks As String = ""
    Dim mActive As Integer = enumActive.NonActive
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum enumActive As Integer
        NonActive
        Active
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property MachineDataCode() As String
        Get
            Return Me.mMachineDataCode
        End Get
        Set(ByVal value As String)
            Me.mMachineDataCode = value
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

    Public Property PrimaryMachineGroup() As Double
        Get
            Return mPrimaryMachineGroup
        End Get
        Set(ByVal value As Double)
            mPrimaryMachineGroup = value
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

    Public Property Active() As String
        Get
            Return Me.mActive
        End Get
        Set(ByVal value As String)
            mActive = value
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mMachineDataCode, Me.mDescription, Me.mPrimaryMachineGroup, Me.Remarks, Me.mActive, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByCode(Me.mMachineDataCode).Value
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
       Core_Module.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                RetVal = DTA.ScalarQuery_GetJmlByCode(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDescription Then
                RetVal = DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByMachineGroup Then
                RetVal = DTA.ScalarQuery_GetJmlByMachineGroup(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mMachineDataCode, Me.mDescription, Me.mPrimaryMachineGroup, Me.mRemarks, Me.mActive, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

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

    Public Function ApakahSudahAdaOnlyActive(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlOnlyActive(vCode).HasValue Then
                If DTA.ScalarQuery_GetJmlOnlyActive(vCode).Value > 0 Then
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryMachine(vPrimary).HasValue Then
                If DTA.ScalarQuery_GetJmlByPrimaryMachine(vPrimary).Value > 0 Then
                    MsgBox("You cannot delete this data, because there is work unit using this data.", _
                   MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterMachineData.tblmaster_machinedata1DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal status As String) As Boolean
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedata1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If status = "Show All" Then
                If SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_CodeLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByMachineGroup Then
                    QueryResult = DTA.FillBy_MachineGroup(TBL, Criteria + "%")
                End If
            ElseIf status = "Only Active" Then
                If SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_CodeLikeActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescLikeActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLikeActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByMachineGroup Then
                    QueryResult = DTA.FillBy_MachineGroup_Active(TBL, Criteria + "%")
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView

            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(1).HeaderText = "Code"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + _
                    " miliseconds"
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter
        Dim TBL As New DseMasterMachineData.tblmaster_machinedataDataTable
        Dim row As DseMasterMachineData.tblmaster_machinedataRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterMachineData.tblmaster_machinedataRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.mMachineDataCode = .Machine_Data_Code
                    Me.mDescription = .Description
                    Me.mPrimaryMachineGroup = .Primary_Machine_Group
                    Me.mRemarks = .Remarks
                    Me.mActive = .Active
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
            TBL.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetCodeByPrimary(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDescByPrimary(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vCode As String) As Double
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJml(vCode)) = False Then
                Return DTA.ScalarQuery_GetJml(vCode)
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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

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
        Dim DTA As New DseMasterMachineDataTableAdapters.tblmaster_machinedataTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByCode(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
