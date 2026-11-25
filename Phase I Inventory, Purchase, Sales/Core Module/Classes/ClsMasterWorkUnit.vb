Option Explicit On
Option Strict Off

Public Class ClsMasterWorkUnit
    Dim mPrimary As Double = -1
    Dim mPrimaryWorkUnitGroup As Double = -1
    Dim mWorkUnitCode As String = ""
    Dim mDesc As String = ""
    Dim mRemarks As String = ""
    Dim mCodeWUGroup As String = ""
    Dim mDescWUGroup As String = ""
    Dim mCodeLocIN As String = ""
    Dim mCodeLocOUT As String = ""
    Dim mDescLocIN As String = ""
    Dim mDescLocOUT As String = ""
    Dim mActive As Integer = enumActive.NonActive
    Dim mPrimaryLocationIN As Double = -1
    Dim mPrimaryLocationOUT As Double = -1
    Dim mPrimaryProductionLine As Double

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum enumActive As Integer
        NonActive
        Active
    End Enum

#Region "Property"
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryWorkUnitGroup() As Double
        Get
            Return mPrimaryWorkUnitGroup
        End Get
        Set(ByVal value As Double)
            mPrimaryWorkUnitGroup = value
        End Set
    End Property

    Public Property WorkUnitCode() As String
        Get
            Return mWorkUnitCode
        End Get
        Set(ByVal value As String)
            mWorkUnitCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property Active() As Integer
        Get
            Return Me.mActive
        End Get
        Set(ByVal value As Integer)
            mActive = value
        End Set
    End Property

    Public Property PrimaryProductionLine() As Double
        Get
            Return Me.mPrimaryProductionLine
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryProductionLine = value
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

    Public Property PrimaryLocationIN() As Double
        Get
            Return mPrimaryLocationIN
        End Get
        Set(ByVal value As Double)
            mPrimaryLocationIN = value
        End Set
    End Property

    Public Property PrimaryLocationOUT() As Double
        Get
            Return mPrimaryLocationOUT
        End Get
        Set(ByVal value As Double)
            mPrimaryLocationOUT = value
        End Set
    End Property

    Public Property CodeWUGroup() As String
        Get
            Return mCodeWUGroup
        End Get
        Set(ByVal value As String)
            mCodeWUGroup = value
        End Set
    End Property

    Public Property DescWUGroup() As String
        Get
            Return mDescWUGroup
        End Get
        Set(ByVal value As String)
            mDescWUGroup = value
        End Set
    End Property

    Public Property CodeLocIN() As String
        Get
            Return mCodeLocIN
        End Get
        Set(ByVal value As String)
            mCodeLocIN = value
        End Set
    End Property

    Public Property CodeLocOUT() As String
        Get
            Return mCodeLocOUT
        End Get
        Set(ByVal value As String)
            mCodeLocOUT = value
        End Set
    End Property

    Public Property DescLocOUT() As String
        Get
            Return mDescLocOUT
        End Get
        Set(ByVal value As String)
            mDescLocOUT = value
        End Set
    End Property

    Public Property DescLocIN() As String
        Get
            Return mDescLocIN
        End Get
        Set(ByVal value As String)
            mDescLocIN = value
        End Set
    End Property
#End Region

    'Public Function GetPrimaryStandByLocation(ByVal vPrimary As Double) As Double
    '    Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryStandByLocatio(vPrimary).HasValue Then
    '            Return DTA.ScalarQuery_GetPrimaryStandByLocation(vPrimary).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetPrimaryPickUpLocation(ByVal vPrimary As Double) As Double
    '    Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryPickUpLocation(vPrimary).HasValue Then
    '            Return DTA.ScalarQuery_GetPrimaryPickUpLocation(vPrimary).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCodeByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCodeByPrimary(vPrimary)
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
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDescByPrimary(vPrimary)
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJml(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
    '    Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetJml(vCode).HasValue Then
    '            Return DTA.ScalarQuery_GetJml(vCode).Value
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function ApakahSudahAdaOnlyActive(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByCodeOnlyActive(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCodeOnlyActive(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String, Optional ByVal vActiveOnly As Boolean = False) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vActiveOnly = False Then
                n = DTA.ScalarQuery_GetJml(vCode)
            Else
                n = DTA.ScalarQuery_GetJmlBy_ActiveOnly(vCode)
            End If

            If n.HasValue Then
                Return n.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItemEmployee As DataGridView, ByVal vDGVItemMachine As DataGridView, _
    ByVal i As FrmMasterWorkUnit) As Boolean
        Dim TmWorkUnitCode As String = ""
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mWorkUnitCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Duplication in work unit code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mWorkUnitCode Then
                    If Me.ApakahSudahAda(Me.mWorkUnitCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Duplication in work unit code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mWorkUnitCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please input work unit code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. nama
            If Me.mDesc = "" Then
                If vShowMsg = True Then
                    MsgBox("Please input work unit name.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            ''Item tidak boleh kosong 
            'If i.DgvEmployee.Rows.Count <= 1 AndAlso i.DgvMachine.Rows.Count <= 1 AndAlso i.dgvSetuptime.Rows.Count <= 1 Then
            '    MsgBox("Please input at lease 1 item data on employee or machine or setuptime ", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If
            'If i.DgvEmployee.Rows.Count <= 1 Then
            '    MsgBox("Please input item data employee", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If
            'If i.DgvMachine.Rows.Count <= 1 Then
            '    MsgBox("Please input item data machine", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If



            'If i.DgvEmployee.Rows.Count <= 0 Then
            '    MsgBox("Please input item data employee", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If
            'If i.DgvMachine.Rows.Count <= 0 Then
            '    MsgBox("Please input item data machine", MsgBoxStyle.Information, "Information ")
            '    Return False
            'End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.apakahsudahada_diallocatemo(vCode) = True Then
                MsgBox("The following work unit data is reference by one or several allocate mo, you can’t delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If DTA.ScalarQuery_GetJmlByPrimaryWorkUnit(vPrimary).HasValue Then
                If DTA.ScalarQuery_GetJmlByPrimaryWorkUnit(vPrimary).Value > 0 Then
                    MsgBox("You cannot delete this data, because there are product using this data.", _
                                                          MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterWorkUnit.tblmaster_workunit1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal status As String) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunit1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If status = "Show All" Then
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_LikeCode(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_LikeDesc(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_LikeRemarks(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByEmployee Then
                    QueryResult = DTA.FillBy_LikeEmployee(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByWorkShift Then
                    QueryResult = DTA.FillBy_LikeWorkshift(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByMachine Then
                    QueryResult = DTA.FillBy_LikeMachine(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByGroupCode Then
                    QueryResult = DTA.FillBy_GroupCodeLIKE(TBL, Criteria + "%", 1, 0)
                ElseIf SearchCriteria = EnumSearchCriteria.ByGroupDesc Then
                    QueryResult = DTA.FillBy_GroupDescLIKE(TBL, Criteria + "%", 1, 0)
                ElseIf SearchCriteria = EnumSearchCriteria.ByCodeLocationIn Then
                    QueryResult = DTA.FillBy_LikeCodeLocationIN(TBL, 1, 0, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByCodeLocationOut Then
                    QueryResult = DTA.FillBy_LikeCodeLocationOut(TBL, 1, 0, Criteria + "%")
                End If
            ElseIf status = "Only Active" Then
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_LikeCodeActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_LikeDescActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_LikeRemarksActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByWorkShift Then
                    QueryResult = DTA.FillBy_LikeWorkshiftActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByMachine Then
                    QueryResult = DTA.FillBy_LikeMachineActive(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByGroupCode Then
                    QueryResult = DTA.FillBy_GroupCodeLIKE(TBL, Criteria + "%", 1, 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByGroupDesc Then
                    QueryResult = DTA.FillBy_GroupDescLIKE(TBL, Criteria + "%", 1, 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByCodeLocationIn Then
                    QueryResult = DTA.FillBy_LikeCodeLocationIN(TBL, 1, 1, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByCodeLocationOut Then
                    QueryResult = DTA.FillBy_LikeCodeLocationOut(TBL, 1, 1, Criteria + "%")
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(4).Visible = False
                DGV.Columns(5).Visible = False
                DGV.Columns(6).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double, ByVal vDGVItemEmployee As DataGridView, _
    ByVal vDGVItemMachine As DataGridView, ByVal vDGVItemSetup As DataGridView, ByVal vObjCD As ClsCekData, ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseMasterWorkUnit.DataTable2DataTable
        Dim row As DseMasterWorkUnit.DataTable2Row
        Dim ObjItem As New ClsMasterWorkUnitItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterWorkUnit.DataTable2Row)
                With row
                    Me.mPrimary = .primary
                    Me.mPrimaryWorkUnitGroup = .primary_work_unit_group
                    Me.mWorkUnitCode = .work_unit_code
                    Me.mDesc = .description
                    Me.mRemarks = .remarks
                    Me.mActive = .active
                    Me.mPrimaryLocationIN = .primary_input_location
                    Me.mPrimaryLocationOUT = .primary_output_location
                    Me.mCodeWUGroup = .Code_WU_Group
                    Me.mDescWUGroup = .Desc_WU_Group
                    Me.CodeLocIN = .Code_Location_In
                    Me.CodeLocOUT = .Code_Location_Out
                    Me.DescLocIN = .Desc_Location_In
                    Me.DescLocOUT = ._Desc_Location__Out_
                    Me.mPrimaryProductionLine = .primary_production_line
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With

                'cari data item employee
                ObjItem.IsiDgvEmployee(Primary, vDGVItemEmployee, i, vObjCD)

                'cari data item machine
                ObjItem.IsiDgvMachine(Primary, vDGVItemMachine, i, vObjCD)
                'cari data item setup
                ObjItem.IsiDgvSetup(Primary, vDGVItemSetup, i, vObjCD)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria, ByVal status As String) As Double
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If status = "Show All" Then
                If vSearchBy = EnumSearchCriteria.ByCode Then
                    RetVal = DTA.ScalarQuery_GetJmlByCode(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    RetVal = DTA.ScalarQuery_GetJmlByDesc(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                    RetVal = DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByMachine Then
                    RetVal = DTA.ScalarQuery_GetJmlByMachineActiveOrNoActive(vKriteria + "%", 1, 0)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByWorkShift Then
                    RetVal = DTA.ScalarQuery_GetJmlByWorkshiftActiveOrNoActive(vKriteria + "%", 1, 0)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByEmployee Then
                    RetVal = DTA.ScalarQuery_GetJmlByEmployeeActvieOrNoActive(vKriteria + "%", 1, 0)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByGroupCode Then
                    RetVal = DTA.ScalarQuery_GetJmlByGroupCodeLIKE(vKriteria + "%", 0, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByGroupDesc Then
                    RetVal = DTA.ScalarQuery_GetJmlByGroupCodeLIKE(vKriteria + "%", 0, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If

                ElseIf vSearchBy = EnumSearchCriteria.ByCodeLocationIn Then
                    RetVal = DTA.ScalarQuery_GetJmlByCodeLocationINLIKE(vKriteria + "%", 0, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If

                ElseIf vSearchBy = EnumSearchCriteria.ByCodeLocationOut Then
                    RetVal = DTA.ScalarQuery_GetJmlByCodeLocationOutLIKE(vKriteria + "%", 0, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                End If
            ElseIf status = "Only Active" Then

                If vSearchBy = EnumSearchCriteria.ByCode Then
                    RetVal = DTA.ScalarQuery_GetJmlByCodeOnlyActive(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    RetVal = DTA.ScalarQuery_GetJmlByDescOnlyActive(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                    RetVal = DTA.ScalarQuery_GetJmlByRemarksOnlyActive(vKriteria + "%")
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByMachine Then
                    RetVal = DTA.ScalarQuery_GetJmlByMachineActiveOrNoActive(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByWorkShift Then
                    RetVal = DTA.ScalarQuery_GetJmlByWorkshiftActiveOrNoActive(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByEmployee Then
                    RetVal = DTA.ScalarQuery_GetJmlByEmployeeActvieOrNoActive(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByGroupCode Then
                    RetVal = DTA.ScalarQuery_GetJmlByGroupCodeLIKE(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByGroupDesc Then
                    RetVal = DTA.ScalarQuery_GetJmlByGroupCodeLIKE(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If

                ElseIf vSearchBy = EnumSearchCriteria.ByCodeLocationIn Then
                    RetVal = DTA.ScalarQuery_GetJmlByCodeLocationINLIKE(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSearchCriteria.ByCodeLocationOut Then
                    RetVal = DTA.ScalarQuery_GetJmlByCodeLocationOutLIKE(vKriteria + "%", 1, 1)
                    If RetVal.HasValue = True Then
                        Return RetVal.Value
                    Else
                        Return 0
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vDGVItemEmployee As DataGridView, ByVal vDGVItemMachine As DataGridView, ByVal vDGVItemSetup As DataGridView, _
    ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim ObjItem As New ClsMasterWorkUnitItem
        'Dim rowDGV As DataGridViewRow
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, True, vDGVItemEmployee, vDGVItemMachine, i) = True Then
                If DTA.InsertQuery(mPrimaryWorkUnitGroup, Me.mWorkUnitCode, Me.mDesc, Me.mRemarks, Me.mActive, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryLocationIN, Me.mPrimaryLocationOUT, Me.PrimaryProductionLine) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mWorkUnitCode).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item Employee
                    ObjItem.SimpanItemEmployee(Me.mPrimary, vDGVItemEmployee, i)

                    'simpan data item Machine
                    ObjItem.SimpanItemMachine(Me.mPrimary, vDGVItemMachine, i)

                    'simpan data item Setup
                    ObjItem.SimpanItemSetup(Me.mPrimary, vDGVItemSetup, i)
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
            ObjItem = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vDGVItemEmployee As DataGridView, ByVal vDGVItemMachine As DataGridView, ByVal vDGVItemSetup As DataGridView, ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim ObjItem As New ClsMasterWorkUnitItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, True, vDGVItemEmployee, vDGVItemMachine, i) = True Then
                If DTA.UpdateQuery(mPrimaryWorkUnitGroup, Me.mWorkUnitCode, Me.mDesc, Me.mRemarks, Me.mActive, vUserCode, Now, Now, Me.mPrimaryLocationIN, Me.mPrimaryLocationOUT, Me.mPrimaryProductionLine, Me.mPrimary, Me.TimeStamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama Employee
                    ObjItem.HapusItemEmployee(Me.mPrimary)

                    'hapus data item lama Machine
                    ObjItem.HapusItemMachine(Me.mPrimary)
                    'hapus data item lama Machine
                    ObjItem.HapusItemSetup(Me.mPrimary)

                    'simpan data item Employee
                    ObjItem.SimpanItemEmployee(Me.mPrimary, vDGVItemEmployee, i)

                    'simpan data item Machine
                    ObjItem.SimpanItemMachine(Me.mPrimary, vDGVItemMachine, i)

                    'simpan data item setup
                    ObjItem.SimpanItemSetup(Me.mPrimary, vDGVItemSetup, i)
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDgvItem As DataGridView, ByVal i As FrmMasterWorkUnit) As Boolean
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim ObjItem As New ClsMasterWorkUnitItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, i.TxtUnitCode.Text) = True Then
                'Hapus item Employee
                ObjItem.HapusItemEmployee(vPrimary)

                'Hapus item Machine
                ObjItem.HapusItemMachine(vPrimary)

                'Hapus item Setup
                ObjItem.HapusItemSetup(vPrimary)

                'Hapus
                DTA.DeleteQuery(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mWorkUnitCode = ""
            mDesc = ""
            mRemarks = ""
            mActive = 0
            mPrimaryLocationIN = -1
            mPrimaryLocationOUT = -1
            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'tambahan batasan untuk allocate mo - 22/juni/2016 09:25
    Public Function apakahsudahada_diallocatemo(ByVal vCode As String) As Boolean
        Dim dta As New Core_Module.DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim retval As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetJmlInAllocateMO(vCode)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
                Return False
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryOutputLocation(ByVal vPrimaryWU As Double) As Double
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryLocationOut1(vPrimaryWU)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryInputLocation(ByVal vPrimaryWU As Double) As Double
        Dim DTA As New DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryLocationIn1(vPrimaryWU)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function GetCodeLocation(ByVal vPrimaryWU As Double) As String
    '    Dim dta As New Core_Module.DseMasterWorkUnitTableAdapters.tblmaster_workunitTableAdapter
    '    Dim retval As string ""
    '    Try
    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        retval = dta.ScalarQuery_GetCodeLocationIN(vPrimaryWU)
    '        If retval.HasValue = True Then
    '            If retval.Value > 0 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '            Return False
    '        End If
    '        Return False
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try

    'End Function
End Class
