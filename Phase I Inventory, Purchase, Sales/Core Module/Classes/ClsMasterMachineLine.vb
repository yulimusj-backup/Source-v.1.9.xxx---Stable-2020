Option Explicit On
Option Strict Off
Public Class ClsMasterMachineLine
    Dim ObjCD As New Core_Module.ClsCekData

    Enum EnumSearchCriteria As Integer
        BY_CODE
        BY_DESC
        BY_REMARKS
    End Enum

    Private mPrimary As Double = -1
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Private mCode As String = ""
    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Private mDescription As String = ""
    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Private mRemarks As String = ""
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Private mCreatedBy As String = ""
    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Private mCreatedDate As Date = Now
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Private mModifiedBy As String = ""
    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Private mModifiedDate As Date = Now
    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Private mTimestamp As Date = Now
    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.Code) = True Then
                    If vShowMsg = True Then
                        MsgBox("Code: " & Me.Code & " already exist.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim Kode_ As String = Me.GetCode(Me.Primary)
                If Kode_ <> Me.Code Then
                    If Me.ApakahSudahAda(Me.Code) = True Then
                        If vShowMsg = True Then
                            MsgBox("Code: " & Me.Code & " already exist.", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            If String.IsNullOrEmpty(mCode) = True Then
                If vShowMsg = True Then
                    MsgBox("Code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            If String.IsNullOrEmpty(mDescription) = True Then
                If vShowMsg = True Then
                    MsgBox("Description cannot be an empty string.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetCountByCode(vCode)
            If Dummy_.HasValue = True Then
                If Dummy_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, Optional ByVal vShowMessage As Boolean = True) As Boolean
        Dim dta As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetCountInPredefined(vPrimary)
            If n.HasValue Then
                If n.Value > 0 Then
                    If vShowMessage Then
                        MsgBox("This master already used in work unit predefined.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            If Me.GetJmlInWorkUnit(vPrimary) > 0 Then
                MsgBox("This master already used in master work unit.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetJmlInWorkUnit(ByVal vPrimaryLine As Double) As Integer
        Dim d As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim val As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            val = d.ScalarQuery_GetCountInWorkUnitByPrimaryLine(vPrimaryLine)
            If val.HasValue Then
                Return val.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True) = True Then
                If DTA.InsertQuery(Me.Code, Me.Description, Me.Remarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False) = True Then
                If DTA.UpdateQuery(Me.Code, Me.Description, Me.Remarks, vUserCode, Now, Now, Me.mPrimary) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
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
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim TBL As New DseMasterMachineLine.tblmaster_machine_lineDataTable
        Dim row As DseMasterMachineLine.tblmaster_machine_lineRow
        Try
            DTA.Fill(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDescription = .description
                    Me.mRemarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetPrimary(vCode)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetCode(vPrimary)
            If String.IsNullOrEmpty(Dummy_) = False Then
                Return Dummy_
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

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Dummy_) = False Then
                Return Dummy_
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
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Date.MinValue
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As DataGridView, ByVal TBL As DseMasterMachineLine.tblmaster_machine_line_findDataTable, _
        ByVal SearchCriteria As ClsMasterMachineLine.EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_line_findTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.BY_CODE Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BY_DESC Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BY_REMARKS Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vCode As String) As Integer
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_lineTableAdapter
        Dim Dummy_ As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy_ = DTA.ScalarQuery_GetCountByCode(vCode)
            If Dummy_.HasValue = True Then
                Return Dummy_.Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ClsMasterMachineLine.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterMachineLineTableAdapters.tblmaster_machine_line_findTableAdapter
        Dim v As New Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.BY_CODE Then
                v = DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.BY_DESC Then
                v = DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.BY_REMARKS Then
                v = DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%")
                If v.HasValue Then
                    Return v.Value
                Else
                    Return 0
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
