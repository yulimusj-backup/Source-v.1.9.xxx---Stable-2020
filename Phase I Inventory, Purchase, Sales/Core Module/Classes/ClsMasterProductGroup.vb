Option Explicit On
Option Strict Off

Public Class ClsMasterProductGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mProductGroupCode As String = ""
    Dim mDescription As String = ""
    Dim mRemarks As String = ""
    ''' <summary>
    ''' nilai product group level silakan start dari 1 -> product group
    ''' </summary>
    ''' <remarks></remarks>
    Dim mProductGroupLevel As Integer = -1
    Dim mPrimaryParentProductGroup As Double = -1
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mKodeAkun As String = ""

    Public Property KodeAkun() As String
        Get
            Return Me.mKodeAkun
        End Get
        Set(ByVal value As String)
            Me.mKodeAkun = value
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

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property ProductGroupCode() As String
        Get
            Return Me.mProductGroupCode
        End Get
        Set(ByVal value As String)
            mProductGroupCode = value
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

    Public Property ProductGroupLevel() As Integer
        Get
            Return Me.mProductGroupLevel
        End Get
        Set(ByVal value As Integer)
            mProductGroupLevel = value
        End Set
    End Property

    Public Property PrimaryParentProductGroup() As Double
        Get
            Return Me.mPrimaryParentProductGroup
        End Get
        Set(ByVal value As Double)
            mPrimaryParentProductGroup = value
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
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mProductGroupCode, Me.mDescription, Me.mProductGroupLevel, _
            Me.mPrimaryParentProductGroup, vUserCode, Now, vUserCode, Now, Now, Me.mRemarks, Me.mKodeAkun) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mProductGroupCode, Me.mProductGroupLevel).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
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
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mProductGroupCode, Me.mDescription, Me.mProductGroupLevel, _
            Me.mPrimaryParentProductGroup, vUserCode, Now, Now, Me.mRemarks, Me.mKodeAkun, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
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
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)
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

    Public Function ApakahSudahAda(ByVal vProductGroupCode As String, ByVal vProductGroupLevel As Integer) As Boolean
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml3(vProductGroupCode, vProductGroupLevel).HasValue Then
                If DTA.ScalarQuery_GetJml3(vProductGroupCode, vProductGroupLevel).Value > 0 Then
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
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh dihapus bila sudah ada yang merefer:
            Dim GroupLevel As Integer = Me.GetProductGroupLevel(vPrimary, DTA)
            If GroupLevel <> 1 Then
                '1. Product Group lain
                If DTA.ScalarQuery_GetJml2(GroupLevel - 1, vPrimary).HasValue Then
                    If DTA.ScalarQuery_GetJml2(GroupLevel - 1, vPrimary).Value > 0 Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            Else
                '2. Product Category
                Dim Obj As New ClsMasterCategory
                If Obj.GetJml(vPrimary) > 0 Then
                    Return False
                Else
                    Return True
                End If
                Obj = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetProductGroupLevel(ByVal Primary As Double, _
    ByVal DTA As dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter) As Integer
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetGroupLevel(Primary).HasValue Then
                Return DTA.ScalarQuery_GetGroupLevel(Primary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal ProductGroupLevel As Integer, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterProductGroup.tblmaster_product_group1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_group1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_LikeGroupCode(TBL, ProductGroupLevel, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_LikeDescription(TBL, Criteria + "%", ProductGroupLevel)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()

            DGV.Columns(0).Visible = False
            If ProductGroupLevel = 3 Then
                DGV.Columns(1).HeaderText = "Product Group Code"
                DGV.Columns(3).Visible = False
            ElseIf ProductGroupLevel = 2 Then
                DGV.Columns(1).HeaderText = "Brand Code"
                DGV.Columns(3).Visible = True
                DGV.Columns(3).HeaderText = "Product Group Code"
            ElseIf ProductGroupLevel = 1 Then
                DGV.Columns(1).HeaderText = "Sub Brand Code"
                DGV.Columns(3).Visible = True
                DGV.Columns(3).HeaderText = "Brand Code"
            End If

            If DGV.Columns.Count > 0 Then
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
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter
        Dim TBL As New dseMasterProductGroup.tblmaster_product_groupDataTable
        Dim row As dseMasterProductGroup.tblmaster_product_groupRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mPrimaryParentProductGroup = .Parent_Group_Primary
                    Me.mProductGroupCode = .Group_Code
                    Me.mProductGroupLevel = .Group_Level
                    Me.mTimestamp = .Timestamp
                    Me.mRemarks = .Remarks
                    Me.mKodeAkun = .Kode_Akun
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vProductGroupCode As String, ByVal vGroupLevel As Integer) As Double
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vProductGroupCode, vGroupLevel).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vProductGroupCode, vGroupLevel).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vGroupLevel As Integer, ByVal vSearchBy As _
    EnumSearchCriteria) As Double
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_group1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml1(vKriteria + "%", vGroupLevel).HasValue Then
                    Return DTA.ScalarQuery_GetJml1(vKriteria + "%", vGroupLevel).Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2(vKriteria + "%", vGroupLevel).HasValue Then
                    Return DTA.ScalarQuery_GetJml2(vKriteria + "%", vGroupLevel).Value
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

    Public Function GetGroupDesc(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetGroupCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetGroupCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetGroupCode(vPrimary)
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

    Public Function IsiListBoxBrand(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterProductGroup.tblmaster_product_group2DataTable) As Boolean
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_group2TableAdapter
        Dim row As dseMasterProductGroup.tblmaster_product_group2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Brand + " - " + row.Description
            Next
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiCboProductGroup(ByVal cbo As ComboBox, ByVal TBL As dseMasterProductGroup.tblmaster_product_group2DataTable)
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_group2TableAdapter
        Dim row As dseMasterProductGroup.tblmaster_product_group2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL)
            cbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Brand + " - " + row.Description
            Next
            cbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiListBoxProductGroup(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterProductGroup.tblmaster_product_group2DataTable) As Boolean
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_group2TableAdapter
        Dim row As dseMasterProductGroup.tblmaster_product_group2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL)
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Brand + " - " + row.Description
            Next
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBLProductSubGroup(ByVal vGroupLevel As Integer, _
    ByVal TBL As dseMasterProductGroup.tblmaster_product_groupDataTable)
        Dim DTA As New dseMasterProductGroupTableAdapters.tblmaster_product_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.ClearBeforeFill = True
            DTA.FillBy_GroupLevel(TBL, vGroupLevel)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
