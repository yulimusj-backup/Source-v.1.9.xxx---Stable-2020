Option Explicit On
Option Strict Off
Public Class clsMasterStorageArea
    Dim mPrimary As Double = -1
    Dim mStorageAreaCode As String = ""
    Dim mDescription As String = ""
    Dim mPrimaryParent As Double = -1
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

    Public Property StorageAreaCode() As String
        Get
            Return Me.mStorageAreaCode
        End Get
        Set(ByVal value As String)
            mStorageAreaCode = value
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

    Public Property PrimaryParent() As Double
        Get
            Return Me.mPrimaryParent
        End Get
        Set(ByVal value As Double)
            mPrimaryParent = value
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
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.Insert(Me.mPrimaryParent, Me.mStorageAreaCode, Me.mDescription, _
            vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mStorageAreaCode).Value
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
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.Update(Me.mPrimaryParent, Me.mStorageAreaCode, Me.mDescription, _
            vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
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
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
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

    Public Function ApakahSudahAda(ByVal vStorageAreaCode As String) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vStorageAreaCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vStorageAreaCode).Value > 0 Then
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
        Dim Obj As New clsMasterProduct
        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. product 
            If Obj.GetJml_ByStorageArea(vPrimary) > 0 Then
                Return False
            Else
                Return True
            End If
            Obj = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterStorageArea.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
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
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBL As New dseMasterStorageArea.tblmaster_storage_areaDataTable
        Dim row As dseMasterStorageArea.tblmaster_storage_areaRow

        Try
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
                    Me.mPrimaryParent = .Primary_Warehouse
                    Me.mStorageAreaCode = .Storage_Area_Code
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

    ''' <summary>
    ''' mengembalikan jumlah storage area yang ada
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml() As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlTotal.HasValue = True Then
                Return DTA.ScalarQuery_GetJmlTotal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_LocationCodeLIKE(ByVal vLocationCodeLIKE As String) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.ScalarQuery_GetJml_ByCodeLIKE(vLocationCodeLIKE).HasValue = True Then
                Return DTA.ScalarQuery_GetJml_ByCodeLIKE(vLocationCodeLIKE).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearchCriteria) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml1(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml1(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml2(vKriteria + "%").Value
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


    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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

    Public Function GetStorageAreaCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCode(vPrimary)
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

    Public Function GetJml_ByWarehouse(ByVal vPrimaryWarehouse As Double) As Integer
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.ScalarQuery_GetJml3(vPrimaryWarehouse).HasValue Then
                Return DTA.ScalarQuery_GetJml3(vPrimaryWarehouse).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Try
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetWarehouseCode(ByVal vPrimaryStorageArea As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetWarehouseCode(vPrimaryStorageArea)) = False Then
                Return DTA.ScalarQuery_GetWarehouseCode(vPrimaryStorageArea)
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

    Public Function GetPrimary(ByVal vStorageAreaCode As String) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vStorageAreaCode).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vStorageAreaCode).Value
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
End Class
