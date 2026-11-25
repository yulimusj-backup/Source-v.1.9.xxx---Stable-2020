Option Explicit On
Option Strict Off

Public Class ClsMasterCategory
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCategoryCode As String = ""
    Dim mDescription As String = ""
    Dim mRemarks As String = ""
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

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property
    Public Property CategoryCode() As String
        Get
            Return Me.mCategoryCode
        End Get
        Set(ByVal value As String)
            mCategoryCode = value
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
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCategoryCode, Me.mDescription, _
            Me.mPrimaryParent, vUserCode, Now, vUserCode, Now, Now, Me.mRemarks) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCategoryCode).Value
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

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal KodeSama As Boolean, ByVal DescSama As Boolean, _
    ByVal Prb As ToolStripProgressBar) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL1 As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row
        Dim Jml As Int64 = 0
        Dim CategoryCodeLama, CategoryDescLama As String
        Dim Obj As New clsMasterProduct

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            CategoryCodeLama = DTA.ScalarQuery_GetCode(Me.mPrimary)
            CategoryDescLama = DTA.ScalarQuery_GetDesc(Me.mPrimary)

            If DTA.Update(Me.mCategoryCode, Me.mDescription, _
            Me.mPrimaryParent, vUserCode, Now, Now, Me.mRemarks, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                Prb.Value = 0
                Prb.Maximum = 0
                If KodeSama = False Or DescSama = False Then
                    Jml = DTA1.FillBy_PrimaryCetegory(TBL1, Me.mPrimary)
                End If
                If KodeSama = False Then
                    Prb.Maximum += Jml
                End If
                If DescSama = False Then
                    Prb.Maximum += Jml
                End If

                If KodeSama = False Then
                    Dim KodeLama As String
                    Dim KodeBaru As String

                    For Each row In TBL1.Rows
                        If row.Category_Code_Bounded = True Then
                            KodeLama = row.Product_Code
                            KodeBaru = Me.CategoryCode + CStr(KodeLama.Substring(CategoryCodeLama.Length))
                            DTA1.UpdateQuery_SetProductCode(KodeBaru, vUserCode, Now, Now, row.Primary, row.Timestamp)
                            row.Timestamp = Obj.GetTimeStamp(row.Primary)
                        End If
                        Prb.Value += 1
                    Next
                End If
                If DescSama = False Then
                    Dim DescLama As String
                    Dim DescBaru As String

                    For Each row In TBL1.Rows
                        If row.Category_Desc_Bounded = True Then
                            DescLama = row.Description
                            DescBaru = Me.Description + " " + DescLama.Substring(CategoryDescLama.Length + 1)
                            DTA1.UpdateQuery_SetProductDesc(vUserCode, Now, Now, DescBaru, row.Primary, row.Timestamp)
                        End If
                        Prb.Value += 1
                    Next
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

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

    Public Function ApakahSudahAda(ByVal vCategoryCode As String, ByVal vPrimarySubBrand As String) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml4(vCategoryCode, vPrimarySubBrand).HasValue Then
                If DTA.ScalarQuery_GetJml4(vCategoryCode, vPrimarySubBrand).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vCategoryCode As String) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJml(vCategoryCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vCategoryCode).Value > 0 Then
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
            '1. Product 
            If Obj.GetJml(vPrimary) > 0 Then
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

    Public Function IsiListBox(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterCategory.tblmaster_category1DataTable) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_category1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterCategory.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterCategoryTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CategoryCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_CategoryDescLIKE(TBL, Criteria + "%")
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
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter
        Dim TBL As New dseMasterCategory.tblmaster_categoryDataTable
        Dim row As dseMasterCategory.tblmaster_categoryRow

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
                    Me.mRemarks = .Remarks
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mPrimaryParent = .Parent_Group_Primary
                    Me.mCategoryCode = .Category_Code
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

    Public Function GetJml(ByVal vPrimaryParent As Double) As Integer
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml3(vPrimaryParent).HasValue Then
                Return DTA.ScalarQuery_GetJml3(vPrimaryParent).Value
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
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetCategoryCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCategoryCode As String) As Double
        Dim DTA As New dseMasterCategoryTableAdapters.tblmaster_categoryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCategoryCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCategoryCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryBrand(ByVal vprimarycategory As Double) As Double
        Dim dta As New dseMasterCategoryTableAdapters.Tbl_Lap_RankTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimaryBrand(vprimarycategory).HasValue Then
                Return dta.ScalarQuery_GetPrimaryBrand(vprimarycategory).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryProductGroup(ByVal vPrimaryCategory As Double) As Double
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetBrandCode(ByVal vprimarycategory As Double) As String
        Dim dta As New dseMasterCategoryTableAdapters.Tbl_Lap_RankTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetBrandCode(vprimarycategory)) = False Then
                Return dta.ScalarQuery_GetBrandCode(vprimarycategory)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetBrandDescription(ByVal vprimarycategory As Double) As String
        Dim dta As New dseMasterCategoryTableAdapters.Tbl_Lap_RankTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetBrandDescription(vprimarycategory)) = False Then
                Return dta.ScalarQuery_GetBrandDescription(vprimarycategory)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
