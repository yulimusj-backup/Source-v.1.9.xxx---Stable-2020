Option Explicit On
Option Strict Off
Public Class ClsMasterProductAccountGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mGroup_level As EnumGroupLevel
    Dim mPrimary_Group As Double = -1
    Dim mType As EnumTipe
    Dim mRemarks As String = ""
    Dim mPrimary_Inventory_Account As Double = -1
    Dim mPrimary_Sales_Account As Double = -1
    Dim mPrimary_Sales_Return_Account As Double = -1
    Dim mPrimary_Cost_Of_Sales_Account As Double = -1
    Dim mPrimary_Adjustment_Accont As Double = -1
    Dim mPrimary_Stock_taking_Account As Double = -1
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
    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            Me.mCode = value
        End Set
    End Property
    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            Me.mDesc = value
        End Set
    End Property
    Public Property Group_Level() As EnumGroupLevel
        Get
            Return Me.mGroup_level
        End Get
        Set(ByVal value As EnumGroupLevel)
            Me.mGroup_level = value
        End Set
    End Property
    Public Property Primary_Group() As Double
        Get
            Return Me.mPrimary_Group
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Group = value
        End Set
    End Property
    Public Property Type() As EnumTipe
        Get
            Return Me.mType
        End Get
        Set(ByVal value As EnumTipe)
            Me.mType = value
        End Set
    End Property
    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            Me.mRemarks = value
        End Set
    End Property
    Public Property Primary_Inventory_Account() As Double
        Get
            Return Me.mPrimary_Inventory_Account
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Inventory_Account = value
        End Set
    End Property
    Public Property Primary_Sales_Account() As Double
        Get
            Return Me.mPrimary_Sales_Account
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Sales_Account = value
        End Set
    End Property
    Public Property Primary_Cost_Of_Sales_Account() As Double
        Get
            Return Me.mPrimary_Cost_Of_Sales_Account
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Cost_Of_Sales_Account = value
        End Set
    End Property
    Public Property Primary_Adjustment_Accont() As Double
        Get
            Return Me.mPrimary_Adjustment_Accont
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Adjustment_Accont = value
        End Set
    End Property
    Public Property Primary_Stock_taking_Account() As Double
        Get
            Return Me.mPrimary_Stock_taking_Account
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Stock_taking_Account = value
        End Set
    End Property
    Public Property Primary_Sales_Return_Account() As Double
        Get
            Return Me.mPrimary_Sales_Return_Account
        End Get
        Set(ByVal value As Double)
            Me.mPrimary_Sales_Return_Account = value
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

    Public Enum EnumGroupLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum

    Public Enum EnumTipe As Integer
        Goods
        Service
    End Enum
    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

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

    Public Function GetPrimary(ByVal vGroupLevel As Integer, ByVal vPrimaryGroup As Double) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary1(vGroupLevel, vPrimaryGroup)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetTimestampByPrimary(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(Me.mCode, Me.mDesc, Me.mGroup_level, Me.mPrimary_Group, Me.mType, Me.mRemarks, Me.mPrimary_Inventory_Account _
                , Me.mPrimary_Sales_Account, Me.mPrimary_Sales_Return_Account, Me.mPrimary_Cost_Of_Sales_Account, Me.mPrimary_Adjustment_Accont _
                 , Me.mPrimary_Stock_taking_Account, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.Code)
                Me.mTimestamp = Me.GetTimeStamp(CDec(Me.mPrimary))
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
    ''
    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
          EnumSearchCriteria) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_group1TableAdapter

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
            ElseIf vSearchBy = EnumSearchCriteria.ByGroup Then
                If DTA.ScalarQuery_GetJmlGroup(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlGroup(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
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
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Update(Me.mCode, Me.mDesc, Me.mGroup_level, Me.mPrimary_Group, Me.mType, Me.mRemarks, Me.mPrimary_Inventory_Account _
                , Me.mPrimary_Sales_Account, Me.mPrimary_Sales_Return_Account, Me.mPrimary_Cost_Of_Sales_Account, Me.mPrimary_Adjustment_Accont _
                 , Me.mPrimary_Stock_taking_Account, vUserCode, Now, Now, Me.mPrimary) > 0 Then
                Me.mTimestamp = GetTimeStamp(Me.mPrimary)
                Me.mModifiedDate = mTimestamp
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

    Public Function Hapus(ByVal vPrimary As Double, ByVal vSalesStaff As String) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vPrimary, vSalesStaff, True) = True Then
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
    Public Function CekPrimaryGroupDanType(ByVal vType As Integer, ByVal vPrimaryGroup As Double) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring


            If DTA.ScalarQuery_GetJmlByPrimaryAndType(vType, vPrimaryGroup).HasValue Then
                If DTA.ScalarQuery_GetJmlByPrimaryAndType(vType, vPrimaryGroup).Value > 0 Then
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
    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vSalesStaff As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Created_By <> vSalesStaff Then
                If vTampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterProductAccountGroup.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroup Then
                QueryResult = DTA.FillBy_Group(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupProductGroup Then
                QueryResult = DTA.FillBy_GroupCodeLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupBrand Then
                QueryResult = DTA.FillBy_GroupBrandLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupSubBrand Then
                QueryResult = DTA.FillBy_GroupSubBrandLIKE(TBL, Criteria & "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupCategory Then
                QueryResult = DTA.FillBy_GroupCategoryLIKE(TBL, Criteria & "%")
            End If

            'update dgv
            TBL.DefaultView.Sort = "code ASC"
            DGV.DataSource = TBL.DefaultView

            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Group"
                DGV.Columns(4).HeaderText = "Remarks"
                DGV.Columns(3).MinimumWidth = 125
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
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim TBL As New DseMasterProductAccountGroup.tblmaster_product_account_groupDataTable
        Dim row As DseMasterProductAccountGroup.tblmaster_product_account_groupRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterProductAccountGroup.tblmaster_product_account_groupRow)
                With row
                    mPrimary = .primary
                    mCode = .code
                    mDesc = .desc
                    mGroup_level = .group_level
                    mPrimary_Group = .primary_group
                    mType = .type
                    mRemarks = .remarks
                    mPrimary_Inventory_Account = .primary_inventory_account
                    mPrimary_Sales_Account = .primary_sales_account
                    mPrimary_Sales_Return_Account = .primary_sales_return_account
                    mPrimary_Cost_Of_Sales_Account = .primary_cost_of_sales_account
                    mPrimary_Adjustment_Accont = .primary_adjustment_account
                    mPrimary_Stock_taking_Account = .primary_stock_taking_account
                    mModifiedBy = .modified_by
                    mModifiedDate = .modified_date
                    mCreatedBy = .created_by
                    mCreatedDate = .created_date
                    mTimestamp = .timestamp
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
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCodeByPrimary(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetCodeByPrimary(CDec(vPrimary))
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
    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByCode(vCode)) = False Then
                Return DTA.ScalarQuery_GetJmlByCode(vCode)
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
    Public Overridable Function ApakahBolehEdit(ByVal vPrimary As Double, _
        ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'batasan edit 
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID : " + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetTypeLevel(ByVal vPrimary As Double) As EnumTipe
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetType(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value = 0 Then
                    Return EnumTipe.Goods
                Else
                    Return EnumTipe.Service
                End If
            Else
                Return EnumTipe.Goods
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipe.Goods
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function CekPrimaryGroup(ByVal vPrimaryGroup As Double) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryGroup(vPrimaryGroup).HasValue Then
                If DTA.ScalarQuery_GetJmlByPrimaryGroup(vPrimaryGroup).Value > 0 Then
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

    '
    '
    '###fungsi dibawah adalah fungsi tambahan yg digunakan di frm masterProduct --->
    '
    '

    Public Function GetPrimaryAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryInvAcc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryInvAcc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimarySalesAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimarySalesAcc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetPrimarySalesAcc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimarySalesReturnAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimarySalesreturnAcc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetPrimarySalesreturnAcc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimaryCOSAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetCOSacc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetCOSacc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimaryAdjustmentAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAdjustmentAcc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryAdjustmentAcc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimaryStockTakingAcc(ByVal vPrimary As Double, ByVal vLvl As Integer, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryStockTakingAcc(vPrimary, vLvl, vType).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryStockTakingAcc(vPrimary, vLvl, vType).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahAda(ByVal vGroupLevel As Integer, ByVal vPrimaryGroup As Double, ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByGroupLevel(vGroupLevel, vPrimaryGroup, vType).HasValue Then
                If DTA.ScalarQuery_GetJmlByGroupLevel(vGroupLevel, vPrimaryGroup, vType).Value > 0 Then
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


    Public Function ApakahAdaPrimary(ByVal vPrimary As Double, _
    ByVal vGroupLevel As Integer, ByVal vPrimaryGroup As Double, _
    ByVal vType As Core_Module.ClsMasterProductAccountGroup.EnumTipe) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryNotSame(vPrimary, vGroupLevel, vPrimaryGroup, vType).HasValue Then
                If DTA.ScalarQuery_GetJmlByPrimaryNotSame(vPrimary, vGroupLevel, vPrimaryGroup, vType).Value > 0 Then
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


    Public Sub IsiTBL(ByVal TBL As DseMasterProductAccountGroup.tblisi_cboDataTable, ByVal vType As ClsMasterProductAccountGroup.EnumTipe)
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblisi_cboTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_tipe(TBL, vType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
    Public Function GetPrimaryGroup(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryGroup(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryGroup(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    '
    Public Function GetGroupLevelByPrimaryGroup(ByVal vPrimaryGroup As Double) As EnumGroupLevel
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetGroupLevelBy(vPrimaryGroup)
            If Dummy.HasValue = True Then
                If Dummy.Value = 0 Then
                    Return EnumGroupLevel.Category
                ElseIf Dummy.Value = 1 Then
                    Return EnumGroupLevel.SubBrand
                ElseIf Dummy.Value = 2 Then
                    Return EnumGroupLevel.Brand
                Else
                    Return EnumGroupLevel.ProductGroup
                End If
            Else
                Return EnumTipe.Goods
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipe.Goods
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTypeByPrimaryGroup(ByVal PrimaryGroup As Double) As EnumTipe
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetTypeBy(PrimaryGroup)
            If Dummy.HasValue = True Then
                If Dummy.Value = 0 Then
                    Return EnumTipe.Goods
                Else
                    Return EnumTipe.Service
                End If
            Else
                Return EnumTipe.Goods
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipe.Goods
        Finally
            DTA.Dispose()
        End Try
    End Function
    '
    Public Function GetGroupLevel(ByVal vPrimary As Double) As EnumGroupLevel
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetGroupLevel(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value = 0 Then
                    Return EnumGroupLevel.Category
                ElseIf Dummy.Value = 1 Then
                    Return EnumGroupLevel.SubBrand
                ElseIf Dummy.Value = 2 Then
                    Return EnumGroupLevel.Brand
                Else
                    Return EnumGroupLevel.ProductGroup
                End If
            Else
                Return EnumTipe.Goods
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipe.Goods
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimaryByDesc(ByVal vDesc As String) As Double
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryByDesc(vDesc).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByDesc(vDesc).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    '---
End Class
