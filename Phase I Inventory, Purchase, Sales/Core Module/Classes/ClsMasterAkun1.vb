Option Explicit On
Option Strict Off

Public Class ClsMasterAkun1
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mAccountCode As String = ""
    Private mDescription As String = ""
    Private mAccountNature As String = ""
    Private mAktif As Integer = 1
    Private mHeader As Integer = 0
    Private mBeginningBalance As Decimal = 0
    Private mTipePerkiraan As EnumTipeAkun = EnumTipeAkun.Asset
    Private mLevelPerkiraan As Integer = 1
    Private mTampilan As String = ""
    Private mSaldoNormal As EnumSaldoNormal = EnumSaldoNormal.Debet
    Private mUrutan As Integer = 0
    Private mPrimaryLv1 As Double = -1
    Private mPrimaryLv2 As Double = -1
    Private mPrimaryLv3 As Double = -1
    Private mTipeAkun2 As Integer = 0
    Private mBeginningBalanceOriginalCurrency As Decimal = 0
    Private mPrimaryCurrency As Double = -1

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property SaldoNormal() As EnumSaldoNormal
        Get
            Return mSaldoNormal
        End Get
        Set(ByVal value As EnumSaldoNormal)
            mSaldoNormal = value
        End Set
    End Property

    Public Enum EnumSaldoNormal As Integer
        Debet
        Kredit
    End Enum

    Public Property TipePerkiraan() As EnumTipeAkun
        Get
            Return mTipePerkiraan
        End Get
        Set(ByVal value As EnumTipeAkun)
            mTipePerkiraan = value
        End Set
    End Property

    Public Enum EnumTipeAkun As Integer
        Asset
        Liabilities
        Equity
        Income
        Expense
        Cost_of_Goods_Sales
        Other_Income
        Other_Expense
        Tax
    End Enum

    Public Enum EnumTipeAkun2 As Integer
        Cash_And_Bank
        Account_Receivable
        Other_Receivable
        Inventory
        Fixed_Asset
        Accumulated_Depreciation
        Other_Asset
        Account_Payable
        Other_Current_Liability
        Long_Term_Liability
        Equity
        Income
        Cost_Of_Sales
        Expense
        Other_Income
        Other_Expense
        Tax
        Asset
        Liabilities
        Other_Payable
    End Enum

    Public Property TipeAkun2() As EnumTipeAkun2
        Get
            Return Me.mTipeAkun2
        End Get
        Set(ByVal value As EnumTipeAkun2)
            mTipeAkun2 = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
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

    Public Property AccountCode() As String
        Get
            Return mAccountCode
        End Get
        Set(ByVal value As String)
            mAccountCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property AccountNature() As String
        Get
            Return mAccountNature
        End Get
        Set(ByVal value As String)
            mAccountNature = value
        End Set
    End Property

    Public Property Aktif() As Integer
        Get
            Return mAktif
        End Get
        Set(ByVal value As Integer)
            mAktif = value
        End Set
    End Property

    Public Property Header() As Integer
        Get
            Return mHeader
        End Get
        Set(ByVal value As Integer)
            mHeader = value
        End Set
    End Property

    Public Property BeginningBalanceOriginalCurrency() As Decimal
        Get
            Return Me.mBeginningBalanceOriginalCurrency
        End Get
        Set(ByVal value As Decimal)
            mBeginningBalanceOriginalCurrency = value
        End Set
    End Property

    Public Property BeginningBalance() As Decimal
        Get
            Return mBeginningBalance
        End Get
        Set(ByVal value As Decimal)
            mBeginningBalance = value
        End Set
    End Property

    Public Property Tampilan() As String
        Get
            Return mTampilan
        End Get
        Set(ByVal value As String)
            mTampilan = value
        End Set
    End Property

    Public Property LevelPerkiraan() As Integer
        Get
            Return mLevelPerkiraan
        End Get
        Set(ByVal value As Integer)
            mLevelPerkiraan = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property PrimaryLv1() As Double
        Get
            Return mPrimaryLv1
        End Get
        Set(ByVal value As Double)
            mPrimaryLv1 = value
        End Set
    End Property

    Public Property PrimaryLv2() As Double
        Get
            Return mPrimaryLv2
        End Get
        Set(ByVal value As Double)
            mPrimaryLv2 = value
        End Set
    End Property

    Public Property PrimaryLv3() As Double
        Get
            Return mPrimaryLv3
        End Get
        Set(ByVal value As Double)
            mPrimaryLv3 = value
        End Set
    End Property

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).Value
            Else
                Return EnumTipeAkun2.Asset
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeAkun2.Asset
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeAccount2(ByVal vPrimary As Double) As EnumTipeAkun2
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAkun(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkun(vPrimary).Value
            Else
                Return EnumTipeAkun2.Asset
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumTipeAkun2.Asset
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAccountCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)
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

    Public Function GetDescription_ByPrimaryAcc(ByVal vPrimary_Acc As String) As String
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary_Acc)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary_Acc)
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

    Public Function GetPrimary(ByVal vAccountCode As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vAccountCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vAccountCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).Value > 0 Then
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

    Public Function ApakahSudahAda_Aktif_NonHeader(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).Value > 0 Then
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

    Public Function ApakahSudahAda_Aktif_Header(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_Header(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJml_Aktif_Header(vAccountCode).Value > 0 Then
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByAccountNo Then
                Return DTA.ScalarQuery_GetJmlByAccountCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlHeader(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByAccountNo Then
                Return DTA.ScalarQuery_GetJmlByAccountCodeLIKE_Header(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE_Header(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV2(ByVal Criteria As String, ByVal DGV As DataGridView, _
   ByVal TBL As DseMasterAkun_1.tblmaster_akunDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
   ByVal vAkunAsset As ClsMasterAsset.EnumAkunAsset) As Boolean
        Dim DTA As New DseMasterAkun_1TableAdapters.tblmaster_akunTableAdapter()
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.OnProcess Or _
                vAkunAsset = ClsMasterAsset.EnumAkunAsset.Asset Then
                    QueryResult = DTA.FillBy_FixedAsset(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.DeprciationExpense Then
                    QueryResult = DTA.FillBy_ExpenseNOtherExpense(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.AccumDepreciation Then
                    QueryResult = DTA.FillBy_AccumDepLikeCode(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.DisposalIncome Or _
                vAkunAsset = ClsMasterAsset.EnumAkunAsset.GainLossDisposal Then
                    QueryResult = DTA.FillBy_IncomeNExpense(TBL, Criteria + "%")
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.OnProcess Or _
                vAkunAsset = ClsMasterAsset.EnumAkunAsset.Asset Then
                    QueryResult = DTA.FillBy_FixedAssetLikeDesc(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.DeprciationExpense Then
                    QueryResult = DTA.FillBy_ExpenseLikeDesc(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.AccumDepreciation Then
                    QueryResult = DTA.FillBy_AccumDepLikeDesc(TBL, Criteria + "%")
                End If
                If vAkunAsset = ClsMasterAsset.EnumAkunAsset.DisposalIncome Or _
                vAkunAsset = ClsMasterAsset.EnumAkunAsset.GainLossDisposal Then
                    QueryResult = DTA.FillBy_IncomeNExpenseLikeDesc(TBL, Criteria + "%")
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Account Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterAkun.tbl_Acari_master_akunDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccountCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Account Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiDGVHeader(ByVal Criteria As String, _
  ByVal DGV As DataGridView, ByVal TBL As DseMasterAkun.tbl_Acari_master_akunDataTable, _
  ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccountCodeLIKE_Header(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE_Header(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Account Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary)
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetLevelPerkiraan(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetLevelPerkiraan(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetLevelPerkiraan(vPrimaryAkun).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeAkun(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipePerkiraan(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetTipePerkiraan(vPrimaryAkun).Value
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

    Public Function GetTipeAkun2(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipePerkiraan2(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetTipePerkiraan2(vPrimaryAkun).Value
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

    Public Function GetUrutan(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetUrutan(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetUrutan(vPrimaryAkun).Value
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

    Public Function GetSaldoNormal(ByVal vPrimaryAkun As Double) As Integer
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetSaldoNormal(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetSaldoNormal(vPrimaryAkun).Value
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

    Public Function GetPrimaryLevel1(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryLevel1(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryLevel1(vPrimaryAkun).Value
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

    Public Function GetPrimaryLevel2(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryLevel2(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryLevel2(vPrimaryAkun).Value
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

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter
        Dim TBL As New DseMasterAkun.tblmaster_akun1DataTable
        Dim row As DseMasterAkun.tblmaster_akun1Row
        Dim Start As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.AccountCode, Me.Description, Me.AccountNature, _
                Me.Aktif, Me.Header, Me.BeginningBalance, Me.TipePerkiraan, Me.LevelPerkiraan, _
                Me.Tampilan, Me.SaldoNormal, Me.Urutan, Me.PrimaryLv1, Me.PrimaryLv2, Me.PrimaryLv3, _
                vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryCurrency, Me.mBeginningBalanceOriginalCurrency, Me.mTipeAkun2) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mAccountCode)
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'update data urutan
                DTA.FillBy_Urutan(TBL, Me.mUrutan, Me.mPrimary)
                Start = Me.Urutan + 1
                For Each row In TBL.Rows
                    DTA.UpdateQuery_UpdateUrutan(Start, row.Primary)
                    Start += 1
                Next

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJmlAnakLvl1(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlAnakLvl1(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlAnakLvl1(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAnakLvl2(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlAnakLvl2(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlAnakLvl2(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAnakLvl3(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlAnakLvl3(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlAnakLvl3(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUrutanMaximalLvl1(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetUrutanMaxLvl1(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetUrutanMaxLvl1(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUrutanMaximalLvl2(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetUrutanMaxLvl2(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetUrutanMaxLvl2(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUrutanMaximalLvl3(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetUrutanMaxLvl3(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetUrutanMaxLvl3(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
