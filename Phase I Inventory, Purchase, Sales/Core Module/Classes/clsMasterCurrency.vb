Option Explicit On
Option Strict Off

Public Class clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCurrencyCode As String = ""
    Dim mDescription As String = ""
    Dim mDefaultPrecision As Integer = -1
    Dim mPrimaryPayableAccount As Double = -1
    Dim mPrimaryReceivableAccount As Double = -1
    Dim mPrimaryUnrealizedExchangGainLossAccount As Double = -1
    Dim mPrimaryRealizedExchangeGainLossAccount As Double = -1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property PrimaryPayableAccount() As Double
        Get
            Return Me.mPrimaryPayableAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryPayableAccount = value
        End Set
    End Property

    Public Property PrimaryReceivableAccount() As Double
        Get
            Return Me.mPrimaryReceivableAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryReceivableAccount = value
        End Set
    End Property

    Public Property PrimaryUnrealizedExchangeGainLossAccount() As Double
        Get
            Return Me.mPrimaryUnrealizedExchangGainLossAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryUnrealizedExchangGainLossAccount = value
        End Set
    End Property

    Public Property PrimaryRealizedExchangeGainLossAccount() As Double
        Get
            Return Me.mPrimaryRealizedExchangeGainLossAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryRealizedExchangeGainLossAccount = value
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

    Public Property CurrencyCode() As String
        Get
            Return Me.mCurrencyCode
        End Get
        Set(ByVal value As String)
            mCurrencyCode = value
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

    Public Property DefaultPrecision() As Integer
        Get
            Return Me.mDefaultPrecision
        End Get
        Set(ByVal value As Integer)
            Me.mDefaultPrecision = value
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

    Public Sub Set_DefaultExchangeRate(ByVal vPrimaryCurr As Double, ByVal vExchangeRate As Decimal)
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetDefaultExchangeRate(vExchangeRate, vPrimaryCurr)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal vFrm As FrmMasterCurrency2) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Dim ObjAkun As New ClsMasterAkun1
        Dim PrimaryAkun As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCurrencyCode, Me.mDescription, Me.mDefaultPrecision, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryPayableAccount, _
            Me.mPrimaryReceivableAccount, Me.mPrimaryUnrealizedExchangGainLossAccount, Me.mPrimaryRealizedExchangeGainLossAccount) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCurrencyCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan akun payable
                PrimaryAkun = Me.GenerateAkun1(vFrm.TxtPrimaryPayableAccountHeader.Text, ObjAkun, vFrm.TxtPayableAccountHeaderNew1.Text, _
                vFrm.TxtPayableAccountHeaderNew2.Text, vFrm.TxtPayableAccountHeaderNewDesc.Text, vUserCode, ClsMasterAkun1.EnumTipeAkun2.Account_Payable)
                If PrimaryAkun <> -1 Then
                    DTA.UpdateQuery_UpdatePayableAccount(PrimaryAkun, Me.mPrimary)
                    Me.mPrimaryPayableAccount = PrimaryAkun
                End If
                'simpan akun receivable
                PrimaryAkun = Me.GenerateAkun1(vFrm.TxtPrimaryReceivableAccountHeader.Text, ObjAkun, vFrm.TxtReceivableAccountHeaderNew1.Text, _
                vFrm.TxtReceivableAccountHeaderNew2.Text, vFrm.TxtReceivableAccountHeaderNewDesc.Text, vUserCode, ClsMasterAkun1.EnumTipeAkun2.Account_Receivable)
                If PrimaryAkun <> -1 Then
                    DTA.UpdateQuery_UpdateReceivablesAccount(PrimaryAkun, Me.mPrimary)
                    Me.mPrimaryReceivableAccount = PrimaryAkun
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjAkun = Nothing
        End Try
    End Function

    Private Function GenerateAkun1(ByVal PrimaryAkunRef As Double, ByVal Objakun As ClsMasterAkun1, ByVal vAkunCode1 As String, ByVal vAkunCode2 As String, _
    ByVal VAkunDesc As String, ByVal vUserCode As String, ByVal vTipeAkun2 As ClsMasterAkun1.EnumTipeAkun2) As Double
        Dim PrimaryLevel1, PrimaryLevel2, PrimaryLevel3, LevelPerkiraan, Urutan As Double

        Try
            LevelPerkiraan = Objakun.GetLevelPerkiraan(PrimaryAkunRef)
            Urutan = Objakun.GetUrutan(PrimaryAkunRef)
            If LevelPerkiraan = 1 Then
                PrimaryLevel1 = PrimaryAkunRef
                PrimaryLevel2 = -1
                PrimaryLevel3 = -1
                If Objakun.GetJmlAnakLvl1(PrimaryAkunRef) > 0 Then
                    Urutan = Objakun.GetUrutanMaximalLvl1(PrimaryAkunRef)
                End If
            ElseIf LevelPerkiraan = 2 Then
                PrimaryLevel1 = Objakun.GetPrimaryLevel1(PrimaryAkunRef)
                PrimaryLevel2 = PrimaryAkunRef
                PrimaryLevel3 = -1
                If Objakun.GetJmlAnakLvl2(PrimaryAkunRef) > 0 Then
                    Urutan = Objakun.GetUrutanMaximalLvl2(PrimaryAkunRef)
                End If
            ElseIf LevelPerkiraan = 3 Then
                PrimaryLevel1 = Objakun.GetPrimaryLevel1(PrimaryAkunRef)
                PrimaryLevel2 = Objakun.GetPrimaryLevel2(PrimaryAkunRef)
                PrimaryLevel3 = PrimaryAkunRef
                If Objakun.GetJmlAnakLvl3(PrimaryAkunRef) > 0 Then
                    Urutan = Objakun.GetUrutanMaximalLvl3(PrimaryAkunRef)
                End If
            End If
            Return Me.GenerateAkun(vAkunCode1, vAkunCode2, VAkunDesc, _
            LevelPerkiraan + 1, Me.mPrimary, vUserCode, vTipeAkun2, _
            Objakun.GetTipeAkun(PrimaryAkunRef), Objakun.GetSaldoNormal(PrimaryAkunRef), PrimaryLevel1, _
            PrimaryLevel2, PrimaryLevel3, Urutan + 1, Objakun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GenerateAkun(ByVal vKodeAkun1 As String, ByVal vKodeAkun2 As String, ByVal vAkunDesc As String, ByVal LevelPerkiraan As Integer, _
    ByVal vPrimaryCurrency As Double, ByVal vUserCode As String, ByVal vTipeAkun2 As ClsMasterAkun1.EnumTipeAkun2, _
    ByVal vTipeAkun As ClsMasterAkun1.EnumTipeAkun, ByVal vSaldoNormal As Integer, ByVal vPrimaryLevel1 As Double, ByVal vPrimaryLevel2 As Double, _
    ByVal vPrimaryLevel3 As Double, ByVal vUrutanAkun As Double, ByVal ObjAkun As ClsMasterAkun1) As Double
        Dim spasi As String = "     "

        Try
            With ObjAkun
                'Masukan data master akun detail
                .AccountCode = vKodeAkun1 + vKodeAkun2
                .Description = vAkunDesc
                .LevelPerkiraan = LevelPerkiraan
                If .LevelPerkiraan = 1 Then
                    .Tampilan = .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 2 Then
                    .Tampilan = spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 3 Then
                    .Tampilan = spasi + spasi + .AccountCode + " - " + .Description
                ElseIf .LevelPerkiraan = 4 Then
                    .Tampilan = spasi + spasi + spasi + .AccountCode + " - " + .Description
                End If
                .Aktif = 1
                .Header = 0
                .BeginningBalance = 0
                .BeginningBalanceOriginalCurrency = 0
                .AccountNature = ""
                .PrimaryCurrency = vPrimaryCurrency
                .TipeAkun2 = vTipeAkun2
                .TipePerkiraan = vTipeAkun
                .SaldoNormal = vSaldoNormal
                .PrimaryLv1 = vPrimaryLevel1
                .PrimaryLv2 = vPrimaryLevel2
                .PrimaryLv3 = vPrimaryLevel3
                .Urutan = vUrutanAkun
            End With

            ObjAkun.SimpanBaru(vUserCode)

            Return ObjAkun.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlByPrimaryAcc_RealizedExchangeGainLossAccount(ByVal vPrimaryAccount As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlPrimaryRealizedExchange(vPrimaryAccount).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPrimaryRealizedExchange(vPrimaryAccount).Value
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

    Public Function GetJmlByPrimaryAcc_UnrealizedExchangeGainLossAccount(ByVal vPrimaryAccount As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryUnrealizedExchangeGainLossAccount(vPrimaryAccount).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryUnrealizedExchangeGainLossAccount(vPrimaryAccount).Value
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

    Public Function GetJmlByPrimaryAcc_AccountReceivable(ByVal vPrimaryAccount As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryReceivableAccount(vPrimaryAccount).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryReceivableAccount(vPrimaryAccount).Value
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

    Public Function GetJmlByPrimaryAcc_AccountPayable(ByVal vPrimaryAccount As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryAccountPayable(vPrimaryAccount).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryAccountPayable(vPrimaryAccount).Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currency1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml(vKriteria + "%").Value
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

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCurrencyCode, Me.mDescription, Me.mDefaultPrecision, _
             vUserCode, Now, Now, Me.mPrimaryPayableAccount, Me.mPrimaryReceivableAccount, _
             Me.mPrimaryUnrealizedExchangGainLossAccount, Me.mPrimaryRealizedExchangeGainLossAccount, Me.mPrimary, Me.mTimestamp) > 0 Then
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
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

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

    Public Function ApakahSudahAda(ByVal vCurrencyCode As String) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vCurrencyCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vCurrencyCode).Value > 0 Then
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
        Try
            'tidak boleh hapus bila data yang bersangkutan sudah dirujuk oleh:
            '0. master product
            Dim ObjProd As New clsMasterProduct
            If ObjProd.GetJmlBy_PrimaryCurrency(vPrimary) > 0 Then
                MsgBox("There are product data using this currency, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            ObjProd = Nothing
            '1. customer-item bank
            Dim ObjCust As New clsMasterCustomer
            If ObjCust.GetJmlBy_PrimaryCurrency(vPrimary) > 0 Then
                MsgBox("There are customer's bank data using this currency, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            ObjCust = Nothing
            '2. supplier-item bank & default currency
            Dim ObjSup As New clsMasterSupplier
            If ObjSup.GetJmlBy_PrimaryCurrency(vPrimary) > 0 Then
                MsgBox("There are suppliers data using this currency, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            ObjSup = Nothing

            Dim ObjLS As New clsMasterListedSoftware
            If ObjLS.IsPurchaseInstalled = True Then
                Dim DTA As New DseMasterUnitCekTransaksiTableAdapters.DataTable1TableAdapter
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                '3. transaksi po
                If DTA.ScalarQuery_GetJmlPO(vPrimary).HasValue = True Then
                    If DTA.ScalarQuery_GetJmlPO(vPrimary).Value > 0 Then
                        MsgBox("There are purchase order data using this currency, you cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
                '4. transaksi aop & aop item cost
                If DTA.ScalarQuery_GetJmlAOP(vPrimary).HasValue = True Then
                    If DTA.ScalarQuery_GetJmlAOP(vPrimary).Value > 0 Then
                        MsgBox("There are approvall of purchase data using this currency, you cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
                If DTA.ScalarQuery_GetJmlAOP_ItemCost(vPrimary).HasValue = True Then
                    If DTA.ScalarQuery_GetJmlAOP_ItemCost(vPrimary).Value > 0 Then
                        MsgBox("There are approvall of purchase cost data using this currency, you cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
                '8. transaksi ipo
                If DTA.ScalarQuery_GetJmlIPO(vPrimary).HasValue = True Then
                    If DTA.ScalarQuery_GetJmlIPO(vPrimary).Value > 0 Then
                        MsgBox("There are internal purchase order data using this currency, you cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                DTA.Dispose()
            End If

            Return True
            'MsgBox("belum cek integritas")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Sub IsiCBO(ByVal vCBO As ComboBox, ByVal vTBL As dseMasterCurrency.tblmaster_currencyDataTable)
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_All_Sorted(vTBL)
            vCBO.DataSource = vTBL.DefaultView
            vCBO.DisplayMember = "Currency Code"
            vCBO.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterCurrency.tblmaster_currency1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currency1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

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
                DGV.Columns(1).HeaderText = "Currency code"
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
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Dim TBL As New dseMasterCurrency.tblmaster_currencyDataTable
        Dim row As dseMasterCurrency.tblmaster_currencyRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    Me.mDefaultPrecision = .Default_Precision
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mCurrencyCode = .Currency_Code
                    Me.mTimestamp = .Timestamp

                    Me.mPrimaryPayableAccount = .Primary_Payable_Account
                    Me.mPrimaryRealizedExchangeGainLossAccount = .Primary_Realized_Exchange_Gain_Loss
                    Me.mPrimaryReceivableAccount = .Primary_Receivable_Account
                    Me.mPrimaryUnrealizedExchangGainLossAccount = .Primary_Unrealized_Exchange_Gain_Loss_Account
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

    Public Function GetDecimal_DefaultCurrency_POS() As Integer
        Try
            Return Me.GetDefaultPrecission(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Function

    Public Function GetDecimal_DefaultCurrency() As Integer
        Try
            Return Me.GetDefaultPrecission(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Function

    Public Function GetDefaultExhangeRate(ByVal vPrimaryCur As Double) As Decimal
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDefaultExchangeRate(vPrimaryCur).HasValue = True Then
                Return DTA.ScalarQuery_GetDefaultExchangeRate(vPrimaryCur).Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDefaultPrecission(ByVal vPrimaryCurrency As Double) As Integer
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDefaultPrecision(vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetDefaultPrecision(vPrimaryCurrency)
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

    Public Function GetPrimary_Akun_Unrealize_Exch_Rate(ByVal vPrimaryCur As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunUnrealizedExch(vPrimaryCur).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunUnrealizedExch(vPrimaryCur).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_Akun_Realize_Exch_Rate(ByVal vPrimaryCur As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunRealizedExchRate(vPrimaryCur).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunRealizedExchRate(vPrimaryCur).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunAR(ByVal vPrimaryCur As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunReceivable(vPrimaryCur).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunReceivable(vPrimaryCur).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunAP(ByVal vPrimaryCur As Double) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunPayable(vPrimaryCur).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunPayable(vPrimaryCur).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCurrencyCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

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
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCurrencyCode As String) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCurrencyCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCurrencyCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescription(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDescription(vPrimary)
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
