Option Explicit On
Option Strict Off

Public Class clsMasterCurrency
    Dim mPrimary As Double = -1
    Dim mCurrencyCode As String = ""
    Dim mDescription As String = ""
    Dim mDefaultPrecision As Integer = -1

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

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Try
            If DTA.Insert(Me.mCurrencyCode, Me.mDescription, Me.mDefaultPrecision, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCurrencyCode).Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currency1TableAdapter

        Try
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
            If DTA.Update(Me.mCurrencyCode, Me.mDescription, Me.mDefaultPrecision, _
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
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

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

    Public Function ApakahSudahAda(ByVal vCurrencyCode As String) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
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

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterCurrency.tblmaster_currency1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currency1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
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
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter
        Dim TBL As New dseMasterCurrency.tblmaster_currencyDataTable
        Dim row As dseMasterCurrency.tblmaster_currencyRow

        Try
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

    Public Function GetCode_DefaultCurrency() As String
        Try
            Return Me.GetCurrencyCode(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
            Return ""
        End Try
    End Function

    Public Function GetDefaultPrecission(ByVal vPrimaryCurrency As Double) As Integer
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
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

    Public Function GetCurrencyCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCategoryCode As String) As Double
        Dim DTA As New dseMasterCurrencyTableAdapters.tblmaster_currencyTableAdapter

        Try
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
End Class
