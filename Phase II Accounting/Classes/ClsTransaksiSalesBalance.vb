Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesBalance
    Private mPrimary As Double = -1
    Private mPrimaryCustomer As Double = -1
    Private mPrimaryAccount As Double = -1
    Private mInvoiceNumber As String = ""
    Private mInvoiceDate As Date = Now
    Private mSales As Double = 0
    Private mVAT As Double = 0
    Private mDutystamp As Double = 0
    Private mTotal As Double = 0
    Private mDueDate As Date = Now
    Private mRecieved As Double = 0
    Private mBalance As Double = 0
    Private mRemarks As String = ""
    Private mPrimaryCurrency As Double = -1
    Private mExchangeRate As Decimal = 1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice
        '2 berikut ini baru ditambahkan di v.1.0.5, berkaitan penggabungan jurnal sales dan jurnal sales return.
        Automatic_SR
        Manual_SR
        '2 berikut ditambahkan ketika update v.1.6.0 untuk iniled
        Credit_Note
        Debet_Note
        Manual_DP 'add aul
    End Enum

    Public Enum EnumTipePayment As Integer
        SalesInvoice_SalesInvoiceDP
        SalesReturn
        AOP_AOPDP
        Purchase_Return
    End Enum

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property PrimaryAccount() As Double
        Get
            Return Me.mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAccount = value
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

    Public Property Primary_Customer() As Double
        Get
            Return mPrimaryCustomer
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomer = value
        End Set
    End Property

    Public Property Invoice_Number() As String
        Get
            Return mInvoiceNumber
        End Get
        Set(ByVal value As String)
            mInvoiceNumber = value
        End Set
    End Property

    Public Property Invoice_Date() As Date
        Get
            Return mInvoiceDate
        End Get
        Set(ByVal value As Date)
            mInvoiceDate = value
        End Set
    End Property

    Public Property Sales() As Double
        Get
            Return mSales
        End Get
        Set(ByVal value As Double)
            mSales = value
        End Set
    End Property

    Public Property VAT() As Double
        Get
            Return mVAT
        End Get
        Set(ByVal value As Double)
            mVAT = value
        End Set
    End Property

    Public Property Dutystamp() As Double
        Get
            Return mDutystamp
        End Get
        Set(ByVal value As Double)
            mDutystamp = value
        End Set
    End Property

    Public Property Total() As Double
        Get
            Return mTotal
        End Get
        Set(ByVal value As Double)
            mTotal = value
        End Set
    End Property

    Public Property DueDate() As Date
        Get
            Return mDueDate
        End Get
        Set(ByVal value As Date)
            mDueDate = value
        End Set
    End Property

    Public Property Recieved() As Double
        Get
            Return mRecieved
        End Get
        Set(ByVal value As Double)
            mRecieved = value
        End Set
    End Property

    Public Property Balance() As Double
        Get
            Return mBalance
        End Get
        Set(ByVal value As Double)
            mBalance = value
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
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If DTA.Insert(Me.Primary_Customer, Me.mInvoiceNumber, Me.mInvoiceDate, Me.mDueDate, _
                          Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryCurrency, Me.mExchangeRate, _
                          Me.mSales, Me.mVAT, Me.mDutystamp, Me.mTotal, Me.mRecieved, Me.mBalance, _
                          Me.mPrimaryAccount) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByInvoice_Number(Me.mInvoiceNumber)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data balance sales
                Me.SimpanDataSalesBalance()

                'nilai balikan
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbl_Caritrans_historicalsalesTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                Return DTA.ScalarQuery_GetJmlByInvoiceNumberLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustName(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If DTA.Update(Me.Primary_Customer, Me.mInvoiceNumber, Me.mInvoiceDate, Me.mDueDate, _
                          Me.mRemarks, vUserCode, Now, Now, Me.mPrimaryCurrency, Me.mExchangeRate, _
                          Me.mSales, Me.mVAT, Me.mDutystamp, Me.mTotal, Me.mRecieved, Me.mBalance, _
                          Me.mPrimaryAccount, Me.mPrimary) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus data balance sales
                Me.HapusDataSalesBalance(Me.mPrimary)
                'simpan data balance sales
                Me.SimpanDataSalesBalance()

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

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimary, vUserCode) = True Then
                'hapus data sales balance
                Me.HapusDataSalesBalance(vPrimary)

                'hapus data transaksi historical
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

    ''
    Public Function ApakahSudahAda(ByVal vInvoiceNumber As String) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByInvoiceNumber(vInvoiceNumber).HasValue Then
                If DTA.ScalarQuery_GetJmlByInvoiceNumber(vInvoiceNumber).Value > 0 Then
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    Optional ByVal vStrictUser As Boolean = True, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim PrimaryBalance As Double = -1
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Jurnal As String = ""

        Try
            'tidak boleh edit jika:
            '1. jika bukan user yang bersangkutan yang melakukan proses edit
            If vStrictUser = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                    If CreatedBy <> vUserCode Then
                        If vTampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '2. jika sudah ada sales receivement jurnal
            Jurnal = ObjJurnal.Get_SalesBalance_Jurnal(vPrimary)
            If Jurnal <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This sales balance referred by sales receivement journal: " + Jurnal + _
                    ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'kalo sampe sini maka boleh edit
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            ObjJurnal = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    Optional ByVal vStrictUser As Boolean = True, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim PrimaryBalance As Double = -1
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Jurnal As String = ""

        Try
            'tidak boleh edit jika:
            '1. jika bukan user yang bersangkutan yang melakukan proses edit
            If vStrictUser = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                    If CreatedBy <> vUserCode Then
                        If vTampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '2. jika sudah ada sales receivement,maka tidak boleh edit
            Jurnal = ObjJurnal.Get_SalesBalance_Jurnal(vPrimary)
            If Jurnal <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This sales balance referred by sales receivement journal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'kalo sampe sini maka boleh edit
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            ObjJurnal = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesBalance.tbl_Caritrans_historicalsalesDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbl_Caritrans_historicalsalesTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                QueryResult = DTA.FillBy_InvoiceNumberLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                QueryResult = DTA.FillBy_CustomerCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                QueryResult = DTA.FillBy_CustomerNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Customer code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Customer name"
                DGV.Columns(3).HeaderText = "Invoice Number"
                ''DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds / 10000) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / 10000, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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

    ''@rosadinnn
    ''@20151022-1008
    Public Function CariDetail(ByVal vTBL As DseTransaksiSalesBalance.tbltrans_historicalsalesDataTable) As Boolean
        Dim row As DseTransaksiSalesBalance.tbltrans_historicalsalesRow
        Try
            row = vTBL.Rows(0)
            With row
                Me.mPrimary = .Primary
                Me.mPrimaryCustomer = .Primary_Customer
                Me.mInvoiceNumber = .Invoice_Number
                Me.Invoice_Date = .Invoice_Date
                Me.mSales = .Sales
                Me.mVAT = .VAT
                Me.mDutystamp = .Dutystamp
                Me.mTotal = .Total
                Me.mDueDate = .Due_Date
                Me.mRecieved = .Recieved
                Me.mBalance = .Balance
                Me.mRemarks = .Remarks
                Me.mCreatedBy = .Created_By
                Me.mCreatedDate = .Created_Date
                Me.mModifiedBy = .Modified_By
                Me.mModifiedDate = .Modified_Date
                Me.mTimestamp = .Timestamp
                Me.mPrimaryCurrency = .primary_currency
                Me.mExchangeRate = .exchange_rate
                Me.mPrimaryAccount = .primary_ar_account
            End With
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TBL As New DseTransaksiSalesBalance.tbltrans_historicalsalesDataTable
        Dim row As DseTransaksiSalesBalance.tbltrans_historicalsalesRow
        Try
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mPrimaryCustomer = .Primary_Customer
                    Me.mInvoiceNumber = .Invoice_Number
                    Me.Invoice_Date = .Invoice_Date
                    Me.mSales = .Sales
                    Me.mVAT = .VAT
                    Me.mDutystamp = .Dutystamp
                    Me.mTotal = .Total
                    Me.mDueDate = .Due_Date
                    Me.mRecieved = .Recieved
                    Me.mBalance = .Balance
                    Me.mRemarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mExchangeRate = .exchange_rate
                    Me.mPrimaryAccount = .primary_ar_account
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

    Public Function GetPrimaryByInvoiceNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter

        Try
            Return DTA.ScalarQuery_GetPrimaryByInvoice_Number(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).Value
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

    Public Function GetInvoiceNumberByPrimary(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetInvoiceNumberByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetInvoiceNumberByPrimary(vPrimary)
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
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_ByPrimaryCustomer(ByVal vPrimaryCustomer As Double) As Double
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimaryByPrimaryCustomer(vPrimaryCustomer).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByInvoiceNumber(ByVal vInvoiceNumber As String) As Double
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByInvoiceNumber(vInvoiceNumber).HasValue Then
                Return DTA.ScalarQuery_GetJmlByInvoiceNumber(vInvoiceNumber).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini menyimpan data sales balance yang muncul di jurnal sales
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SimpanDataSalesBalance()
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = EnumTipeSalesBalance.Manual
        Dim Obj_Kartu As New ClsTransaksi_Kartu_APAR
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecBase As Integer = 0
        Dim PrimaryAkunAR As Double = -1

        Try
            DefPrecBase = ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)
            PrimaryAkunAR = Me.mPrimaryAccount
            ' PrimaryAkunAR = ObjCur.GetPrimaryAkunAR(Me.mPrimaryCurrency)

            DTA.Insert(Me.mPrimaryCustomer, Me.mPrimary, TipeBalance, Me.mInvoiceNumber, Me.mInvoiceDate, Me.mRemarks, Me.mDueDate, _
            Math.Round(Me.mBalance * Me.mExchangeRate, DefPrecBase), 0, 0, Math.Round(Me.mBalance * Me.mExchangeRate, DefPrecBase), _
            0, Me.mCreatedBy, Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, _
            Me.mPrimaryCurrency, Me.mBalance, Me.mExchangeRate, PrimaryAkunAR)

            'simpan data kartu AR
            Obj_Kartu.Simpan(Me.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual, Me.mInvoiceNumber, _
            Me.mPrimaryCustomer, Math.Round(Me.mBalance * Me.mExchangeRate, DefPrecBase), 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            Me.mInvoiceDate, Me.Primary, Me.mInvoiceDate, Math.Abs((Me.mDueDate - Me.mInvoiceDate).Days), _
            ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual, Me.PrimaryCurrency, Me.mBalance, 0, Me.mExchangeRate, PrimaryAkunAR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj_Kartu = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub HapusDataSalesBalance(ByVal vPrimarySalesBalance As Double)
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = EnumTipeSalesBalance.Manual
        Dim obj As New ClsTransaksi_Kartu_APAR

        Try
            DTA.Delete(vPrimarySalesBalance, TipeBalance)
            obj.Hapus(vPrimarySalesBalance, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            obj = Nothing
        End Try
    End Sub

    Public Function GetBalanceOriginal(ByVal vPrimary As Double, ByVal DTA As DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter) As Decimal
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetBalanceOriginal(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetBalanceOriginal(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiSalesBalance.balance_invoice_viewDataTable, _
        ByVal ij As FrmTransaksiJournal_Sales, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimaryCust As Double, ByVal vPrimaryJurnal As Double, _
        ByVal Default_Prec As Integer, ByVal TBLSR As DseTransaksiSalesReturnBalance.balance_sr_viewDataTable)
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.balance_invoice_viewTableAdapter
        Dim TBLDummy As New DseTransaksiSalesBalance.balance_invoice_viewDataTable
        Dim DTAJurnal As New DseTransaksiSalesBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBLJurnal As New DseTransaksiSalesBalance.tbltrans_jurnal_item_paymentDataTable
        Dim rowJurnal As DseTransaksiSalesBalance.tbltrans_jurnal_item_paymentRow
        Dim Filter As String = ""
        Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim row, row1 As DseTransaksiSalesBalance.balance_invoice_viewRow
        Dim rowdgv As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ketemu As Boolean = False

        Dim TipeBalance1 As EnumTipeSalesBalance = EnumTipeSalesBalance.Manual_SR
        Dim TipeBalance2 As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic_SR
        Dim TipeBalance3 As EnumTipeSalesBalance = EnumTipeSalesBalance.Credit_Note
        Dim TipeBalance4 As EnumTipeSalesBalance = EnumTipeSalesBalance.Debet_Note
        Dim ObjSRBalance As New ClsTransaksiSalesReturnBalance
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim ObjCDNote As New ClsTransaksiCreditNote
        Dim AdaNonIDR As Boolean = False
        Dim objAkun As New ClsMasterAkun

        'Update SMESS 2026
        Dim DTACustPO As New DseTransaksiSalesBalanceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLCustPO As New DseTransaksiSalesBalance.tbltrans_salesinvoiceDataTable
        Dim DTAReturnCustPO As New DseTransaksiSalesBalanceTableAdapters.tblReturnCustPONoTableAdapter
        Dim TBLReturnCustPO As New DseTransaksiSalesBalance.tblReturnCustPONoDataTable
        Try
            'isikan data dari database
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()

            If vPrimaryCust = -1 Then
                DTA.FillBy_All(TBL)
            Else
                DTA.FillBy_PrimaryCust(TBL, vPrimaryCust)
            End If

            'Update SMESS 2026
            DTACustPO.Connection.ConnectionString = Core_Module.GetConstring
            DTACustPO.FillBy_PrimaryCust(TBLCustPO, vPrimaryCust)
            DTAReturnCustPO.Connection.ConnectionString = Core_Module.GetConstring
            DTAReturnCustPO.FillBy_PrimaryCust(TBLReturnCustPO, vPrimaryCust)

            'modifikasi nilai received dan receivement berdasarkan data jurnal
            'bila proses edit
            If vPrimaryJurnal <> -1 Then
                DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)

                For Each rowJurnal In TBLJurnal.Rows
                    TBLDummy.Rows.Clear()
                    DTA.FillBy_DocumenAcuan(TBLDummy, vPrimaryCust, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan)

                    If TBLDummy.Rows.Count > 0 Then
                        For Each row In TBLDummy.Rows
                            If rowJurnal.primary_currency = row.primary_currency And rowJurnal.primary_akun_ap_ar = row.primary_akun_AR Then

                                If row.Tipe = TipeBalance1 Then
                                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                                    ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                                End If
                                If row.Tipe = TipeBalance2 Then
                                    ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
                                End If
                                If row.Tipe = TipeBalance3 OrElse row.Tipe = TipeBalance4 Then
                                    'ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                                    ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice, row.Due_Date)
                                End If

                                Filter = "Primary = " + row.Primary.ToString

                                row.Received -= rowJurnal.Payment
                                row.Receivement = rowJurnal.Payment
                                row.received_original -= rowJurnal.payment_original
                                row.Receivement_Original = rowJurnal.payment_original
                                If ij.ObjJurnal.ApakahSudahRevaluasi(ij.dtpDate.Value) = True Then
                                    row.exchange_rate = rowJurnal.exchange_rate
                                End If

                                TBL.DefaultView.RowFilter = Filter
                                If TBL.DefaultView.Count = 0 Then
                                    row1 = TBL.Newbalance_invoice_viewRow
                                    With row1
                                        .Balance = row.Balance
                                        .Due_Date = row.Due_Date
                                        .Duty_stamp = row.Duty_stamp
                                        .Invoice_Date = row.Invoice_Date
                                        .Invoice_Number = row.Invoice_Number
                                        .Primary = row.Primary
                                        .Primary_Customer = row.Primary_Customer
                                        .Primary_Invoice = row.Primary_Invoice
                                        .Received = row.Received
                                        .Receivement = row.Receivement
                                        .Remarks = row.Remarks
                                        .Sales = row.Sales
                                        .Sales_VAT = row.Sales_VAT
                                        .Tipe = row.Tipe
                                        .primary_currency = row.primary_currency
                                        .Default_Precision = row.Default_Precision
                                        .Currency_Code = row.Currency_Code
                                        .received_original = row.received_original
                                        .Receivement_Original = row.Receivement_Original
                                        .exchange_rate = row.exchange_rate
                                        .primary_akun_AR = row.primary_akun_AR
                                        .balance_original = row.balance_original
                                    End With
                                    TBL.Rows.Add(row1)
                                Else
                                    row1 = TBL.DefaultView(0).Row
                                    row1.Received = row.Received
                                    row1.Receivement = row.Receivement
                                    row1.received_original = row.received_original
                                    row1.Receivement_Original = row.Receivement_Original
                                    row1.exchange_rate = row.exchange_rate
                                End If
                            End If
                        Next
                    End If
                Next
            End If

            'khusus untuk sales return maka isikan no transaksi invoice, no sr, dan remarksnya
            For Each row In TBL.Rows
                If row.Tipe = TipeBalance1 Then
                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                    ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                End If
                If row.Tipe = TipeBalance2 Then
                    ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
                End If

                If row.Tipe = TipeBalance3 OrElse row.Tipe = TipeBalance4 Then
                    'ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                    ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice, row.Due_Date)
                End If
            Next

            'isikan ke dalam datagridview
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Invoice Date] , [Invoice Number], primary"

            For Each dv As DataRowView In TBL.DefaultView
                row = dv.Row
                'Update SMESS 2026
                If row.Tipe <> TipeBalance1 And row.Tipe <> TipeBalance2 And row.Tipe <> TipeBalance3 And row.Tipe <> TipeBalance4 Then
                    row.Customer_PO_No = Me.isiCustPONo(TBLCustPO, row.Primary_Invoice)
                ElseIf row.Tipe = TipeBalance2 Then
                    row.Customer_PO_No = Me.isiReturnCustPONo(TBLReturnCustPO, row.Primary_Invoice)
                End If

                If row.Balance <> 0 Or row.Receivement <> 0 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowdgv
                        If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                            AdaNonIDR = True
                        End If
                        If row.Sales = 0 Then
                            row.Invoice_Number = row.Invoice_Number + " (Tax + Dutystamp)"
                        End If
                        .Cells(ij.ColInvoiceNo.Index).Value = row.Invoice_Number
                        .Cells(ij.ColInvoiceDate.Index).Value = row.Invoice_Date
                        .Cells(ij.ColSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales, Default_Prec)
                        .Cells(ij.ColVatSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales_VAT, Default_Prec)
                        .Cells(ij.ColDutystamp.Index).Value = ObjCD.DigitGrouping(row.Duty_stamp, Default_Prec)
                        .Cells(ij.ColTotalSales_VAT_Value.Index).Value = ObjCD.DigitGrouping(row.Sales + row.Sales_VAT + row.Duty_stamp, Default_Prec)
                        .Cells(ij.ColReceived.Index).Value = ObjCD.DigitGrouping(row.Received, Default_Prec)
                        .Cells(ij.ColBalanceSales.Index).Value = ObjCD.DigitGrouping(row.Balance, Default_Prec)
                        .Cells(ij.ColReceivement.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Receivement_Original * row.exchange_rate, Default_Prec), Default_Prec)
                        .Cells(ij.ColPrimarySalesBalance.Index).Value = row.Primary
                        .Cells(ij.ColDueDateInvoice.Index).Value = row.Due_Date
                        .Cells(ij.ColPrimaryDocumenAcuan_Sales.Index).Value = row.Primary_Invoice
                        .Cells(ij.ColTipeDocumentAcuan_Sales.Index).Value = row.Tipe
                        .Cells(ij.ColInvRemarks.Index).Value = row.Remarks
                        .Cells(ij.ColPayInv.Index).Value = "Full Payment"
                        .Cells(ij.ColPrimaryCust.Index).Value = row.Primary_Customer
                        .Cells(ij.ColCurrencyBalance.Index).Value = row.Currency_Code
                        .Cells(ij.ColCurrencyDefPrec.Index).Value = row.Default_Precision
                        .Cells(ij.ColPrimaryCurrencyBalance.Index).Value = row.primary_currency
                        .Cells(ij.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                        .Cells(ij.ColReceivementOriginal.Index).Value = ObjCD.DigitGrouping(row.Receivement_Original, row.Default_Precision)
                        .Cells(ij.ColReceivedOriginal.Index).Value = ObjCD.DigitGrouping(row.received_original, row.Default_Precision)
                        .Cells(ij.ColExhangeRateBalance.Index).Value = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                        .Cells(ij.ColBalance1.Index).Value = CDec(.Cells(ij.ColBalanceSales.Index).Value) + _
                        Math.Round(CDec(row.Receivement_Original * row.exchange_rate), Default_Prec) + _
                        CDec(.Cells(ij.ColReceived.Index).Value)
                        ' nilai diatas ini sebenarnya sama dengan received, 
                        'hanya saja menjadi salah untuk journal yang sudah direaluasi kemudian belum dibetulkan, diganti untuk mengakomodasi proses pembetulan
                        .Cells(ij.ColBalanceOriginal02.Index).Value = .Cells(ij.ColBalanceOriginal.Index).Value
                        .Cells(ij.ColBalanceOriginal01.Index).Value = CDec(.Cells(ij.ColBalanceOriginal.Index).Value) + _
                        CDec(.Cells(ij.ColReceivementOriginal.Index).Value) + CDec(.Cells(ij.ColReceivedOriginal.Index).Value)
                        .Cells(ij.ColPrimaryAkunBalance.Index).Value = row.primary_akun_AR
                        .Cells(ij.ColARAccount.Index).Value = objAkun.GetAccountCode(row.primary_akun_AR) & " - " & objAkun.GetDescription_ByPrimaryAcc(row.primary_akun_AR)

                        'Update SMESS 2026
                        .Cells(ij.ColPONo.Index).Value = row.Customer_PO_No
                    End With
                End If
            Next

            'jika semua baris currencynya IDR, maka exchage rate tidak perlu ada
            If AdaNonIDR = True Then
                ij.ColExhangeRateBalance.Visible = True
            Else
                ij.ColExhangeRateBalance.Visible = False
            End If
            ij.ColInvoiceDate.DefaultCellStyle.Format = "dd MMM yyyy"
            ij.ColDueDateInvoice.DefaultCellStyle.Format = "dd MMM yyyy"

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTAJurnal.Dispose()
            TBLJurnal.Dispose()
        End Try
    End Sub

    'Update SMESS 2026
    Public Function isiCustPONo(ByVal vTBLCustPO As DseTransaksiSalesBalance.tbltrans_salesinvoiceDataTable, ByVal vPrimaryInvoice As Double) As String
        Dim retval As String = ""
        Dim vTemp As String = ""
        Dim vMaxLength As Integer = 150
        Dim vSuffix As String = " ..."
        Dim row As DseTransaksiSalesBalance.tbltrans_salesinvoiceRow
        Dim i As Integer = 0
        Try
            vTBLCustPO.DefaultView.RowFilter = "[Primary]=" & vPrimaryInvoice
            If vTBLCustPO.DefaultView.Count > 0 Then
                For i = 0 To vTBLCustPO.DefaultView.Count - 1
                    row = vTBLCustPO.DefaultView(i).Row
                    If String.IsNullOrEmpty(row.Cust_Po_No) = False Then
                        vTemp = row.Cust_Po_No
                    ElseIf String.IsNullOrEmpty(row.Cust_PO_No_DP) = False Then
                        vTemp = row.Cust_PO_No_DP
                    Else
                        vTemp = ""
                    End If

                    If vTemp <> "" Then
                        If String.IsNullOrEmpty(retval) Then
                            vTemp = vTemp
                        Else
                            vTemp = retval & ", " & vTemp
                        End If

                        If vTemp.Length > 150 Then
                            retval = vTemp.Substring(0, vMaxLength) & vSuffix
                            Exit For
                        Else
                            retval = vTemp
                        End If
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        Finally
            vTBLCustPO.DefaultView.RowFilter = Nothing
        End Try
    End Function

    'Update SMESS 2026
    Public Function isiReturnCustPONo(ByVal vTBLReturnCustPO As DseTransaksiSalesBalance.tblReturnCustPONoDataTable, ByVal vPrimaryInvoice As Double) As String
        Dim retval As String = ""
        Dim vTemp As String = ""
        Dim vMaxLength As Integer = 150
        Dim vSuffix As String = " ..."
        Dim row As DseTransaksiSalesBalance.tblReturnCustPONoRow
        Dim i As Integer = 0
        Try
            vTBLReturnCustPO.DefaultView.RowFilter = "[Primary]=" & vPrimaryInvoice
            If vTBLReturnCustPO.DefaultView.Count > 0 Then
                For i = 0 To vTBLReturnCustPO.DefaultView.Count - 1
                    row = vTBLReturnCustPO.DefaultView(i).Row
                    If String.IsNullOrEmpty(row.Cust_PO_No) = False Then
                        If String.IsNullOrEmpty(retval) Then
                            vTemp = row.Cust_PO_No
                        Else
                            vTemp = retval & ", " & row.Cust_PO_No
                        End If

                        If vTemp.Length > 150 Then
                            retval = vTemp.Substring(0, vMaxLength) & vSuffix
                            Exit For
                        Else
                            retval = vTemp
                        End If
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retval
        Finally
            vTBLReturnCustPO.DefaultView.RowFilter = Nothing
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan form sales journal untuk mengisikan daftar balance invoice
    ''' </summary>
    ''' <param name="vDGV"></param>
    ''' <remarks></remarks>
    'Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiSalesBalance.balance_invoice_viewDataTable, _
    'ByVal ij As FrmTransaksiJournal_Sales, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimaryCust As Double, ByVal vPrimaryJurnal As Double, _
    'ByVal Default_Prec As Integer, ByVal TBLSR As DseTransaksiSalesReturnBalance.balance_sr_viewDataTable)
    '    Dim DTA As New DseTransaksiSalesBalanceTableAdapters.balance_invoice_viewTableAdapter
    '    Dim TBLDummy As New DseTransaksiSalesBalance.balance_invoice_viewDataTable
    '    Dim DTAJurnal As New DseTransaksiSalesBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '    Dim TBLJurnal As New DseTransaksiSalesBalance.tbltrans_jurnal_item_paymentDataTable
    '    Dim rowJurnal As DseTransaksiSalesBalance.tbltrans_jurnal_item_paymentRow
    '    Dim Filter As String = ""
    '    Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesInvoice_SalesInvoiceDP
    '    Dim row, row1 As DseTransaksiSalesBalance.balance_invoice_viewRow
    '    Dim rowdgv As DataGridViewRow
    '    Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '    Dim ketemu As Boolean = False

    '    Dim TipeBalance1 As EnumTipeSalesBalance = EnumTipeSalesBalance.Manual_SR
    '    Dim TipeBalance2 As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic_SR
    '    Dim ObjSRBalance As New ClsTransaksiSalesReturnBalance
    '    Dim ObjSR As New ClsTransaksiSalesReturn
    '    Dim AdaNonIDR As Boolean = False

    '    Try
    '        'isikan data dari database
    '        DTA.ClearBeforeFill = False
    '        TBL.Rows.Clear()

    '        If vPrimaryCust = -1 Then
    '            DTA.FillBy_All(TBL)
    '        Else
    '            DTA.FillBy_PrimaryCust(TBL, vPrimaryCust)
    '        End If

    '        'modifikasi nilai received dan receivement berdasarkan data jurnal
    '        'bila proses edit
    '        If vPrimaryJurnal <> -1 Then
    '            DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)

    '            For Each rowJurnal In TBLJurnal.Rows
    '                TBLDummy.Rows.Clear()
    '                DTA.FillBy_DocumenAcuan(TBLDummy, vPrimaryCust, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan)

    '                If TBLDummy.Rows.Count > 0 Then
    '                    For Each row In TBLDummy.Rows
    '                        If rowJurnal.primary_currency = row.primary_currency And rowJurnal.primary_akun_ap_ar = row.primary_akun_AR Then

    '                            If row.Tipe = TipeBalance1 Then
    '                                'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
    '                                ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
    '                            End If
    '                            If row.Tipe = TipeBalance2 Then
    '                                ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
    '                            End If

    '                            Filter = "Primary = " + row.Primary.ToString

    '                            row.Received -= rowJurnal.Payment
    '                            row.Receivement = rowJurnal.Payment
    '                            row.received_original -= rowJurnal.payment_original
    '                            row.Receivement_Original = rowJurnal.payment_original
    '                            If ij.ObjJurnal.ApakahSudahRevaluasi(ij.dtpDate.Value) = True Then
    '                                row.exchange_rate = rowJurnal.exchange_rate
    '                            End If

    '                            TBL.DefaultView.RowFilter = Filter
    '                            If TBL.DefaultView.Count = 0 Then
    '                                row1 = TBL.Newbalance_invoice_viewRow
    '                                With row1
    '                                    .Balance = row.Balance
    '                                    .Due_Date = row.Due_Date
    '                                    .Duty_stamp = row.Duty_stamp
    '                                    .Invoice_Date = row.Invoice_Date
    '                                    .Invoice_Number = row.Invoice_Number
    '                                    .Primary = row.Primary
    '                                    .Primary_Customer = row.Primary_Customer
    '                                    .Primary_Invoice = row.Primary_Invoice
    '                                    .Received = row.Received
    '                                    .Receivement = row.Receivement
    '                                    .Remarks = row.Remarks
    '                                    .Sales = row.Sales
    '                                    .Sales_VAT = row.Sales_VAT
    '                                    .Tipe = row.Tipe
    '                                    .primary_currency = row.primary_currency
    '                                    .Default_Precision = row.Default_Precision
    '                                    .Currency_Code = row.Currency_Code
    '                                    .received_original = row.received_original
    '                                    .Receivement_Original = row.Receivement_Original
    '                                    .exchange_rate = row.exchange_rate
    '                                    .primary_akun_AR = row.primary_akun_AR
    '                                    .balance_original = row.balance_original
    '                                End With
    '                                TBL.Rows.Add(row1)
    '                            Else
    '                                row1 = TBL.DefaultView(0).Row
    '                                row1.Received = row.Received
    '                                row1.Receivement = row.Receivement
    '                                row1.received_original = row.received_original
    '                                row1.Receivement_Original = row.Receivement_Original
    '                                row1.exchange_rate = row.exchange_rate
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            Next
    '        End If

    '        'khusus untuk sales return maka isikan no transaksi invoice, no sr, dan remarksnya
    '        For Each row In TBL.Rows
    '            If row.Tipe = TipeBalance1 Then
    '                'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
    '                ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
    '            End If
    '            If row.Tipe = TipeBalance2 Then
    '                ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
    '            End If
    '        Next

    '        'isikan ke dalam datagridview
    '        vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        vDGV.Rows.Clear()
    '        TBL.DefaultView.RowFilter = ""
    '        TBL.DefaultView.Sort = "[Invoice Date] , [Invoice Number], primary"

    '        For Each dv As DataRowView In TBL.DefaultView
    '            row = dv.Row
    '            If row.Balance <> 0 Or row.Receivement <> 0 Then
    '                vDGV.Rows.Add()
    '                rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
    '                With rowdgv
    '                    If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
    '                        AdaNonIDR = True
    '                    End If
    '                    If row.Sales = 0 Then
    '                        row.Invoice_Number = row.Invoice_Number + " (Tax + Dutystamp)"
    '                    End If
    '                    .Cells(ij.ColInvoiceNo.Index).Value = row.Invoice_Number
    '                    .Cells(ij.ColInvoiceDate.Index).Value = Format(row.Invoice_Date, "dd MMM yyyy")
    '                    .Cells(ij.ColSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales, Default_Prec)
    '                    .Cells(ij.ColVatSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales_VAT, Default_Prec)
    '                    .Cells(ij.ColDutystamp.Index).Value = ObjCD.DigitGrouping(row.Duty_stamp, Default_Prec)
    '                    .Cells(ij.ColTotalSales_VAT_Value.Index).Value = ObjCD.DigitGrouping(row.Sales + row.Sales_VAT + row.Duty_stamp, Default_Prec)
    '                    .Cells(ij.ColReceived.Index).Value = ObjCD.DigitGrouping(row.Received, Default_Prec)
    '                    .Cells(ij.ColBalanceSales.Index).Value = ObjCD.DigitGrouping(row.Balance, Default_Prec)
    '                    .Cells(ij.ColReceivement.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Receivement_Original * row.exchange_rate, Default_Prec), Default_Prec)
    '                    .Cells(ij.ColPrimarySalesBalance.Index).Value = row.Primary
    '                    .Cells(ij.ColDueDateInvoice.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
    '                    .Cells(ij.ColPrimaryDocumenAcuan_Sales.Index).Value = row.Primary_Invoice
    '                    .Cells(ij.ColTipeDocumentAcuan_Sales.Index).Value = row.Tipe
    '                    .Cells(ij.ColInvRemarks.Index).Value = row.Remarks
    '                    .Cells(ij.ColPayInv.Index).Value = "Full Payment"
    '                    .Cells(ij.ColPrimaryCust.Index).Value = row.Primary_Customer
    '                    .Cells(ij.ColCurrencyBalance.Index).Value = row.Currency_Code
    '                    .Cells(ij.ColCurrencyDefPrec.Index).Value = row.Default_Precision
    '                    .Cells(ij.ColPrimaryCurrencyBalance.Index).Value = row.primary_currency
    '                    .Cells(ij.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
    '                    .Cells(ij.ColReceivementOriginal.Index).Value = ObjCD.DigitGrouping(row.Receivement_Original, row.Default_Precision)
    '                    .Cells(ij.ColReceivedOriginal.Index).Value = ObjCD.DigitGrouping(row.received_original, row.Default_Precision)
    '                    .Cells(ij.ColExhangeRateBalance.Index).Value = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
    '                    .Cells(ij.ColBalance1.Index).Value = CDec(.Cells(ij.ColBalanceSales.Index).Value) + _
    '                    Math.Round(CDec(row.Receivement_Original * row.exchange_rate), Default_Prec) + _
    '                    CDec(.Cells(ij.ColReceived.Index).Value)
    '                    ' nilai diatas ini sebenarnya sama dengan received, 
    '                    'hanya saja menjadi salah untuk journal yang sudah direaluasi kemudian belum dibetulkan, diganti untuk mengakomodasi proses pembetulan
    '                    .Cells(ij.ColBalanceOriginal02.Index).Value = .Cells(ij.ColBalanceOriginal.Index).Value
    '                    .Cells(ij.ColBalanceOriginal01.Index).Value = CDec(.Cells(ij.ColBalanceOriginal.Index).Value) + _
    '                    CDec(.Cells(ij.ColReceivementOriginal.Index).Value) + CDec(.Cells(ij.ColReceivedOriginal.Index).Value)
    '                    .Cells(ij.ColPrimaryAkunBalance.Index).Value = row.primary_akun_AR
    '                End With
    '            End If
    '        Next

    '        'jika semua baris currencynya IDR, maka exchage rate tidak perlu ada
    '        If AdaNonIDR = True Then
    '            ij.ColExhangeRateBalance.Visible = True
    '        Else
    '            ij.ColExhangeRateBalance.Visible = False
    '        End If

    '        vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    '        vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '        DTAJurnal.Dispose()
    '        TBLJurnal.Dispose()
    '    End Try
    'End Sub

    Public Sub GetStartingBalance(ByVal vConstringLama As String)
        Try
            'buat laporan ar aging

            'simpan datanya ke dalam form sales balance baru

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakLaporanARDueDate(ByVal vIsDueDate As Boolean, ByVal vDueDate As Date, _
           ByVal vIsAllCust As Boolean, ByVal vPrimaryCust As Double, ByVal TBL As DseTransaksiSalesBalance.balance_invoice_viewDataTable, _
           ByVal vDef_Prec As Integer, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.balance_invoice_viewTableAdapter
        Dim Filter As String = ""
        Dim row As DseTransaksiSalesBalance.balance_invoice_viewRow
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeBalance2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
        Dim TipeBalance1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim ObjSRBalance As New ClsTransaksiSalesReturnBalance

        Dim dsCetak As New DseTransaksiPurchaseBalance
        Dim tblCetak As DseTransaksiPurchaseBalance.APAR_PrintDataTable = dsCetak.APAR_Print
        Dim rowCetak As DseTransaksiPurchaseBalance.APAR_PrintRow
        Dim tblCurrency As New DseTransaksiPurchaseBalance.ListOfCurrencyDataTable
        Dim rowCurrency As DseTransaksiPurchaseBalance.ListOfCurrencyRow
        Dim reportFile As New CrpLaporanReminder

        Dim paramCurrencyCode As String = ""
        Dim paramBalanceOriginal As String = ""
        Dim paramBalanceIdr As String = ""
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim dtaCust As New DseTransaksiPurchaseBalanceTableAdapters.tblmaster_customerTableAdapter
        Dim tblCust As New DseTransaksiPurchaseBalance.tblmaster_customerDataTable
        Dim rowCust As DseTransaksiPurchaseBalance.tblmaster_customerRow

        Try
            ''set def value
            vDef_Prec = ObjCur.GetDecimal_DefaultCurrency
            dtaCust.Fill(tblCust)

            If vIsAllCust = True Then
                DTA.FillBy_All(TBL)
            Else
                DTA.FillBy_PrimaryCust(TBL, vPrimaryCust)
            End If

            'khusus untuk sales return maka isikan no transaksi invoice, no sr, dan remarksnya
            For Each row In TBL.Rows
                If row.Tipe = TipeBalance1 Then
                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                    ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                End If
                If row.Tipe = TipeBalance2 Then
                    ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
                End If
            Next

            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Invoice Date] , [Invoice Number], primary"

            For Each dv As DataRowView In TBL.DefaultView
                row = dv.Row
                If row.Balance <> 0 Or row.Receivement <> 0 Then

                    ''batasan due date
                    If vIsDueDate = True Then
                        If Not row.Due_Date > vDueDate Then

                            ''simpan komponen
                            With row
                                rowCetak = tblCetak.NewAPAR_PrintRow
                                rowCetak.Primary = .Primary
                                rowCetak.Primary_Supplier_Customer = .Primary_Customer
                                tblCust.DefaultView.RowFilter = "Primary =" & .Primary_Customer.ToString
                                If tblCust.DefaultView.Count > 0 Then
                                    rowCust = tblCust.DefaultView(0).Row
                                    rowCetak.Supplier_Customer_Code = rowCust.Customer_Code
                                    rowCetak.Supplier_Customer_Desc = rowCust.Customer_Name
                                Else
                                    rowCetak.Supplier_Customer_Code = "N/A"
                                    rowCetak.Supplier_Customer_Desc = "N/A"
                                End If
                                rowCetak.Primary_Currency = .primary_currency
                                rowCetak.Currency_Code = .Currency_Code
                                rowCetak.Currency_Desc = "n/a"

                                If row.Sales = 0 Then
                                    row.Invoice_Number = row.Invoice_Number + " (Tax + Dutystamp)"
                                End If
                                rowCetak.Invoice_No = row.Invoice_Number
                                rowCetak.Invoice_Date = row.Invoice_Date
                                rowCetak.Due_Date = .Due_Date
                                rowCetak.Balance_Original = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                                rowCetak.Exchange_Rate = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                                rowCetak.Balance_IDR = CDec(ObjCD.DigitGrouping(row.Balance, vDef_Prec)) + _
                                        Math.Round(CDec(row.Receivement_Original * row.exchange_rate), vDef_Prec) + _
                                        CDec(ObjCD.DigitGrouping(row.Received, vDef_Prec))
                                rowCetak.Precision = .Default_Precision
                                tblCetak.Rows.Add(rowCetak)

                                ''update list of currency
                                tblCurrency.DefaultView.RowFilter = "[Currency Code] = '" & row.Currency_Code & "'"
                                If tblCurrency.DefaultView.Count = 0 Then
                                    rowCurrency = tblCurrency.NewListOfCurrencyRow
                                    rowCurrency.Primary = 99999999
                                    rowCurrency.Currency_Code = rowCetak.Currency_Code
                                    rowCurrency.Currency_Desc = "n/a"
                                    If row.Currency_Code = "IDR" Then
                                        rowCurrency.Balance_Original = 0
                                    Else
                                        rowCurrency.Balance_Original = rowCetak.Balance_Original
                                    End If
                                    rowCurrency.Balance_IDR = rowCetak.Balance_IDR
                                    rowCurrency.Precision = rowCetak.Precision
                                    tblCurrency.Rows.Add(rowCurrency)
                                Else
                                    rowCurrency = tblCurrency.DefaultView(0).Row
                                    If row.Currency_Code <> "IDR" Then
                                        rowCurrency.Balance_Original += rowCetak.Balance_Original
                                    End If
                                    rowCurrency.Balance_IDR += rowCetak.Balance_IDR
                                End If
                            End With
                        End If
                    Else

                        ''tanpa batasan due date
                        ''simpan komponen
                        With row
                            rowCetak = tblCetak.NewAPAR_PrintRow
                            rowCetak.Primary = .Primary
                            rowCetak.Primary_Supplier_Customer = .Primary_Customer
                            tblCust.DefaultView.RowFilter = "Primary =" & .Primary_Customer.ToString
                            If tblCust.DefaultView.Count > 0 Then
                                rowCust = tblCust.DefaultView(0).Row
                                rowCetak.Supplier_Customer_Code = rowCust.Customer_Code
                                rowCetak.Supplier_Customer_Desc = rowCust.Customer_Name
                            Else
                                rowCetak.Supplier_Customer_Code = "N/A"
                                rowCetak.Supplier_Customer_Desc = "N/A"
                            End If
                            rowCetak.Primary_Currency = .primary_currency
                            rowCetak.Currency_Code = .Currency_Code
                            rowCetak.Currency_Desc = "n/a"

                            If row.Sales = 0 Then
                                row.Invoice_Number = row.Invoice_Number + " (Tax + Dutystamp)"
                            End If
                            rowCetak.Invoice_No = row.Invoice_Number
                            rowCetak.Invoice_Date = row.Invoice_Date
                            rowCetak.Due_Date = .Due_Date
                            rowCetak.Balance_Original = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                            rowCetak.Exchange_Rate = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                            rowCetak.Balance_IDR = CDec(ObjCD.DigitGrouping(row.Balance, vDef_Prec)) + _
                                    Math.Round(CDec(row.Receivement_Original * row.exchange_rate), vDef_Prec) + _
                                    CDec(ObjCD.DigitGrouping(row.Received, vDef_Prec))
                            rowCetak.Precision = .Default_Precision
                            tblCetak.Rows.Add(rowCetak)

                            ''update list of currency
                            tblCurrency.DefaultView.RowFilter = "[Currency Code] = '" & row.Currency_Code & "'"
                            If tblCurrency.DefaultView.Count = 0 Then
                                rowCurrency = tblCurrency.NewListOfCurrencyRow
                                rowCurrency.Primary = 99999999
                                rowCurrency.Currency_Code = rowCetak.Currency_Code
                                rowCurrency.Currency_Desc = "n/a"
                                If row.Currency_Code = "IDR" Then
                                    rowCurrency.Balance_Original = 0
                                Else
                                    rowCurrency.Balance_Original = rowCetak.Balance_Original
                                End If
                                rowCurrency.Balance_IDR = rowCetak.Balance_IDR
                                rowCurrency.Precision = rowCetak.Precision
                                tblCurrency.Rows.Add(rowCurrency)
                            Else
                                rowCurrency = tblCurrency.DefaultView(0).Row
                                If row.Currency_Code <> "IDR" Then
                                    rowCurrency.Balance_Original += rowCetak.Balance_Original
                                End If
                                rowCurrency.Balance_IDR += rowCetak.Balance_IDR
                            End If
                        End With

                    End If
                End If
            Next

            ''stop if no data
            If TBL.Rows.Count <= 0 Then
                MsgBox("No data available", MsgBoxStyle.Information, "Information")
                Return False
                Exit Function
            End If

            ''pecah list of currecy
            tblCurrency.DefaultView.RowFilter = ""
            tblCurrency.DefaultView.Sort = "[Currency Code]"
            For Each drv As DataRowView In tblCurrency.DefaultView
                rowCurrency = drv.Row
                If paramCurrencyCode.Length = 0 Then
                    paramCurrencyCode = rowCurrency.Currency_Code
                    If rowCurrency.Currency_Code = "IDR" Then
                        paramBalanceOriginal = "-"
                    Else
                        paramBalanceOriginal = Me.ObjCD.DigitGrouping(rowCurrency.Balance_Original.ToString, rowCurrency.Precision)
                    End If
                    paramBalanceIdr = Me.ObjCD.DigitGrouping(rowCurrency.Balance_IDR.ToString, vDef_Prec)
                Else
                    paramCurrencyCode = paramCurrencyCode & vbNewLine & rowCurrency.Currency_Code
                    If rowCurrency.Currency_Code = "IDR" Then
                        paramBalanceOriginal = paramBalanceOriginal & vbNewLine & "-"
                    Else
                        paramBalanceOriginal = paramBalanceOriginal & vbNewLine & Me.ObjCD.DigitGrouping(rowCurrency.Balance_Original.ToString, rowCurrency.Precision)
                    End If
                    paramBalanceIdr = paramBalanceIdr & vbNewLine & Me.ObjCD.DigitGrouping(rowCurrency.Balance_IDR.ToString, vDef_Prec)
                End If
            Next

            reportFile.SetDataSource(dsCetak)

            reportFile.SetParameterValue("TransactionType", "AR")
            If vIsDueDate = True Then
                reportFile.SetParameterValue("TitleOfReport", "ACCOUNT RECEIVEABLE DUE UNTIL  " & Format(vDueDate, "dd MMMM yyyy"))
            Else
                reportFile.SetParameterValue("TitleOfReport", "ACCOUNT RECEIVEABLE DUE")
            End If
            reportFile.SetParameterValue("DueDate", vDueDate)
            reportFile.SetParameterValue("TitleOfCurrency", paramCurrencyCode)
            reportFile.SetParameterValue("TotalBalanceOriginal", paramBalanceOriginal)
            reportFile.SetParameterValue("TotalBalanceIDR", paramBalanceIdr)
            reportFile.SetParameterValue("Precision", vDef_Prec)
            reportFile.SetParameterValue("PayableOrReceiveable", "Receiveable")
            reportFile.SetParameterValue("SupplierOrCustomer", "CUSTOMER")
            reportFile.SetParameterValue("DefPrecExchangeRate", Core_Module.DefPrecExchangeRate)
            reportFile.SetParameterValue(reportFile.Parameter_TipeDokumen.ParameterFieldName, "INVOICE")

            vCrv.ReportSource = reportFile
            vCrv.Zoom(90)
            vCrv.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjSR = Nothing
            ObjSRBalance = Nothing
        End Try
    End Function
End Class
