Option Explicit On
Option Strict Off

Public Class ClsTransaksiBalanceManual_DP
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mTipeDP As EnumTipeDP = EnumTipeDP.DP_AOP
    Dim mPrimaryDocAcuan As Double = -1
    Dim mPrimaryPihakAcuan As Double = -1
    Dim mDPNo As String = ""
    Dim mInvNo As String = ""
    Dim mDPDate As Date = Now
    Dim mDueDate As Date = Now
    'Dim mPurchaseOrSales As Decimal = 0
    'Dim mVAT As Decimal = 0
    'Dim mDutystamp As Decimal = 0
    'Dim mTotal As Decimal = 0
    'Dim mPaidOrReceived As Decimal = 0
    Dim mBalance As Decimal = 0
    Dim mPrimaryCurrency As Double = -1
    Dim mExchangeRate As Decimal = 1
    Dim mPrimaryAccount As Double = -1
    Dim mPrimaryAccountDP As Double = -1
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumTipeDP As Integer
        DP_AOP
        DP_Invoice
    End Enum

    Public Property TipeDP() As EnumTipeDP
        Get
            Return Me.mTipeDP
        End Get
        Set(ByVal value As EnumTipeDP)
            mTipeDP = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryDocAcuan() As Double
        Get
            Return Me.mPrimaryDocAcuan
        End Get
        Set(ByVal value As Double)
            mPrimaryDocAcuan = value
        End Set
    End Property

    Public Property PrimaryPihakAcuan() As Double
        Get
            Return Me.mPrimaryPihakAcuan
        End Get
        Set(ByVal value As Double)
            mPrimaryPihakAcuan = value
        End Set
    End Property

    Public Property DPNo() As String
        Get
            Return Me.mDPNo
        End Get
        Set(ByVal value As String)
            mDPNo = value
        End Set
    End Property

    Public Property InvNo() As String
        Get
            Return Me.mInvNo
        End Get
        Set(ByVal value As String)
            mInvNo = value
        End Set
    End Property

    Public Property DPDate() As Date
        Get
            Return mDPDate
        End Get
        Set(ByVal value As Date)
            mDPDate = value
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

    'Public Property PurchaseOrSales() As Decimal
    '    Get
    '        Return mPurchaseOrSales
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mPurchaseOrSales = value
    '    End Set
    'End Property

    'Public Property VAT() As Decimal
    '    Get
    '        Return mVAT
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mVAT = value
    '    End Set
    'End Property

    'Public Property Dutystamp() As Decimal
    '    Get
    '        Return mDutystamp
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mDutystamp = value
    '    End Set
    'End Property

    'Public Property Total() As Decimal
    '    Get
    '        Return mTotal
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mTotal = value
    '    End Set
    'End Property

    'Public Property PaidOrReceived() As Decimal
    '    Get
    '        Return mPaidOrReceived
    '    End Get
    '    Set(ByVal value As Decimal)
    '        mPaidOrReceived = value
    '    End Set
    'End Property

    Public Property Balance() As Decimal
        Get
            Return mBalance
        End Get
        Set(ByVal value As Decimal)
            mBalance = value
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

    Public Property ExchangeRate() As Double
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Double)
            mExchangeRate = value
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

    Public Property PrimaryAccountDP() As Double
        Get
            Return Me.mPrimaryAccountDP
        End Get
        Set(ByVal value As Double)
            mPrimaryAccountDP = value
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

    Public Function ApakahSudahAda(ByVal vDPNo As String, ByVal vTipeDP As EnumTipeDP) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            Retval = DTA.ScalarQuery_GetJml(vDPNo, vTipeDP)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vTipeDP As EnumTipeDP, Optional ByVal vStrictUser As Boolean = True, _
    Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim CreatedBy As String = ""
        Dim Journal_Ref_Number As String = ""
        Dim TransaksiDP_No As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi sales journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            If vPrimary <> -1 Then
                Journal_Ref_Number = Me.GetBalance_Jurnal(vPrimary, vTipeDP)
                If Journal_Ref_Number <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + Journal_Ref_Number + _
                        ". You cannot delete this data", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If

                TransaksiDP_No = Me.ApaSudahDigunakan_Transaksi(vPrimary, vTipeDP)
                If TransaksiDP_No <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + TransaksiDP_No + _
                        ". You cannot delete this data", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. jika bukan user yang bersangkutan yang melakukan proses edit
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

            'kalo sampe sini maka boleh edit
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApaSudahDigunakan_Transaksi(ByVal vPrimaryDP As Double, ByVal vTipe As EnumTipeDP) As String
        Dim dta As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_aopTableAdapter
        Dim tbl As New DseTransaksiBalanceManual_DP.tbltrans_aopDataTable
        Dim row As DseTransaksiBalanceManual_DP.tbltrans_aopRow
        Dim dta1 As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim tbl1 As New DseTransaksiBalanceManual_DP.tbltrans_salesinvoiceDataTable
        Dim row1 As DseTransaksiBalanceManual_DP.tbltrans_salesinvoiceRow
        Dim retval As String = ""

        Try
            If vTipe = EnumTipeDP.DP_AOP Then
                dta.Connection.ConnectionString = core_module.getconstring
                dta.Fill(tbl, vPrimaryDP)
                For Each row In tbl.Rows
                    If retval = "" Then
                        retval = row.Aop_No
                    Else
                        retval = retval + ", " + row.Aop_No
                    End If
                Next
            Else
                dta1.Connection.ConnectionString = core_module.getconstring
                dta1.Fill(tbl1, vPrimaryDP)
                For Each row1 In tbl1.Rows
                    If retval = "" Then
                        retval = row1.Invoice_No
                    Else
                        retval = retval + ", " + row1.Invoice_No
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetBalance_Jurnal(ByVal vPrimaryPurchaseBalance As Double, ByVal vTipeDP As EnumTipeDP) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiJournal.DataTable2DataTable
        Dim row As DseTransaksiJournal.DataTable2Row
        Dim retval As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim TipeSalesBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP

        Try
            If vTipeDP = EnumTipeDP.DP_Invoice Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
            End If

            DTA.Fill(TBL, TipePayment, vPrimaryPurchaseBalance, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vTipeDP As EnumTipeDP, Optional ByVal vStrictUser As Boolean = True, _
    Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim CreatedBy As String = ""
        Dim Journal_Ref_Number As String = ""
        Dim TransaksiDP_No As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi sales journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            If vPrimary <> -1 Then
                Journal_Ref_Number = Me.GetBalance_Jurnal(vPrimary, vTipeDP)
                If Journal_Ref_Number <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + Journal_Ref_Number + _
                        ". You cannot edit this data", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If

                TransaksiDP_No = Me.ApaSudahDigunakan_Transaksi(vPrimary, vTipeDP)
                If TransaksiDP_No <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + TransaksiDP_No + _
                        ". You cannot edit this data", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. jika bukan user yang bersangkutan yang melakukan proses edit
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

            'kalo sampe sini maka boleh edit
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TBL As New DseTransaksiBalanceManual_DP.tbltrans_historicalbalancemanual_dpDataTable
        Dim row As DseTransaksiBalanceManual_DP.tbltrans_historicalbalancemanual_dpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mTipeDP = .Tipe_DP
                    Me.mPrimaryDocAcuan = .Primary_Doc_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDPNo = .DP_No
                    Me.mInvNo = .Invoice_No
                    Me.mDPDate = .DP_Date
                    Me.mDueDate = .Due_Date
                    Me.mBalance = .Balance
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mExchangeRate = .exchange_rate
                    Me.mPrimaryAccount = .primary_ap_ar_account
                    Me.mPrimaryAccountDP = .primary_dp_account
                    Me.mRemarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
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

    Public Function ApakahDocAcuan_SudahDigunakan(ByVal vPrimaryDocAcuan As Double, ByVal vPrimaryTransaksi As Double, ByVal vTipeDP As EnumTipeDP) As Boolean
        Dim dta As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Dp_No As String = ""
        Dim DocAcuan_No As String = ""

        Try
            'cek apakah document acuan sudah di gunakan di transaksi historical lain
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dp_No = dta.ScalarQuery_GetDPNo_Jml(vPrimaryDocAcuan, vTipeDP, vPrimaryTransaksi)
            If String.IsNullOrEmpty(Dp_No) = False Then
                If vTipeDP = EnumTipeDP.DP_AOP Then
                    Dim ObjPO As New ClsTransaksiPO
                    DocAcuan_No = ObjPO.GetNoTransaksiPO(vPrimaryDocAcuan)
                    ObjPO = Nothing
                Else
                    Dim ObjSO As New ClsTransaksiSalesOrder
                    DocAcuan_No = ObjSO.GetNoSO(vPrimaryDocAcuan)
                    ObjSO = Nothing
                End If

                MsgBox("This " & DocAcuan_No & " referred by historical balance: " & Dp_No, MsgBoxStyle.Information, "Information")

                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
    ByVal vEnumTipeDP As EnumTipeDP) As Double
        Dim DTAAop As New DseTransaksiBalanceManual_DPTableAdapters.tbl_aop_dp_findTableAdapter
        Dim DTAInv As New DseTransaksiBalanceManual_DPTableAdapters.tbl_inv_dp_findTableAdapter
        Dim QueryResult As Double = 0

        Try
            If vEnumTipeDP = EnumTipeDP.DP_AOP Then
                DTAAop.Connection.ConnectionString = Core_Module.GetConstring
                If vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                    QueryResult = DTAAop.ScalarQuery_GetJml_AOP_DP(vEnumTipeDP, vKriteria + "%", "%", "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                    QueryResult = DTAAop.ScalarQuery_GetJml_AOP_DP(vEnumTipeDP, "%", vKriteria + "%", "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                    QueryResult = DTAAop.ScalarQuery_GetJml_AOP_DP(vEnumTipeDP, "%", "%", vKriteria + "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.BySupplierInvoiceNo Then
                    QueryResult = DTAAop.ScalarQuery_GetJml_AOP_DP(vEnumTipeDP, "%", "%", "%", vKriteria + "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTAAop.ScalarQuery_GetJml_AOP_DP(vEnumTipeDP, "%", "%", "%", "%", vKriteria + "%").Value
                End If
            Else
                DTAInv.Connection.ConnectionString = Core_Module.GetConstring
                If vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                    QueryResult = DTAInv.ScalarQuery_GetJml_INV_DP(vEnumTipeDP, vKriteria + "%", "%", "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                    QueryResult = DTAInv.ScalarQuery_GetJml_INV_DP(vEnumTipeDP, "%", vKriteria + "%", "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByInvoiceNumber Then
                    QueryResult = DTAInv.ScalarQuery_GetJml_INV_DP(vEnumTipeDP, "%", "%", vKriteria + "%", "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.BySupplierInvoiceNo Then
                    QueryResult = DTAInv.ScalarQuery_GetJml_INV_DP(vEnumTipeDP, "%", "%", "%", vKriteria + "%", "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTAInv.ScalarQuery_GetJml_INV_DP(vEnumTipeDP, "%", "%", "%", "%", vKriteria + "%").Value
                End If
            End If

            Return QueryResult
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTAAop.Dispose()
            DTAInv.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vKriteria As String, ByVal DGV As DataGridView, ByVal vTBLAop_DP As DseTransaksiBalanceManual_DP.tbl_aop_dp_findDataTable, _
    ByVal vTBLInv_DP As DseTransaksiBalanceManual_DP.tbl_inv_dp_findDataTable, ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    ByVal vEnumTipeDP As EnumTipeDP) As Boolean
        Dim DTAAop As New DseTransaksiBalanceManual_DPTableAdapters.tbl_aop_dp_findTableAdapter
        Dim DTAInv As New DseTransaksiBalanceManual_DPTableAdapters.tbl_inv_dp_findTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTAAop.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'vTanggalDari = Format(CDate(vTanggalDari), "yyyy-MM-dd")
            'vTanggalSampai = Format(CDate(vTanggalSampai), "yyyy-MM-dd")

            If vEnumTipeDP = EnumTipeDP.DP_AOP Then
                DTAAop.Connection.ConnectionString = Core_Module.GetConstring
                If SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                    QueryResult = DTAAop.Fill(vTBLAop_DP, vEnumTipeDP, vKriteria + "%", "%", "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                    QueryResult = DTAAop.Fill(vTBLAop_DP, vEnumTipeDP, "%", vKriteria + "%", "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                    QueryResult = DTAAop.Fill(vTBLAop_DP, vEnumTipeDP, "%", "%", vKriteria + "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierInvoiceNo Then
                    QueryResult = DTAAop.Fill(vTBLAop_DP, vEnumTipeDP, "%", "%", "%", vKriteria + "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTAAop.Fill(vTBLAop_DP, vEnumTipeDP, "%", "%", "%", "%", vKriteria + "%")
                End If
                DGV.DataSource = vTBLAop_DP.DefaultView
            Else
                DTAInv.Connection.ConnectionString = Core_Module.GetConstring
                If SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                    QueryResult = DTAInv.Fill(vTBLInv_DP, vEnumTipeDP, vKriteria + "%", "%", "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                    QueryResult = DTAInv.Fill(vTBLInv_DP, vEnumTipeDP, "%", vKriteria + "%", "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByInvoiceNumber Then
                    QueryResult = DTAInv.Fill(vTBLInv_DP, vEnumTipeDP, "%", "%", vKriteria + "%", "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierInvoiceNo Then
                    QueryResult = DTAInv.Fill(vTBLInv_DP, vEnumTipeDP, "%", "%", "%", vKriteria + "%", "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTAInv.Fill(vTBLInv_DP, vEnumTipeDP, "%", "%", "%", "%", vKriteria + "%")
                End If
                DGV.DataSource = vTBLInv_DP.DefaultView
            End If

            'update dgv
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                'DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                If vEnumTipeDP = EnumTipeDP.DP_Invoice Then
                    DGV.Columns(4).Visible = False
                End If
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
            DTAAop.Dispose()
            DTAInv.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(Me.mTipeDP, Me.mPrimaryDocAcuan, Me.mPrimaryPihakAcuan, Me.mDPNo, _
                          Me.mInvNo, Me.mDPDate, Me.mDueDate, Me.Balance, Me.mPrimaryCurrency, _
                          Me.mExchangeRate, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.Remarks, _
                          vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mDPNo, Me.mTipeDP)
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data balance
                Me.SimpanDataBalanceDP(Me.mTipeDP)

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
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Try
            If DTA.Update(Me.mTipeDP, Me.mPrimaryDocAcuan, Me.mPrimaryPihakAcuan, Me.mDPNo, _
                          Me.mInvNo, Me.mDPDate, Me.mDueDate, Me.Balance, Me.mPrimaryCurrency, _
                          Me.mExchangeRate, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.Remarks, _
                          vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus data balance lama
                Me.HapusDataBalanceDP(Me.mPrimary, Me.mTipeDP)

                'simpan data balance baru
                Me.SimpanDataBalanceDP(Me.mTipeDP)

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

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vTipeDP As EnumTipeDP) As Boolean
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vTipeDP) = True Then
                'hapus data balance
                Me.HapusDataBalanceDP(vPrimary, vTipeDP)

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

    Private Sub SimpanDataBalanceDP(ByVal vTipeDP As EnumTipeDP)
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
        Dim JenisTransaksiAwal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual
        Dim JenisKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
        Dim Obj As New ClsTransaksi_Kartu_APAR
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecBase As Integer = 0
        Dim BalanceBase As Decimal = 0
        Dim PrimaryAkun As Double = -1
        Dim DebitValue As Decimal = 0
        Dim DebitValueOriginal As Decimal = 0
        Dim KreditValue As Decimal = 0
        Dim KreditValueOriginal As Decimal = 0

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            BalanceBase = Math.Round(Me.mBalance * Me.mExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero)
            PrimaryAkun = Me.mPrimaryAccount
            '  PrimaryAkun = ObjCur.GetPrimaryAkunAP(Me.mPrimaryCurrency)

            If vTipeDP = EnumTipeDP.DP_AOP Then
                Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
                DTA.Insert(Me.mPrimaryPihakAcuan, Me.mPrimary, TipeBalance, Me.mDPNo, Me.mDPDate, Me.mRemarks, Me.mDueDate, _
                BalanceBase, 0, 0, BalanceBase, 0, Me.mCreatedBy, Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, _
                Me.mPrimaryCurrency, Me.mBalance, Me.mExchangeRate, PrimaryAkun)
                DTA.Dispose()

                KreditValue = BalanceBase
                KreditValueOriginal = Me.mBalance
            Else
                JenisTransaksiAwal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual
                JenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR

                Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
                DTA.Insert(Me.mPrimaryPihakAcuan, Me.mPrimary, TipeBalance, Me.mDPNo, Me.mDPDate, Me.mRemarks, Me.mDueDate, _
                BalanceBase, 0, 0, BalanceBase, 0, Me.mCreatedBy, Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, _
                Me.mPrimaryCurrency, Me.mBalance, Me.mExchangeRate, PrimaryAkun)
                DTA.Dispose()

                DebitValue = BalanceBase
                DebitValueOriginal = Me.mBalance
            End If

            'simpan data kartu pr
            BalanceBase = Math.Abs(BalanceBase)
            Obj.Simpan(Me.mPrimary, JenisTransaksiAwal, Me.mDPNo, Me.mPrimaryPihakAcuan, DebitValue, KreditValue, "-", JenisKartu, Me.mDPDate, _
            Me.mPrimary, Me.mDPDate, Math.Abs((Me.mDueDate - Me.mDPDate).Days), JenisTransaksiAwal, Me.mPrimaryCurrency, Math.Abs(DebitValueOriginal), _
            Math.Abs(KreditValueOriginal), Me.mExchangeRate, PrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub HapusDataBalanceDP(ByVal vPrimaryTransaksiBalance As Double, ByVal vTipeDP As EnumTipeDP)
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
        Dim JenisTransaksiAwal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            If vTipeDP = EnumTipeDP.DP_AOP Then
                Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
                DTA.Delete(vPrimaryTransaksiBalance, TipeBalance)
                DTA.Dispose()
            Else
                Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
                JenisTransaksiAwal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual
                DTA.Delete(vPrimaryTransaksiBalance, TipeBalance)
                DTA.Dispose()
            End If

            Obj.Hapus(vPrimaryTransaksiBalance, JenisTransaksiAwal)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Public Function GetSuppInvoiceNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetSuppInvNo(vPrimary)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
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

    Public Function GetPrimary(ByVal vDPNo As String, ByVal vTipe As EnumTipeDP) As Double
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimary(vDPNo, vTipe)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetPrimaryCurrency(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryCurr(vPrimaryTransaksi)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetPrimaryAPAR_Acc(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryAPARAcc(vPrimaryTransaksi)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetDPNo(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim Retval As String

        Try
            Retval = DTA.ScalarQuery_GetDPNo(vPrimaryTransaksi)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
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
        Dim DTA As New DseTransaksiBalanceManual_DPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return CDate("#1/1/1950#")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return CDate("#1/1/1950#")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
