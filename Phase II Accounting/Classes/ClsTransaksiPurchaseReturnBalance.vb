Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturnBalance
    Private mPrimary As Double = -1
    Private mPrimarysupp As Double = -1
    Private mPrimaryAccount As Double = -1
    Private mAOPRNumber As String = ""
    Private mAOPRDate As Date = Now
    Private mPurchaseReturn As Double = 0
    Private mVAT As Double = 0
    Private mDutystamp As Double = 0
    Private mTotal As Double = 0
    Private mDueDate As Date = Now
    Private mPaid As Double = 0
    Private mBalance As Double = 0
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mPrimaryCurrency As Double = -1
    Dim mExchangeRate As Decimal = 1

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
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

    Public Property Primary_Supplier() As Double
        Get
            Return mPrimarysupp
        End Get
        Set(ByVal value As Double)
            mPrimarysupp = value
        End Set
    End Property

    Public Property PrimaryAccount() As Double
        Get
            Return mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAccount = value
        End Set
    End Property

    Public Property AOPRNumber() As String
        Get
            Return mAOPRNumber
        End Get
        Set(ByVal value As String)
            mAOPRNumber = value
        End Set
    End Property

    Public Property AOPRDate() As Date
        Get
            Return mAOPRDate
        End Get
        Set(ByVal value As Date)
            mAOPRDate = value
        End Set
    End Property

    Public Property Purchase_Return() As Double
        Get
            Return mPurchaseReturn
        End Get
        Set(ByVal value As Double)
            mPurchaseReturn = value
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

    Public Property Paid() As Double
        Get
            Return mPaid
        End Get
        Set(ByVal value As Double)
            mPaid = value
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
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If DTA.Insert(Me.Primary_Supplier, Me.mAOPRNumber, Me.mAOPRDate, Me.Purchase_Return, _
                          Me.VAT, Me.Dutystamp, Me.Total, Me.DueDate, Me.Paid, Me.Balance, _
                          Me.Remarks, vUserCode, Now, vUserCode, Now, Now, _
                          Me.mPrimaryCurrency, Me.mExchangeRate, Me.mPrimaryAccount) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByAOPRNumber(Me.mAOPRNumber)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data balance
                Me.SimpanDataPurchaseReturnBalance()

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

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tblcaritrans_historicalpurchasereturnbalancemanualviewTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByAOPRNumber Then
                Return DTA.ScalarQuery_GetJmlByAOPRNumberLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.BySupplierCode Then
                Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.BySupplierName Then
                Return DTA.ScalarQuery_GetJmlBySuppNameLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If DTA.Update(Me.mPrimarysupp, Me.mAOPRNumber, Me.mAOPRDate, Me.mDueDate, _
                          Me.mRemarks, vUserCode, Now, Now, Me.mPrimaryCurrency, Me.mExchangeRate, _
                          Me.mPrimaryAccount, Me.mPrimary) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus data balance lama
                Me.HapusDataPurchaseReturnBalance(Me.mPrimary)

                'simpan data balance baru
                Me.SimpanDataPurchaseReturnBalance()

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
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimary, vUserCode) = True Then
                'hapus data balance
                Me.HapusDataPurchaseReturnBalance(vPrimary)

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

    Public Function ApakahSudahAda(ByVal vInvoiceNumber As String) As Boolean
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByAOPRNumber(vInvoiceNumber).HasValue Then
                If DTA.ScalarQuery_GetJmlByAOPRNumber(vInvoiceNumber).Value > 0 Then
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
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi sales journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            If vPrimary <> -1 Then
                Journal_Ref_Number = Obj.Get_PurchaseReturnBalance_Jurnal(vPrimary)
                If Journal_Ref_Number <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + Journal_Ref_Number + _
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
            Obj = Nothing
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    Optional ByVal vStrictUser As Boolean = True, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi sales journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            If vPrimary <> -1 Then
                Journal_Ref_Number = Obj.Get_PurchaseReturnBalance_Jurnal(vPrimary)
                If Journal_Ref_Number <> "" Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction referred by: " + Journal_Ref_Number + _
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
            Obj = Nothing
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaskiPurchaseReturnBalance.tblcaritrans_historicalpurchasereturnbalancemanualviewDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tblcaritrans_historicalpurchasereturnbalancemanualviewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAOPRNumber Then
                QueryResult = DTA.FillBy_AOPRNumberLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierCode Then
                QueryResult = DTA.FillBy_SuppCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SuppNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Supplier code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Supplier name"
                DGV.Columns(3).HeaderText = "PR Number"
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

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBL As New DseTransaskiPurchaseReturnBalance.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim row As DseTransaskiPurchaseReturnBalance.tbltrans_historicalpurchasereturnbalancemanualRow

        Try
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mPrimarysupp = .Primary_Supplier
                    Me.mAOPRNumber = .AOPR_Number
                    Me.mAOPRDate = .AOPR_Date
                    Me.mPurchaseReturn = .Purchase_Return
                    Me.mVAT = .VAT
                    Me.Dutystamp = .Dutystamp
                    Me.mTotal = .Total
                    Me.mDueDate = .Due_Date
                    Me.mPaid = .Paid
                    Me.mBalance = .Balance
                    Me.mRemarks = .Remarks
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mExchangeRate = .Exchange_Rate
                    Me.mPrimaryAccount = .primary_ap_account
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

    Public Function GetPrimaryByAOPRNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimaryByAOPRNumber(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetNoTransaksi_Remarks_Tanggal(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vDate As Date, ByVal vPrimary As Double, _
    ByRef vSuppInvNumber As String)
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBL As New DseTransaskiPurchaseReturnBalance.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim row As DseTransaskiPurchaseReturnBalance.tbltrans_historicalpurchasereturnbalancemanualRow

        Try
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.Remarks
                vNoTransaksi = row.AOPR_Number
                vDate = row.AOPR_Date
                vSuppInvNumber = ""
            Else
                vRemarks = ""
                vNoTransaksi = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetAOPRNumberByPrimary(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAOPRNumberByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAOPRNumberByPrimary(vPrimary)
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
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
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

    Public Function GetPrimary_ByPrimarySupplier(ByVal vPrimaryCustomer As Double) As Double
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimaryByPrimarySupp(vPrimaryCustomer).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByAOPRNumber(ByVal vInvoiceNumber As String) As Double
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByAOPRNumber(vInvoiceNumber).HasValue Then
                Return DTA.ScalarQuery_GetJmlByAOPRNumber(vInvoiceNumber).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SimpanDataPurchaseReturnBalance()
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim Obj As New ClsTransaksi_Kartu_APAR
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecBase As Integer = 0
        Dim BalanceBase As Decimal = 0
        Dim PrimaryAkun As Double = -1

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            BalanceBase = Math.Round(Me.mBalance * Me.mExchangeRate * -1, DefPrecBase, MidpointRounding.AwayFromZero)
            PrimaryAkun = Me.mPrimaryAccount
            '  PrimaryAkun = ObjCur.GetPrimaryAkunAP(Me.mPrimaryCurrency)

            DTA.Insert(Me.mPrimarysupp, Me.mPrimary, TipeBalance, Me.mAOPRNumber, Me.mAOPRDate, Me.mRemarks, Me.mDueDate, _
            BalanceBase, 0, 0, BalanceBase, 0, Me.mCreatedBy, _
            Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, Me.mPrimaryCurrency, Me.mBalance * -1, _
            Me.mExchangeRate, PrimaryAkun)

            'simpan data kartu pr
            BalanceBase = Math.Abs(BalanceBase)
            Obj.Simpan(Me.mPrimary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual, Me.mAOPRNumber, _
            Me.mPrimarysupp, BalanceBase, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, Me.mAOPRDate, _
            Me.mPrimary, Me.mAOPRDate, Math.Abs((Me.mDueDate - Me.mAOPRDate).Days), ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual, _
            Me.mPrimaryCurrency, Math.Abs(Me.mBalance), 0, Me.mExchangeRate, PrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub HapusDataPurchaseReturnBalance(ByVal vPrimaryPurchaseReturnBalance As Double)
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            DTA.Delete(vPrimaryPurchaseReturnBalance, TipeBalance)
            Obj.Hapus(vPrimaryPurchaseReturnBalance, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaskiPurchaseReturnBalance.balance_purchase_return_viewDataTable, _
   ByVal ij As FrmTransaksiJournal_Purchase, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimarySupp As Double, _
   ByVal vPrimaryJurnal As Double, ByVal vDef_Prec As Integer)
        Dim DTA As New DseTransaskiPurchaseReturnBalanceTableAdapters.balance_purchase_return_viewTableAdapter
        Dim TBLDummy As New DseTransaskiPurchaseReturnBalance.balance_purchase_return_viewDataTable
        Dim DTAJurnal As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBLJurnal As New DseTransaskiPurchaseReturnBalance.tbltrans_jurnal_item_paymentDataTable
        Dim rowJurnal As DseTransaskiPurchaseReturnBalance.tbltrans_jurnal_item_paymentRow
        Dim Filter As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.Purchase_Return
        Dim row, row1 As DseTransaskiPurchaseReturnBalance.balance_purchase_return_viewRow
        Dim rowdgv As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            'isikan data dari database
            DTA.FillBy_PrimarySupp(TBL, vPrimarySupp)

            'modifikasi nilai received dan receivement berdasarkan data jurnal
            'bila proses edit
            If vPrimaryJurnal <> -1 Then
                DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)
                For Each rowJurnal In TBLJurnal.Rows
                    TBLDummy.Rows.Clear()
                    DTA.FillBy_DocumentAcuan(TBLDummy, vPrimarySupp, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan)
                    If TBLDummy.Rows.Count > 0 Then
                        row = TBLDummy.Rows(0)
                        Filter = "Primary = " + row.Primary.ToString
                        row.Received -= rowJurnal.Payment
                        row.Receivement = rowJurnal.Payment
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count = 0 Then
                            row1 = TBL.Newbalance_purchase_return_viewRow
                            With row1
                                .Balance = row.Balance
                                .Due_Date = row.Due_Date
                                .Duty_stamp = row.Duty_stamp
                                .AOPR_Date = row.AOPR_Date
                                .AOPR_Number = row.AOPR_Number
                                .Primary = row.Primary
                                .Primary_Supplier = row.Primary_Supplier
                                .Primary_AOPR = row.Primary_AOPR
                                .Received = row.Received
                                .Receivement = row.Receivement
                                .Remarks = row.Remarks
                                .Purchase_Return = row.Purchase_Return
                                .Purchase_Return_VAT = row.Purchase_Return_VAT
                                .Tipe = row.Tipe
                            End With
                            TBL.Rows.Add(row1)
                        Else
                            row1 = TBL.DefaultView(0).Row
                            row1.Received = row.Received
                            row1.Receivement = row.Receivement
                        End If
                    End If
                Next
            End If

            'isikan ke dalam datagridview
            vDGV.Rows.Clear()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "AOPR Number"

            For i As Double = 0 To TBL.DefaultView.Count - 1
                row = TBL.DefaultView(i).Row
                If row.Balance <> 0 Or row.Receivement <> 0 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowdgv
                        .Cells(ij.ColPurchaseReturnNumber.Index).Value = row.AOPR_Number
                        .Cells(ij.ColAOPRDate.Index).Value = Format(row.AOPR_Date, "dd MMM yyyy")
                        .Cells(ij.ColPurchaseReturnValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase_Return, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColVATReturnValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase_Return_VAT, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColDutystampPR.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColPurchaseReturn_VATReturn_Value.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase_Return + row.Purchase_Return_VAT + row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColReceivedReturn.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Received, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColBalanceReturn.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Balance, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColReturnReceivement.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Receivement, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColPrimaryPurchaseReturnBalance.Index).Value = row.Primary
                        .Cells(ij.ColDueDatePurchaseReturn.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                        .Cells(ij.Col_TipeDocumentAcuanPR.Index).Value = row.Tipe
                        .Cells(ij.Col_PrimaryDocumenAcuan_PR.Index).Value = row.Primary_AOPR
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub
End Class
