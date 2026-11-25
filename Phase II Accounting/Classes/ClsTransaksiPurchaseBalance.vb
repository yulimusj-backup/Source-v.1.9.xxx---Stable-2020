Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseBalance
    Private mPrimary As Double = -1
    Private mPrimarysupp As Double = -1
    Private mPrimaryAP As Double = -1
    Private mAOPNumber As String = ""
    Private mAOPDate As Date = Now
    Private mPurchase As Double = 0
    Private mVAT As Double = 0
    Private mTotal As Double = 0
    Private mDueDate As Date = Now
    Private mPaid As Double = 0
    Private mBalance As Double = 0
    Private mRemarks As String = ""
    Private mDutystamp As Double = 0
    Private mPrimaryCurrency As Double = 0
    Private mExchangeRate As Decimal = 1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim objCD As New Core_Module.ClsCekData

    Enum EnumCetakanInvoiceNo As Integer
        SupplierInv
        AOPNo
    End Enum

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property PrimaryAP() As Double
        Get
            Return Me.mPrimaryAP
        End Get
        Set(ByVal value As Double)
            mPrimaryAP = value
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

    Public Property DutyStamp() As Double
        Get
            Return Me.mDutystamp
        End Get
        Set(ByVal value As Double)
            mDutystamp = value
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

    Public Property AOPNumber() As String
        Get
            Return mAOPNumber
        End Get
        Set(ByVal value As String)
            mAOPNumber = value
        End Set
    End Property

    Public Property AOPDate() As Date
        Get
            Return mAOPDate
        End Get
        Set(ByVal value As Date)
            mAOPDate = value
        End Set
    End Property

    Public Property Purchase() As Double
        Get
            Return mPurchase
        End Get
        Set(ByVal value As Double)
            mPurchase = value
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter

        Try
            If DTA.Insert(Me.Primary_Supplier, Me.mAOPNumber, Me.mAOPDate, Me.mDueDate, _
                          Me.Remarks, vUserCode, Now, vUserCode, Now, Now, _
                          Me.mPrimaryCurrency, Me.mExchangeRate, Me.mDutystamp, _
                          Me.mPurchase, Me.mVAT, Me.mTotal, Me.mPaid, Me.mBalance, Me.mPrimaryAP) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByAOPNumber(Me.mAOPNumber)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data balance
                Me.SimpanDataPurchaseBalance()

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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tblcaritrans_historicalpurchasebalancemanualviewTableAdapter
        Try
            If vSearchBy = EnumSearchCriteria.ByAOPNumber Then
                Return DTA.ScalarQuery_GetJmlByAOPNumberLIKE(vKriteria + "%").Value
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Try
            If DTA.Update(Me.mPrimarysupp, Me.mAOPNumber, Me.mAOPDate, Me.mDueDate, _
                          Me.mRemarks, vUserCode, Now, Now, Me.mPrimaryCurrency, _
                          Me.mExchangeRate, Me.mDutystamp, Me.mPurchase, Me.mVAT, _
                          Me.mTotal, Me.mPaid, Me.mBalance, Me.mPrimaryAP, Me.mPrimary) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus data balance lama
                Me.HapusDataPurchaseBalance(Me.mPrimary)
                'simpan data balance
                Me.SimpanDataPurchaseBalance()

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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimary, vUserCode) = True Then
                'hapus data purchase balance
                Me.HapusDataPurchaseBalance(vPrimary)

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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByAOPNumber(vInvoiceNumber).HasValue Then
                If DTA.ScalarQuery_GetJmlByAOPNumber(vInvoiceNumber).Value > 0 Then
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi sales journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            If vPrimary <> -1 Then
                Journal_Ref_Number = Obj.Get_PurchaseBalance_Jurnal(vPrimary)
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Journal_Ref_Number As String = ""

        Try
            'tidak boleh edit jika:
            '1. sudah ada transaksi purchase journal yang merujuk ke transaksi ini, sebutkan nomor transaksinya
            Journal_Ref_Number = Obj.Get_PurchaseBalance_Jurnal(vPrimary)
            If Journal_Ref_Number <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This transaction referred by: " + Journal_Ref_Number + _
                    ". You cannot edit this data", MsgBoxStyle.Information, "Information")
                End If
                Return False
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
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPurchaseBalance.tblcaritrans_historicalpurchasebalancemanualviewDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tblcaritrans_historicalpurchasebalancemanualviewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAOPNumber Then
                QueryResult = DTA.FillBy_AOPNumberLIKE(TBL, Criteria + "%")
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
                DGV.Columns(3).HeaderText = "AOP Number"
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
    Public Function CariDetail(ByVal vTBL As DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualDataTable, _
            ByVal Primary As Double) As Boolean
        Dim row As DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualRow
        Try
            vTBL.DefaultView.RowFilter = "[Primary] = " & Primary.ToString

            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                With row
                    Me.mPrimary = .Primary
                    Me.mPrimarysupp = .Primary_Supplier
                    Me.mAOPNumber = .AOP_Number
                    Me.mAOPDate = .AOP_Date
                    Me.mPurchase = .Purchase
                    Me.mVAT = .VAT
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
                    Me.mDutystamp = .Dutystamp
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mExchangeRate = .exchange_rate
                    Me.mPrimaryAP = .primary_ap_account
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBL.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim TBL As New DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualDataTable
        Dim row As DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualRow
        Try
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mPrimarysupp = .Primary_Supplier
                    Me.mAOPNumber = .AOP_Number
                    Me.mAOPDate = .AOP_Date
                    Me.mPurchase = .Purchase
                    Me.mVAT = .VAT
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
                    Me.mDutystamp = .Dutystamp
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mExchangeRate = .exchange_rate
                    Me.mPrimaryAP = .primary_ap_account
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

    Public Function GetPrimaryByAOPNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter

        Try
            Return DTA.ScalarQuery_GetPrimaryByAOPNumber(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAOPNumberByPrimary(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAOPNumberByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAOPNumberByPrimary(vPrimary)
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
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
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimaryByPrimarySupplier(vPrimaryCustomer).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByAOPNumber(ByVal vInvoiceNumber As String) As Double
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByAOPNumber(vInvoiceNumber).HasValue Then
                Return DTA.ScalarQuery_GetJmlByAOPNumber(vInvoiceNumber).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub SimpanDataPurchaseBalance()
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Obj As New ClsTransaksi_Kartu_APAR
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjCD As New Core_Module.ClsCekData
        Dim BalanceBaru As Decimal = 0
        Dim PrimaryAkunAP As Double = -1

        Try
            PrimaryAkunAP = Me.mPrimaryAP
            '  PrimaryAkunAP = ObjCur.GetPrimaryAkunAP(Me.mPrimaryCurrency)
            BalanceBaru = ObjCD.DigitGrouping(Me.mBalance * Me.mExchangeRate, ObjCur.GetDecimal_DefaultCurrency)

            DTA.Insert(Me.mPrimarysupp, Me.mPrimary, TipeBalance, Me.mAOPNumber, Me.mAOPDate, Me.mRemarks, Me.mDueDate, _
            BalanceBaru, 0, 0, BalanceBaru, _
            0, Me.mCreatedBy, _
            Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, Me.mPrimaryCurrency, _
            Me.mBalance, Me.mExchangeRate, PrimaryAkunAP)

            'simpan data kartu AP
            Obj.Simpan(Me.mPrimary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual, Me.mAOPNumber, _
            Me.mPrimarysupp, 0, ObjCD.DigitGrouping(Me.mBalance * Me.mExchangeRate, ObjCur.GetDecimal_DefaultCurrency), _
            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, Me.mAOPDate, _
            Me.mPrimary, Me.mAOPDate, Math.Abs((Me.mDueDate - Me.mAOPDate).Days), ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual, _
            Me.mPrimaryCurrency, 0, Me.mBalance, Me.mExchangeRate, PrimaryAkunAP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
            ObjCur = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub HapusDataPurchaseBalance(ByVal vPrimaryPurchaseBalance As Double)
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            DTA.Delete(vPrimaryPurchaseBalance, TipeBalance)
            'hapus data kartu
            Obj.Hapus(vPrimaryPurchaseBalance, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter

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

    Public Function GetBalanceOriginal(ByVal vPrimary As Double, ByVal DTA As DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter) As Decimal
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

    Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiPurchaseBalance.balance_purchase_viewDataTable, _
       ByVal ij As FrmTransaksiJournal_Purchase, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimarySupp As Double, _
       ByVal vPrimaryJurnal As Double, ByVal vDef_Prec As Integer)
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.balance_purchase_viewTableAdapter
        Dim TBLDummy As New DseTransaksiPurchaseBalance.balance_purchase_viewDataTable
        Dim DTAJurnal As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBLJurnal As New DseTransaksiPurchaseBalance.tbltrans_jurnal_item_paymentDataTable
        Dim rowJurnal As DseTransaksiPurchaseBalance.tbltrans_jurnal_item_paymentRow
        Dim Filter As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim row, row1 As DseTransaksiPurchaseBalance.balance_purchase_viewRow
        Dim rowdgv As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ketemu As Boolean = False
        Dim AdaNonIDR As Boolean = False
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeBalance2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
        Dim TipeBalance1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim TipeBalance3 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
        Dim TipeBalance4 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
        Dim TipeBalance5 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
        Dim TipeBalance6 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
        Dim ObjHistoricalDP As New ClsTransaksiBalanceManual_DP 'add aul
        Dim ObjPR As New ClsTransaksiPurchaseReturn
        Dim ObjPRBalance As New ClsTransaksiPurchaseReturnBalance
        Dim ObjCDNote As New ClsTransaksiCreditNote
        Dim ObjAkun As New ClsMasterAkun

        'Update SMESS 2026
        Dim DTAPONo As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_aopTableAdapter
        Dim TBLPONo As New DseTransaksiPurchaseBalance.tbltrans_aopDataTable
        Dim DTAReturnPONo As New DseTransaksiPurchaseBalanceTableAdapters.tblReturnPONoTableAdapter
        Dim TBLReturnPONo As New DseTransaksiPurchaseBalance.tblReturnPONoDataTable

        Try
            vDGV.Rows.Clear()
            'isikan data dari database
            If vPrimarySupp = -1 Then
                DTA.FillBy_All(TBL)
            Else
                DTA.FillBy_PrimarySupp(TBL, vPrimarySupp)
            End If

            'Update SMESS 2026
            DTAPONo.Connection.ConnectionString = Core_Module.GetConstring
            DTAPONo.FillBy_PrimarySupplier(TBLPONo, vPrimarySupp)
            DTAReturnPONo.Connection.ConnectionString = Core_Module.GetConstring
            DTAReturnPONo.FillBy_PrimarySupplier(TBLReturnPONo, vPrimarySupp)

            'modifikasi nilai received dan receivement berdasarkan data jurnal
            'bila proses edit
            If vPrimaryJurnal <> -1 Then
                DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)

                For Each rowJurnal In TBLJurnal.Rows
                    TBLDummy.Rows.Clear()
                    DTA.FillBy_DocumentAcuan(TBLDummy, vPrimarySupp, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan)

                    If TBLDummy.Rows.Count > 0 Then
                        For Each row In TBLDummy.Rows
                            If rowJurnal.primary_currency = row.primary_currency And rowJurnal.primary_akun_ap_ar = row.primary_akun_ap Then
                                If row.Tipe = TipeBalance3 OrElse row.Tipe = TipeBalance4 Then
                                    'Credit/Debit Note
                                    'ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No)
                                    ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No, row.Due_Date)
                                    'row.Supplier_Invoice_No = "-"
                                    '  ObjPR.GetNoTransaksi_Remarks_NoInvoice(row.AOP_Number, row.Remarks, row.Primary_AOP, row.AOP_Date, row.Supplier_Invoice_No)
                                End If

                                Filter = "Primary = " + row.Primary.ToString
                                row.Paid -= rowJurnal.Payment
                                row.Payment = rowJurnal.Payment
                                row.paid_original -= rowJurnal.payment_original
                                row.Payment_Original = rowJurnal.payment_original
                                If ij.ObjJurnal.ApakahSudahRevaluasi(ij.dtpDate.Value) = True Then
                                    row.exchange_rate = rowJurnal.exchange_rate
                                End If

                                TBL.DefaultView.RowFilter = Filter
                                If TBL.DefaultView.Count = 0 Then
                                    row1 = TBL.Newbalance_purchase_viewRow
                                    With row1
                                        .Balance = row.Balance
                                        .Due_Date = row.Due_Date
                                        .Duty_stamp = row.Duty_stamp
                                        .AOP_Date = row.AOP_Date
                                        .AOP_Number = row.AOP_Number
                                        .Primary = row.Primary
                                        .Primary_Supplier = row.Primary_Supplier
                                        .Primary_AOP = row.Primary_AOP
                                        .Paid = row.Paid
                                        .Payment = row.Payment
                                        .Remarks = row.Remarks
                                        .Purchase = row.Purchase
                                        .Purchase_VAT = row.Purchase_VAT
                                        .Tipe = row.Tipe
                                        .Supplier_Invoice_No = row.Supplier_Invoice_No
                                        .primary_currency = row.primary_currency
                                        .Default_Precision = row.Default_Precision
                                        .Currency_Code = row.Currency_Code
                                        .paid_original = row.paid_original
                                        .Payment_Original = row.Payment_Original
                                        .exchange_rate = row.exchange_rate
                                        .primary_akun_ap = row.primary_akun_ap
                                        .balance_original = row.balance_original
                                    End With
                                    TBL.Rows.Add(row1)
                                Else
                                    row1 = TBL.DefaultView(0).Row
                                    row1.Paid = row.Paid
                                    row1.Payment = row.Payment
                                    row1.paid_original = row.paid_original
                                    row1.Payment_Original = row.Payment_Original
                                    row1.exchange_rate = row.exchange_rate
                                End If
                            End If
                        Next
                    End If
                Next
            End If

            'kkhusus untuk purchase return, kembali no transaksi dan remarks
            For Each row In TBL.Rows
                If row.Tipe = TipeBalance1 Then
                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                    ObjPRBalance.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No)
                End If
                If row.Tipe = TipeBalance2 Then
                    'purchase return 
                    ObjPR.GetNoTransaksi_Remarks_NoInvoice(row.AOP_Number, row.Remarks, row.Primary_AOP, row.AOP_Date, row.Supplier_Invoice_No)
                End If

                If row.Tipe = TipeBalance3 OrElse row.Tipe = TipeBalance4 Then
                    'Credit/Debit Note
                    'row.Supplier_Invoice_No = ObjCDNote.GetSuppInvoiceNoByPrimaryCDNote(row.Primary_AOP)

                    'ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No)
                    'row.Supplier_Invoice_No = "-"
                    '  ObjPR.GetNoTransaksi_Remarks_NoInvoice(row.AOP_Number, row.Remarks, row.Primary_AOP, row.AOP_Date, row.Supplier_Invoice_No)
                    ObjCDNote.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No, row.Due_Date)
                End If

                If row.Tipe = TipeBalance5 Then
                    'Purchase balance DP 
                    row.Supplier_Invoice_No = ObjHistoricalDP.GetSuppInvoiceNo(row.Primary_AOP)
                End If
                If row.Tipe = TipeBalance6 Then
                    'Purchase balance Manual
                    row.Supplier_Invoice_No = ""
                End If
            Next

            If vDGV Is Nothing Then
                MessageBox.Show("vDGV masih Nothing.")
                Exit Sub
            End If

            'isikan ke dalam datagridview
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[AOP Date], [AOP Number], Primary"

            For Each dv As DataRowView In TBL.DefaultView
                row = dv.Row

                'Update SMESS 2026
                If row.Tipe = 0 Then
                    row.PO_No = Me.IsiPONo(TBLPONo, row.Primary_AOP)
                End If
                If row.Tipe = 3 Then
                    row.PO_No = Me.IsiReturnPONo(TBLReturnPONo, row.Primary_AOP)
                End If
                If row.Balance <> 0 Or row.Payment <> 0 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowdgv
                        If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                            AdaNonIDR = True
                        End If
                        If row.Purchase = 0 Then
                            row.AOP_Number = row.AOP_Number + " (Tax + Dutystamp + etc)"
                        End If
                        .Cells(ij.ColAOPNumber.Index).Value = row.AOP_Number
                        .Cells(ij.ColAOPDate.Index).Value = row.AOP_Date
                        .Cells(ij.ColPurchaseValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColVatPurchaseValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase_VAT, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColDutyStamp.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColTotalPurchase_VAT_Value.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase + row.Purchase_VAT + row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColPaid.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Paid, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColBalancePurchase.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Balance, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColPayment.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Payment, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
                        .Cells(ij.ColPrimaryPurchaseBalance.Index).Value = row.Primary
                        .Cells(ij.ColDueDatePurchase.Index).Value = row.Due_Date
                        .Cells(ij.ColTipeDocumentAcuan_Purchase.Index).Value = row.Tipe
                        .Cells(ij.ColPrimaryDocumentAcuan_Purchase.Index).Value = row.Primary_AOP
                        TipeBalance = row.Tipe
                        'If TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic OrElse _
                        '    TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                        '    .Cells(ij.ColSupplierInvoiceNumber.Index).Value = row.Supplier_Invoice_No
                        'Else
                        '    .Cells(ij.ColSupplierInvoiceNumber.Index).Value = ""
                        'End If
                        .Cells(ij.ColSupplierInvoiceNumber.Index).Value = row.Supplier_Invoice_No

                        .Cells(ij.ColAOPRemarks.Index).Value = row.Remarks
                        .Cells(ij.ColPrimarySupplier.Index).Value = row.Primary_Supplier
                        .Cells(ij.Col_Purchase_Pay.Index).Value = "Full Payment"
                        .Cells(ij.ColCurrencyBalance.Index).Value = row.Currency_Code
                        .Cells(ij.ColCurrencyDefPrec.Index).Value = row.Default_Precision
                        .Cells(ij.ColPrimaryCurrencyBalance.Index).Value = row.primary_currency
                        .Cells(ij.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                        .Cells(ij.ColPaymentOriginal.Index).Value = ObjCD.DigitGrouping(row.Payment_Original, row.Default_Precision)
                        .Cells(ij.ColPaidOriginal.Index).Value = ObjCD.DigitGrouping(row.paid_original, row.Default_Precision)
                        .Cells(ij.ColExchangeRateBalance.Index).Value = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                        .Cells(ij.ColBalance1.Index).Value = CDec(.Cells(ij.ColBalancePurchase.Index).Value) + _
                        Math.Round(CDec(row.Payment_Original * row.exchange_rate), vDef_Prec) + _
                        CDec(.Cells(ij.ColPaid.Index).Value)
                        ' nilai diatas ini sebenarnya sama dengan received, hanya saja menjadi salah untuk journal yang sudah direaluasi kemudian belum dibetulkan, diganti untuk mengakomodasi proses pembetulan
                        .Cells(ij.ColBalanceOriginal02.Index).Value = .Cells(ij.ColBalanceOriginal.Index).Value
                        .Cells(ij.ColBalanceOriginal1.Index).Value = CDec(.Cells(ij.ColBalanceOriginal.Index).Value) + _
                        CDec(.Cells(ij.ColPaymentOriginal.Index).Value) + CDec(.Cells(ij.ColPaidOriginal.Index).Value)
                        .Cells(ij.ColPrimaryAkunBalance.Index).Value = row.primary_akun_ap
                        .Cells(ij.ColAPAccount.Index).Value = ObjAkun.GetAccountCode(row.primary_akun_ap) & " - " & ObjAkun.GetDescription_ByPrimaryAcc(row.primary_akun_ap)

                        'Update SMESS 2026
                        .Cells(ij.ColPONo.Index).Value = row.PO_No
                    End With
                End If
            Next

            'jika semua baris currencynya IDR, maka exchage rate tidak perlu ada
            If AdaNonIDR = True Then
                ij.ColExchangeRateBalance.Visible = True
            Else
                ij.ColExchangeRateBalance.Visible = False
            End If
            ij.ColAOPDate.DefaultCellStyle.Format = "dd MMM yyyy"
            ij.ColDueDatePurchase.DefaultCellStyle.Format = "dd MMM yyyy"

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            If vPrimarySupp <> -1 Then
                vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjPR = Nothing
            ObjPRBalance = Nothing

            'Update SMESS 2026
            DTAPONo.Dispose()
            TBLPONo.Dispose()
            DTAReturnPONo.Dispose()
            TBLReturnPONo.Dispose()
        End Try
    End Sub

    'Update SMESS 2026
    'Public Function IsiPONo(ByVal vTBLPONo As DseTransaksiPurchaseBalance.tbltrans_aopDataTable, ByVal vPrimaryAOP As Double) As String
    '    Dim retval As String = ""
    '    Dim vTemp As String = ""
    '    Dim vMaxLength As Integer = 150
    '    Dim vSuffix As String = "..."
    '    Dim row As DseTransaksiPurchaseBalance.tbltrans_aopRow
    '    Dim i As Integer = 0
    '    Try
    '        vTBLPONo.DefaultView.RowFilter = "[Primary]=" & vPrimaryAOP
    '        If vTBLPONo.DefaultView.Count > 0 Then
    '            For i = 0 To vTBLPONo.DefaultView.Count - 1
    '                row = vTBLPONo.DefaultView(i).Row
    '                If String.IsNullOrEmpty(row.PO_No) = False Then
    '                    If String.IsNullOrEmpty(retval) Then
    '                        retval = row.PO_No
    '                    Else
    '                        retval = retval & ", " & row.PO_No
    '                    End If
    '                ElseIf String.IsNullOrEmpty(row.PO_No_DP) = False Then
    '                    If String.IsNullOrEmpty(retval) Then
    '                        retval = row.PO_No_DP
    '                    Else
    '                        retval = retval & ", " & row.PO_No_DP
    '                    End If
    '                End If
    '            Next
    '        End If
    '        Return retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return retval
    '    Finally
    '        vTBLPONo.DefaultView.RowFilter = Nothing
    '    End Try
    'End Function
    Public Function IsiPONo(ByVal vTBLPONo As DseTransaksiPurchaseBalance.tbltrans_aopDataTable, ByVal vPrimaryAOP As Double) As String
        Dim retval As String = ""
        Dim vTemp As String = ""
        Dim vMaxLength As Integer = 150
        Dim vSuffix As String = " ..."
        Dim row As DseTransaksiPurchaseBalance.tbltrans_aopRow
        Dim i As Integer = 0
        Try
            vTBLPONo.DefaultView.RowFilter = "[Primary]=" & vPrimaryAOP
            If vTBLPONo.DefaultView.Count > 0 Then
                For i = 0 To vTBLPONo.DefaultView.Count - 1
                    row = vTBLPONo.DefaultView(i).Row

                    If String.IsNullOrEmpty(row.PO_No) = False Then
                        vTemp = row.PO_No
                    ElseIf String.IsNullOrEmpty(row.PO_No_DP) = False Then
                        vTemp = row.PO_No_DP
                    Else
                        vTemp = ""
                    End If

                    If vTemp <> "" Then
                        If String.IsNullOrEmpty(retval) Then
                            vTemp = vTemp
                        Else
                            vTemp = retval & ", " & vTemp
                        End If
                        If vTemp.Length > vMaxLength Then
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
            vTBLPONo.DefaultView.RowFilter = Nothing
        End Try
    End Function

    'Update SMESS 2026
    Public Function IsiReturnPONo(ByVal vTBLPONo As DseTransaksiPurchaseBalance.tblReturnPONoDataTable, ByVal vPrimaryAOP As Double) As String
        Dim retval As String = ""
        Dim vTemp As String = ""
        Dim vMaxLength As Integer = 150
        Dim vSuffix As String = " ..."
        Dim row As DseTransaksiPurchaseBalance.tblReturnPONoRow
        Dim i As Integer = 0
        Try
            vTBLPONo.DefaultView.RowFilter = "[Primary]=" & vPrimaryAOP
            If vTBLPONo.DefaultView.Count > 0 Then
                For i = 0 To vTBLPONo.DefaultView.Count - 1
                    row = vTBLPONo.DefaultView(i).Row
                    If String.IsNullOrEmpty(row.No_PO) = False Then
                        If String.IsNullOrEmpty(retval) Then
                            vTemp = row.No_PO
                        Else
                            vTemp = retval & ", " & row.No_PO
                        End If

                        If vTemp.Length > vMaxLength Then
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
            vTBLPONo.DefaultView.RowFilter = Nothing
        End Try
    End Function

    ' Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiPurchaseBalance.balance_purchase_viewDataTable, _
    'ByVal ij As FrmTransaksiJournal_Purchase, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimarySupp As Double, _
    'ByVal vPrimaryJurnal As Double, ByVal vDef_Prec As Integer)
    '     Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.balance_purchase_viewTableAdapter
    '     Dim TBLDummy As New DseTransaksiPurchaseBalance.balance_purchase_viewDataTable
    '     Dim DTAJurnal As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
    '     Dim TBLJurnal As New DseTransaksiPurchaseBalance.tbltrans_jurnal_item_paymentDataTable
    '     Dim rowJurnal As DseTransaksiPurchaseBalance.tbltrans_jurnal_item_paymentRow
    '     Dim Filter As String = ""
    '     Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
    '     Dim row, row1 As DseTransaksiPurchaseBalance.balance_purchase_viewRow
    '     Dim rowdgv As DataGridViewRow
    '     Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '     Dim ketemu As Boolean = False
    '     Dim AdaNonIDR As Boolean = False
    '     Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '     Dim TipeBalance2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '     Dim TipeBalance1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '     Dim ObjPR As New ClsTransaksiPurchaseReturn
    '     Dim ObjPRBalance As New ClsTransaksiPurchaseReturnBalance

    '     Try
    '         vDGV.Rows.Clear()
    '         'isikan data dari database
    '         If vPrimarySupp = -1 Then
    '             DTA.FillBy_All(TBL)
    '         Else
    '             DTA.FillBy_PrimarySupp(TBL, vPrimarySupp)
    '         End If

    '         'modifikasi nilai received dan receivement berdasarkan data jurnal
    '         'bila proses edit
    '         If vPrimaryJurnal <> -1 Then
    '             DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)

    '             For Each rowJurnal In TBLJurnal.Rows
    '                 TBLDummy.Rows.Clear()
    '                 DTA.FillBy_DocumentAcuan(TBLDummy, vPrimarySupp, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan)

    '                 If TBLDummy.Rows.Count > 0 Then
    '                     For Each row In TBLDummy.Rows
    '                         If rowJurnal.primary_currency = row.primary_currency And rowJurnal.primary_akun_ap_ar = row.primary_akun_ap Then
    '                             Filter = "Primary = " + row.Primary.ToString
    '                             row.Paid -= rowJurnal.Payment
    '                             row.Payment = rowJurnal.Payment
    '                             row.paid_original -= rowJurnal.payment_original
    '                             row.Payment_Original = rowJurnal.payment_original
    '                             If ij.ObjJurnal.ApakahSudahRevaluasi(ij.dtpDate.Value) = True Then
    '                                 row.exchange_rate = rowJurnal.exchange_rate
    '                             End If

    '                             TBL.DefaultView.RowFilter = Filter
    '                             If TBL.DefaultView.Count = 0 Then
    '                                 row1 = TBL.Newbalance_purchase_viewRow
    '                                 With row1
    '                                     .Balance = row.Balance
    '                                     .Due_Date = row.Due_Date
    '                                     .Duty_stamp = row.Duty_stamp
    '                                     .AOP_Date = row.AOP_Date
    '                                     .AOP_Number = row.AOP_Number
    '                                     .Primary = row.Primary
    '                                     .Primary_Supplier = row.Primary_Supplier
    '                                     .Primary_AOP = row.Primary_AOP
    '                                     .Paid = row.Paid
    '                                     .Payment = row.Payment
    '                                     .Remarks = row.Remarks
    '                                     .Purchase = row.Purchase
    '                                     .Purchase_VAT = row.Purchase_VAT
    '                                     .Tipe = row.Tipe
    '                                     .Supplier_Invoice_No = row.Supplier_Invoice_No
    '                                     .primary_currency = row.primary_currency
    '                                     .Default_Precision = row.Default_Precision
    '                                     .Currency_Code = row.Currency_Code
    '                                     .paid_original = row.paid_original
    '                                     .Payment_Original = row.Payment_Original
    '                                     .exchange_rate = row.exchange_rate
    '                                     .primary_akun_ap = row.primary_akun_ap
    '                                     .balance_original = row.balance_original
    '                                 End With
    '                                 TBL.Rows.Add(row1)
    '                             Else
    '                                 row1 = TBL.DefaultView(0).Row
    '                                 row1.Paid = row.Paid
    '                                 row1.Payment = row.Payment
    '                                 row1.paid_original = row.paid_original
    '                                 row1.Payment_Original = row.Payment_Original
    '                                 row1.exchange_rate = row.exchange_rate
    '                             End If
    '                         End If
    '                     Next
    '                 End If
    '             Next
    '         End If

    '         'kkhusus untuk purchase return, kembali no transaksi dan remarks
    '         For Each row In TBL.Rows
    '             If row.Tipe = TipeBalance1 Then
    '                 'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
    '                 ObjPRBalance.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No)
    '             End If
    '             If row.Tipe = TipeBalance2 Then
    '                 'purchase return 
    '                 ObjPR.GetNoTransaksi_Remarks_NoInvoice(row.AOP_Number, row.Remarks, row.Primary_AOP, row.AOP_Date, row.Supplier_Invoice_No)
    '             End If

    '         Next

    '         'isikan ke dalam datagridview
    '         vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '         vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '         vDGV.Rows.Clear()
    '         TBL.DefaultView.RowFilter = ""
    '         TBL.DefaultView.Sort = "[AOP Date], [AOP Number], Primary"

    '         For Each dv As DataRowView In TBL.DefaultView
    '             row = dv.Row

    '             If row.Balance <> 0 Or row.Payment <> 0 Then
    '                 vDGV.Rows.Add()
    '                 rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
    '                 With rowdgv
    '                     If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
    '                         AdaNonIDR = True
    '                     End If
    '                     If row.Purchase = 0 Then
    '                         row.AOP_Number = row.AOP_Number + " (Tax + Dutystamp + etc)"
    '                     End If
    '                     .Cells(ij.ColAOPNumber.Index).Value = row.AOP_Number
    '                     .Cells(ij.ColAOPDate.Index).Value = Format(row.AOP_Date, "dd MMM yyyy")
    '                     .Cells(ij.ColPurchaseValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColVatPurchaseValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase_VAT, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColDutyStamp.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColTotalPurchase_VAT_Value.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Purchase + row.Purchase_VAT + row.Duty_stamp, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColPaid.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Paid, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColBalancePurchase.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Balance, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColPayment.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Payment, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)
    '                     .Cells(ij.ColPrimaryPurchaseBalance.Index).Value = row.Primary
    '                     .Cells(ij.ColDueDatePurchase.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
    '                     .Cells(ij.ColTipeDocumentAcuan_Purchase.Index).Value = row.Tipe
    '                     .Cells(ij.ColPrimaryDocumentAcuan_Purchase.Index).Value = row.Primary_AOP
    '                     TipeBalance = row.Tipe
    '                     If TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Or TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
    '                         .Cells(ij.ColSupplierInvoiceNumber.Index).Value = row.Supplier_Invoice_No
    '                     Else
    '                         .Cells(ij.ColSupplierInvoiceNumber.Index).Value = ""
    '                     End If
    '                     .Cells(ij.ColAOPRemarks.Index).Value = row.Remarks
    '                     .Cells(ij.ColPrimarySupplier.Index).Value = row.Primary_Supplier
    '                     .Cells(ij.Col_Purchase_Pay.Index).Value = "Full Payment"
    '                     .Cells(ij.ColCurrencyBalance.Index).Value = row.Currency_Code
    '                     .Cells(ij.ColCurrencyDefPrec.Index).Value = row.Default_Precision
    '                     .Cells(ij.ColPrimaryCurrencyBalance.Index).Value = row.primary_currency
    '                     .Cells(ij.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
    '                     .Cells(ij.ColPaymentOriginal.Index).Value = ObjCD.DigitGrouping(row.Payment_Original, row.Default_Precision)
    '                     .Cells(ij.ColPaidOriginal.Index).Value = ObjCD.DigitGrouping(row.paid_original, row.Default_Precision)
    '                     .Cells(ij.ColExchangeRateBalance.Index).Value = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
    '                     .Cells(ij.ColBalance1.Index).Value = CDec(.Cells(ij.ColBalancePurchase.Index).Value) + _
    '                     Math.Round(CDec(row.Payment_Original * row.exchange_rate), vDef_Prec) + _
    '                     CDec(.Cells(ij.ColPaid.Index).Value)
    '                     ' nilai diatas ini sebenarnya sama dengan received, hanya saja menjadi salah untuk journal yang sudah direaluasi kemudian belum dibetulkan, diganti untuk mengakomodasi proses pembetulan
    '                     .Cells(ij.ColBalanceOriginal02.Index).Value = .Cells(ij.ColBalanceOriginal.Index).Value
    '                     .Cells(ij.ColBalanceOriginal1.Index).Value = CDec(.Cells(ij.ColBalanceOriginal.Index).Value) + _
    '                     CDec(.Cells(ij.ColPaymentOriginal.Index).Value) + CDec(.Cells(ij.ColPaidOriginal.Index).Value)
    '                     .Cells(ij.ColPrimaryAkunBalance.Index).Value = row.primary_akun_ap
    '                 End With
    '             End If
    '         Next

    '         'jika semua baris currencynya IDR, maka exchage rate tidak perlu ada
    '         If AdaNonIDR = True Then
    '             ij.ColExchangeRateBalance.Visible = True
    '         Else
    '             ij.ColExchangeRateBalance.Visible = False
    '         End If
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         DTA.Dispose()
    '         If vPrimarySupp <> -1 Then
    '             vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    '             vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '         End If
    '         Windows.Forms.Cursor.Current = Cursors.Arrow
    '         ObjPR = Nothing
    '         ObjPRBalance = Nothing
    '     End Try
    ' End Sub

    Public Function CetakLaporanAPDueDate(ByVal vIsDueDate As Boolean, ByVal vDueDate As Date, _
          ByVal vIsAllSupplier As Boolean, ByVal vPrimarySupp As Double, ByVal TBL As DseTransaksiPurchaseBalance.balance_purchase_viewDataTable, _
          ByVal vDef_Prec As Integer, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vInvoiceNoType As EnumCetakanInvoiceNo) As Boolean
        Dim DTA As New DseTransaksiPurchaseBalanceTableAdapters.balance_purchase_viewTableAdapter
        Dim Filter As String = ""
        Dim row As DseTransaksiPurchaseBalance.balance_purchase_viewRow
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        Dim TipeBalance2 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
        Dim TipeBalance1 As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim ObjPR As New ClsTransaksiPurchaseReturn
        Dim ObjPRBalance As New ClsTransaksiPurchaseReturnBalance

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

        Dim dtaSupp As New DseTransaksiPurchaseBalanceTableAdapters.tblmaster_supplierTableAdapter
        Dim tblSupp As New DseTransaksiPurchaseBalance.tblmaster_supplierDataTable
        Dim rowSupp As DseTransaksiPurchaseBalance.tblmaster_supplierRow

        Try
            ''set def value
            vDef_Prec = ObjCur.GetDecimal_DefaultCurrency
            dtaSupp.Fill(tblSupp)

            If vIsAllSupplier = True Then
                DTA.FillBy_All(TBL)
            Else
                DTA.FillBy_PrimarySupp(TBL, vPrimarySupp)
            End If

            'kkhusus untuk purchase return, kembali no transaksi dan remarks
            For Each row In TBL.Rows
                If row.Tipe = TipeBalance1 Then
                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                    ObjPRBalance.GetNoTransaksi_Remarks_Tanggal(row.AOP_Number, row.Remarks, row.AOP_Date, row.Primary_AOP, row.Supplier_Invoice_No)
                End If
                If row.Tipe = TipeBalance2 Then
                    'purchase return 
                    ObjPR.GetNoTransaksi_Remarks_NoInvoice(row.AOP_Number, row.Remarks, row.Primary_AOP, row.AOP_Date, row.Supplier_Invoice_No)
                End If
            Next

            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[AOP Date], [AOP Number], Primary"

            For Each dv As DataRowView In TBL.DefaultView
                row = dv.Row
                If row.Balance <> 0 Or row.Payment <> 0 Then

                    ''batasan due date
                    If vIsDueDate = True Then
                        If Not row.Due_Date > vDueDate Then

                            ''simpan komponen
                            With row
                                rowCetak = tblCetak.NewAPAR_PrintRow
                                rowCetak.Primary = .Primary
                                rowCetak.Primary_Supplier_Customer = .Primary_Supplier
                                tblSupp.DefaultView.RowFilter = "Primary =" & .Primary_Supplier.ToString
                                If tblSupp.DefaultView.Count > 0 Then
                                    rowSupp = tblSupp.DefaultView(0).Row
                                    rowCetak.Supplier_Customer_Code = rowSupp.Supplier_Code
                                    rowCetak.Supplier_Customer_Desc = rowSupp.Supplier_Name
                                Else
                                    rowCetak.Supplier_Customer_Code = "N/A"
                                    rowCetak.Supplier_Customer_Desc = "N/A"
                                End If
                                rowCetak.Primary_Currency = .primary_currency
                                rowCetak.Currency_Code = .Currency_Code
                                rowCetak.Currency_Desc = "n/a"
                                TipeBalance = row.Tipe
                                If vInvoiceNoType = EnumCetakanInvoiceNo.SupplierInv Then
                                    If TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Or TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                                        rowCetak.Invoice_No = .Supplier_Invoice_No
                                    Else
                                        rowCetak.Invoice_No = ""
                                    End If
                                ElseIf vInvoiceNoType = EnumCetakanInvoiceNo.AOPNo Then
                                    rowCetak.Invoice_No = .AOP_Number
                                End If
                                rowCetak.Invoice_Date = row.AOP_Date
                                rowCetak.Due_Date = .Due_Date
                                rowCetak.Balance_Original = objCD.DigitGrouping(row.balance_original, row.Default_Precision)
                                rowCetak.Exchange_Rate = objCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                                rowCetak.Balance_IDR = CDec(objCD.DigitGrouping(Math.Round(row.Balance, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)) + _
                                        Math.Round(CDec(row.Payment_Original * row.exchange_rate), vDef_Prec) + _
                                        CDec(objCD.DigitGrouping(Math.Round(row.Paid, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec))
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

                        ''tanpa filter due date
                        ''simpan komponen
                        With row
                            rowCetak = tblCetak.NewAPAR_PrintRow
                            rowCetak.Primary = .Primary
                            rowCetak.Primary_Supplier_Customer = .Primary_Supplier
                            tblSupp.DefaultView.RowFilter = "Primary =" & .Primary_Supplier.ToString
                            If tblSupp.DefaultView.Count > 0 Then
                                rowSupp = tblSupp.DefaultView(0).Row
                                rowCetak.Supplier_Customer_Code = rowSupp.Supplier_Code
                                rowCetak.Supplier_Customer_Desc = rowSupp.Supplier_Name
                            Else
                                rowCetak.Supplier_Customer_Code = "N/A"
                                rowCetak.Supplier_Customer_Desc = "N/A"
                            End If
                            rowCetak.Primary_Currency = .primary_currency
                            rowCetak.Currency_Code = .Currency_Code
                            rowCetak.Currency_Desc = "n/a"
                            TipeBalance = row.Tipe
                            If TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Or TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                                rowCetak.Invoice_No = .Supplier_Invoice_No
                            Else
                                rowCetak.Invoice_No = ""
                            End If
                            rowCetak.Invoice_Date = row.AOP_Date
                            rowCetak.Due_Date = .Due_Date
                            rowCetak.Balance_Original = objCD.DigitGrouping(row.balance_original, row.Default_Precision)
                            rowCetak.Exchange_Rate = objCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                            rowCetak.Balance_IDR = CDec(objCD.DigitGrouping(Math.Round(row.Balance, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec)) + _
                                    Math.Round(CDec(row.Payment_Original * row.exchange_rate), vDef_Prec) + _
                                    CDec(objCD.DigitGrouping(Math.Round(row.Paid, vDef_Prec, MidpointRounding.AwayFromZero), vDef_Prec))
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
                        paramBalanceOriginal = Me.objCD.DigitGrouping(rowCurrency.Balance_Original.ToString, rowCurrency.Precision)
                    End If
                    paramBalanceIdr = Me.objCD.DigitGrouping(rowCurrency.Balance_IDR.ToString, vDef_Prec)
                Else
                    paramCurrencyCode = paramCurrencyCode & vbNewLine & rowCurrency.Currency_Code
                    If rowCurrency.Currency_Code = "IDR" Then
                        paramBalanceOriginal = paramBalanceOriginal & vbNewLine & "-"
                    Else
                        paramBalanceOriginal = paramBalanceOriginal & vbNewLine & Me.objCD.DigitGrouping(rowCurrency.Balance_Original.ToString, rowCurrency.Precision)
                    End If
                    paramBalanceIdr = paramBalanceIdr & vbNewLine & Me.objCD.DigitGrouping(rowCurrency.Balance_IDR.ToString, vDef_Prec)
                End If
            Next

            reportFile.SetDataSource(dsCetak)

            reportFile.SetParameterValue("TransactionType", "AP")
            If vIsDueDate = True Then
                reportFile.SetParameterValue("TitleOfReport", "ACCOUNT PAYABLE DUE UNTIL  " & Format(vDueDate, "dd MMMM yyyy"))
            Else
                reportFile.SetParameterValue("TitleOfReport", "ACCOUNT PAYABLE DUE")
            End If
            reportFile.SetParameterValue("DueDate", vDueDate)
            reportFile.SetParameterValue("TitleOfCurrency", paramCurrencyCode)
            reportFile.SetParameterValue("TotalBalanceOriginal", paramBalanceOriginal)
            reportFile.SetParameterValue("TotalBalanceIDR", paramBalanceIdr)
            reportFile.SetParameterValue("Precision", vDef_Prec)
            reportFile.SetParameterValue("PayableOrReceiveable", "Payable")
            reportFile.SetParameterValue("SupplierOrCustomer", "SUPPLIER")
            reportFile.SetParameterValue("DefPrecExchangeRate", Core_Module.DefPrecExchangeRate)

            If vInvoiceNoType = EnumCetakanInvoiceNo.AOPNo Then
                reportFile.SetParameterValue(reportFile.Parameter_TipeDokumen.ParameterFieldName, "AOP")
            Else
                reportFile.SetParameterValue(reportFile.Parameter_TipeDokumen.ParameterFieldName, "INVOICE")
            End If


            vCrv.ReportSource = reportFile
            vCrv.Zoom(90)
            vCrv.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjPR = Nothing
            ObjPRBalance = Nothing
        End Try
    End Function
End Class
