Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturnBalance

    Private mPrimary As Double = -1
    Private mPrimaryCustomer As Double = -1
    Private mPrimaryAccount As Double = -1
    Private mAOSRNumber As String = ""
    Private mAOSRDate As Date = Now
    Private mSalesReturn As Double = 0
    Private mVATReturn As Double = 0
    Private mTotal As Double = 0
    Private mDueDate As Date = Now
    Private mPaid As Double = 0
    Private mBalance As Double = 0
    Private mRemarks As String = ""
    Private mDutystamp As Double = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim mPrimaryCurrency As Double = -1
    Dim mExchangeRate As Decimal = 1

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            Me.mExchangeRate = value
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

    Public Property Dutystamp() As Double
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

    Public Property Primary_Customer() As Double
        Get
            Return mPrimaryCustomer
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomer = value
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

    Public Property AOSRNumber() As String
        Get
            Return mAOSRNumber
        End Get
        Set(ByVal value As String)
            mAOSRNumber = value
        End Set
    End Property

    Public Property AOSRDate() As Date
        Get
            Return mAOSRDate
        End Get
        Set(ByVal value As Date)
            mAOSRDate = value
        End Set
    End Property

    Public Property SalesReturn() As Double
        Get
            Return mSalesReturn
        End Get
        Set(ByVal value As Double)
            mSalesReturn = value
        End Set
    End Property

    Public Property VATReturn() As Double
        Get
            Return mVATReturn
        End Get
        Set(ByVal value As Double)
            mVATReturn = value
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter

        Try
            If DTA.Insert(Me.Primary_Customer, Me.mAOSRNumber, Me.mAOSRDate, _
                          Me.mDueDate, Me.mRemarks, vUserCode, Now, vUserCode, _
                          Now, Now, Me.mPrimaryCurrency, Me.mExchangeRate, _
                          Me.mSalesReturn, Me.mVATReturn, Me.mTotal, Me.mPaid, _
                          Me.mBalance, Me.mDutystamp, Me.mPrimaryAccount) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByAOSRNumber(Me.mAOSRNumber)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data balance
                Me.SimpanData_SalesReturnBalance()

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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter

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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tblcaritrans_historicalsalesreturnmanual_viewTableAdapter

        Try
            If vSearchBy = EnumSearchCriteria.ByAOSRNumber Then
                Return DTA.ScalarQuery_GetJmlByAOSRNumberLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQuery_GetJmlByCustomerCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustomerNameLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Try
            If DTA.Update(Me.Primary_Customer, Me.mAOSRNumber, Me.mAOSRDate, _
                          Me.mDueDate, Me.mRemarks, vUserCode, Now, Now, _
                          Me.mPrimaryCurrency, Me.mExchangeRate, Me.mSalesReturn, _
                          Me.mVATReturn, Me.mTotal, Me.mPaid, Me.mBalance, Me.mDutystamp, _
                          Me.mPrimaryAccount, Me.mPrimary) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'simpan data item
                Me.HapusData_SalesReturnBalance(Me.mPrimary)

                'simpan data balance
                Me.SimpanData_SalesReturnBalance()

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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimary, vUserCode) = True Then
                'hapus data item
                Me.HapusData_SalesReturnBalance(vPrimary)

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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByAOSRNumber(vInvoiceNumber).HasValue Then
                If DTA.ScalarQuery_GetJmlByAOSRNumber(vInvoiceNumber).Value > 0 Then
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim PrimaryBalance As Double = -1
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
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

            '2. jika sudah ada sales return payment,maka tidak boleh edit --> obslete pa v.1.0.5
            'Jurnal = ObjJurnal.Get_SalesReturnBalance_Jurnal(vPrimary)
            'If Jurnal <> "" Then
            'If vTampilkanPesan = True Then
            'MsgBox("This sales return balance referred by sales return payment journal: " + Jurnal + _
            '". You cannot delete this data.", MsgBoxStyle.Information, "Information")
            'End If

            'Return False
            'End If

            '2.b. jika sudah ada sales receivement,maka tidak boleh edit
            Jurnal = ObjJurnal.Get_SalesBalance_SR_Jurnal(vPrimary)
            If Jurnal <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This sales return balance referred by sales journal: " + Jurnal + _
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim Obj As New ClsTransaksiJournal
        Dim PrimaryBalance As Double = -1
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
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

            '2. jika sudah ada sales return payment,maka tidak boleh edit --> obselete sejak penggabung sales receivement dan sales payment
            ' jurnal pada v.1.0.5
            Jurnal = ObjJurnal.Get_SalesReturnBalance_Jurnal(vPrimary)
            If Jurnal <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This sales return balance referred by sales return payment journal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            '2.b. jika sudah ada sales receivement,maka tidak boleh edit
            Jurnal = ObjJurnal.Get_SalesBalance_SR_Jurnal(vPrimary)
            If Jurnal <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("This sales return balance referred by sales journal: " + Jurnal + _
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
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesReturnBalance.tblcaritrans_historicalsalesreturnmanual_viewDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tblcaritrans_historicalsalesreturnmanual_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAOSRNumber Then
                QueryResult = DTA.FillBy_AOSRNumberLIKE(TBL, Criteria + "%")
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
                DGV.Columns(3).HeaderText = "SR Number"
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBL As New DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualDataTable
        Dim row As DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualRow
        Try
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mPrimaryCustomer = .Primary_Customer
                    Me.mAOSRNumber = .AOSR_Number
                    Me.mAOSRDate = .AOSR_Date
                    Me.mSalesReturn = .Sales_Return
                    Me.mVATReturn = .VAT_Return
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
                    Me.mExchangeRate = .exchange_rate
                    Me.mPrimaryCurrency = .primary_currency
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

    Public Function GetPrimaryByAOSRNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter

        Try
            Return DTA.ScalarQuery_GetPrimaryByAOSRNumber(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetNoTransaksi_Remarks_Tanggal(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vDate As Date, ByVal vPrimary As Double)
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBL As New DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualDataTable
        Dim row As DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualRow

        Try
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.Remarks
                vNoTransaksi = row.AOSR_Number
                vDate = row.AOSR_Date
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

    Public Function GetAOSRNumberByPrimary(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAOSRNumberByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAOSRNumberByPrimary(vPrimary)
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
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
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Try
            Return DTA.ScalarQuery_GetPrimaryByPrimaryCustomer(vPrimaryCustomer).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByAOSRNumber(ByVal vInvoiceNumber As String) As Double
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByAOSRNumber(vInvoiceNumber).HasValue Then
                Return DTA.ScalarQuery_GetJmlByAOSRNumber(vInvoiceNumber).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SimpanData_SalesReturnBalance()
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim Obj_Kartu As New ClsTransaksi_Kartu_APAR
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkunAR As Double = -1
        Dim ReturnVal As Decimal = 0

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            PrimaryAkunAR = Me.mPrimaryAccount
            'PrimaryAkunAR = ObjCur.GetPrimaryAkunAR(Me.mPrimaryCurrency)
            ReturnVal = Math.Round(Me.mBalance * Me.mExchangeRate * -1, DefPrec, MidpointRounding.AwayFromZero)

            DTA.Insert(Me.mPrimaryCustomer, Me.mPrimary, TipeBalance, Me.mAOSRNumber, Me.mAOSRDate, Me.mRemarks, Me.mDueDate, _
            ReturnVal, 0, 0, ReturnVal, 0, Me.mCreatedBy, _
            Me.mCreatedDate, Me.mModifiedBy, Me.mModifiedDate, Me.mTimestamp, Me.mPrimaryCurrency, Me.mBalance * -1, Me.mExchangeRate, PrimaryAkunAR)

            'simpan data kartu AR
            Obj_Kartu.Simpan(Me.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual, Me.mAOSRNumber, _
            Me.mPrimaryCustomer, 0, Math.Abs(ReturnVal), "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            Me.mAOSRDate, Me.Primary, Me.mAOSRDate, Math.Abs((Me.mDueDate - Me.mAOSRDate).Days), _
            ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual, Me.mPrimaryCurrency, 0, Math.Abs(Me.mBalance), Me.mExchangeRate, PrimaryAkunAR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj_Kartu = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub HapusData_SalesReturnBalance(ByVal vPrimarySalesReturnBalance As Double)
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
        Dim obj As New ClsTransaksi_Kartu_APAR

        Try
            DTA.Delete(vPrimarySalesReturnBalance, TipeBalance)
            obj.Hapus(vPrimarySalesReturnBalance, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            obj = Nothing
        End Try
    End Sub

    Public Sub IsiDGVBalance(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiSalesReturnBalance.balance_sr_viewDataTable, _
   ByVal ij As FrmTransaksiJournal_Sales, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimaryCust As Double, ByVal vPrimaryJurnal As Double, _
    ByVal Default_Prec As Integer)
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.balance_sr_viewTableAdapter
        Dim TBLDummy As New DseTransaksiSalesReturnBalance.balance_sr_viewDataTable
        Dim DTAJurnal As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBLJurnal As New DseTransaksiSalesReturnBalance.tbltrans_jurnal_item_paymentDataTable
        Dim rowJurnal As DseTransaksiSalesReturnBalance.tbltrans_jurnal_item_paymentRow
        Dim Filter As String = ""
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesReturn
        Dim row, row1 As DseTransaksiSalesReturnBalance.balance_sr_viewRow
        Dim rowdgv As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ketemu As Boolean = False

        Try
            'isikan data dari database
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()
            DTA.FillBy(TBL, vPrimaryCust)

            'modifikasi nilai received, receivement, dan balance berdasarkan data jurnal
            'bila proses edit
            If vPrimaryJurnal <> -1 Then
                DTAJurnal.Fill(TBLJurnal, vPrimaryJurnal, TipePayment)
                For Each rowJurnal In TBLJurnal.Rows
                    TBLDummy.Rows.Clear()
                    DTA.FillBy_DocumenAcuan(TBLDummy, rowJurnal.Primary_Document_Acuan, rowJurnal.Tipe_Document_Acuan, vPrimaryCust)
                    If TBLDummy.Rows.Count > 0 Then
                        row = TBLDummy.Rows(0)
                        Filter = "Primary = " + row.Primary.ToString
                        row.Paid -= rowJurnal.Payment
                        row.Payment = rowJurnal.Payment
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count = 0 Then
                            row1 = TBL.Newbalance_sr_viewRow
                            With row1
                                .Balance = row.Balance
                                .Due_Date = row.Due_Date
                                .Duty_stamp = row.Duty_stamp
                                .SR_Date = row.SR_Date
                                .SR_Number = row.SR_Number
                                .Primary = row.Primary
                                .Primary_Customer = row.Primary_Customer
                                .Primary_SR = row.Primary_SR
                                .Paid = row.Paid
                                .Payment = row.Payment
                                .Remarks = row.Remarks
                                .Sales_Return = row.Sales_Return
                                .Sales_Return_VAT = row.Sales_Return_VAT
                                .Tipe = row.Tipe
                            End With
                            TBL.Rows.Add(row1)
                        Else
                            row1 = TBL.DefaultView(0).Row
                            row1.Paid = row.Paid
                            row1.Payment = row.Payment
                        End If

                    End If
                Next
            End If

            'isikan ke dalam datagridview
            vDGV.Rows.Clear()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "SR Number"

            For i As Double = 0 To TBL.DefaultView.Count - 1
                row = TBL.DefaultView(i).Row
                If row.Balance <> 0 Or row.Payment <> 0 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowdgv
                        .Cells(ij.ColSalesReturnNumber.Index).Value = row.SR_Number
                        .Cells(ij.ColSRDate.Index).Value = Format(row.SR_Date, "dd MMM yyyy")
                        .Cells(ij.ColSalesReturnValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Sales_Return, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColVATReturnValue.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Sales_Return_VAT, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColDutystampReturn.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Duty_stamp, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColSalesReturn_VATReturn_Value.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Sales_Return + row.Sales_Return_VAT + row.Duty_stamp, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColPaidReturn.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Paid, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColBalanceReturn.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Balance, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColReturnPayment.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Payment, Default_Prec, MidpointRounding.AwayFromZero), Default_Prec)
                        .Cells(ij.ColPrimarySalesBalance.Index).Value = row.Primary
                        .Cells(ij.ColDueDate_SR.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                        .Cells(ij.ColPrimaryDocumenAcuan_SR.Index).Value = row.Primary_SR
                        .Cells(ij.ColTipeDocumentAcuan_SR.Index).Value = row.Tipe
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
