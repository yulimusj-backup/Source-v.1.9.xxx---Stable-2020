Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturn
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mNoUrut As Double = -1
    Private mPrimaryStockInPurchaseReturn As Double = -1
    Private mPrimaryGRN As Double = -1
    Private mTanggalTransaksi As Date = Now
    Private mSubTotal As Double = 0
    Private mTaxExclude As Boolean = True
    Private mTax As Double = 0
    Private mTaxAmount As Double = 0
    Private mTotal As Double = 0
    Private mReasonOfReturn As String = ""
    Private mRemarks As String = ""
    Private mSupplierDONumber As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice
    End Enum

    Public Enum EnumTipePayment As Integer
        SalesInvoice_SalesInvoiceDP
        SalesReturn
        AOP_AOPDP
        Purchase_Return
    End Enum

    Public Enum EnumTipe_PurchaseReturn As Integer
        Revisi_AOP
        Nota_Retur
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return (mNoTransaksi)
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
        End Set
    End Property

    Public Property PrimaryStockInPurchaseReturn() As Double
        Get
            Return mPrimaryStockInPurchaseReturn
        End Get
        Set(ByVal value As Double)
            mPrimaryStockInPurchaseReturn = value
        End Set
    End Property
    Public Property PrimaryGRN() As Double
        Get
            Return mPrimaryGRN
        End Get
        Set(ByVal value As Double)
            mPrimaryGRN = value
        End Set
    End Property
    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property
    Public Property SubTotal() As Double
        Get
            Return mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
        End Set
    End Property
    Public Property TaxExclude() As Boolean
        Get
            Return mTaxExclude
        End Get
        Set(ByVal value As Boolean)
            mTaxExclude = value
        End Set
    End Property
    Public Property Tax() As Double
        Get
            Return mTax
        End Get
        Set(ByVal value As Double)
            mTax = value
        End Set
    End Property
    Public Property TaxAmount() As Double
        Get
            Return mTaxAmount
        End Get
        Set(ByVal value As Double)
            mTaxAmount = value
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
    Public Property ReasonOfReturn() As String
        Get
            Return mReasonOfReturn
        End Get
        Set(ByVal value As String)
            mReasonOfReturn = value
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
    Public Property SupplierDONumber() As String
        Get
            Return mSupplierDONumber
        End Get
        Set(ByVal value As String)
            mSupplierDONumber = value
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

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
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

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    ''Di pakai sama AOP
    Public Sub IsiTBLPRByPrimaryGRN(ByVal ListGRN As List(Of Double), ByVal TBL As DseTransaksiPurchaseReturn.DataTable1DataTable)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable1TableAdapter
        Dim PrimaryGRN As Double = 0

        Try
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()
            For Each PrimaryGRN In ListGRN
                DTA.FillBy_PrimaryGRN(TBL, PrimaryGRN)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan tipe transaksi purchase return berdasarkan primary stock out purchase return
    ''' </summary>
    ''' <param name="vPrimaryStockOut"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTipePR_ByPrimaryStockOut(ByVal vPrimaryStockOut As Double) As EnumTipe_PurchaseReturn
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TipePR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Revisi_AOP

        Try
            If DTA.ScalarQuery_GetTipePR(vPrimaryStockOut).HasValue = True Then
                TipePR = DTA.ScalarQuery_GetTipePR(vPrimaryStockOut).Value
            Else
                TipePR = EnumTipe_PurchaseReturn.Revisi_AOP
            End If

            Return TipePR
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_getUrutanMax(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_getUrutanMax(TanggalStart, TanggalEnd).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 5 huruf
            While NoUrut.Length < 5
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/PRTN/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If DTA.ScalarQuery_getJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_getJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn1DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn1Row
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.Primary = .Primary
                    Me.NoTransaksi = .No_Transaksi
                    Me.PrimaryStockInPurchaseReturn = .Primary_Stock_Out_Purchase_Return
                    Me.PrimaryGRN = .Primary_GRN
                    Me.TanggalTransaksi = .Tanggal_Transaksi
                    Me.SubTotal = .Sub_Total
                    Me.TaxExclude = .Tax_Exclude
                    Me.Tax = .Tax
                    Me.TaxAmount = .Tax_Amount
                    Me.Total = .Total
                    Me.ReasonOfReturn = .Reason_Of_Return
                    Me.Remarks = .Remarks
                    Me.SupplierDONumber = .Supplier_DO_Number

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV_Multi(vDgvItem, Me.mPrimary)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function
    Public Function IsiDGV(ByVal vCriteria As String, _
       ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPurchaseReturn.DataTable3DataTable, _
       ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vStartDate As Date = (#1/1/1950#), Optional ByVal vEndDate As Date = (#1/1/2100#)) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vStartDate, vEndDate)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, vCriteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                QueryResult = DTA.FillBy_SupplierNameLike(TBL, vCriteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SupplierCodeLike(TBL, vCriteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                QueryResult = DTA.FillBy_noGRNLike(TBL, vCriteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                QueryResult = DTA.FillBy_NoStockInPurchaseLike(TBL, vCriteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns("Tanggal Transaksi").DefaultCellStyle = ij
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
    Public Overridable Function ApakahBolehEdit(Optional ByVal TampilkanPesan As Boolean = False) As Boolean
        Try

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(Optional ByVal TampilkanPesan As Boolean = False) As Boolean
        Try

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetPrimaryStockOutPurchaseReturn_Minus_Satu(ByVal vPrimaryStockOutPurchaseReturn As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            DTA.UpdateQuery_SetPrimaryStockOutPurchaseReturnMinus(vPrimaryStockOutPurchaseReturn)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetPrimaryGRN(ByVal vPrimaryGRNLama As Double, ByVal vPrimaryGRNBaru As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            DTA.UpdateQuery_SetPrimaryGRN(vPrimaryGRNBaru, vPrimaryGRNLama)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetPrimaryStockOutPurchaseReturn(ByVal vPrimaryPR As Double, ByVal vPrimaryStockOut As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            DTA.UpdateQuery_SetPrimaryStockOut(vPrimaryStockOut, vPrimaryPR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Obj As New ClsTransaksiPurchaseReturnItem
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Try
            If Me.ApakahBolehHapus = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function GetPrimary_ByPrimaryStockOutPR(ByVal vPrimaryStockOutPR As Double) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryPR(vPrimaryStockOutPR).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryPR(vPrimaryStockOutPR).Value
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Try
            If DTA.ScalarQuery_getPrimary(vNoTransaksi).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_getPrimary(vNoTransaksi).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetInfoSupplierNameNCodeNTransNo(ByVal vPrimary As Double, ByRef vSupName As String, _
    ByRef vSupCode As String, ByRef vNoPR As String, ByRef vSupplierDONO As String)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn2TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn2DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn2Row

        Try
            DTA.FillBy_PrimaryTransaksi(TBL, vPrimary)
            vSupName = ""
            vSupCode = ""
            vNoPR = ""
            vSupplierDONO = ""
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vSupName = row.SupplierName
                vSupCode = row.Supplier_No
                vNoPR = row.PR_No
                vSupplierDONO = row.Supplier_DO_Number
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini mengembalikan no transaksi stock out purchase return jika purchase return yang dimaksud sudah punya stock out purchase return
    ''' selain dari stock out purchase return yang dimaksud.
    ''' Jika mengambalikan string kosong maka purchase return tersebut bisa dirujuk oleh stock out purchase return tsb
    ''' </summary>
    ''' <param name="vPrimaryStockOut"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahBolehDiRujukStockOutPurchaseReturn(ByVal vPrimaryStockOut As Boolean, ByVal vPrimaryPurchaseReturn As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn1DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn1Row
        Dim Retval As String = ""

        Try
            DTA.FillBy_Primary(TBL, vPrimaryPurchaseReturn)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.Primary_Stock_Out_Purchase_Return = -1 Then
                    'purchase return ini belum pernah dirujuk oleh transaksi stock out purchase return mana pun
                    Return ""
                ElseIf row.Primary_Stock_Out_Purchase_Return <> vPrimaryStockOut Then
                    Dim ObjStockInOut As New clsTransaksiStockInOut
                    Retval = ObjStockInOut.GetNoTransaksi(row.Primary_Stock_Out_Purchase_Return)
                    ObjStockInOut = Nothing
                    Return Retval
                End If
            Else
                Return ""
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksi(vPrimaryTransaksi)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_getNoTransaksi(vPrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable31TableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.getJmlBy_NoTransaksiLike(Criteria + "%").HasValue = True Then
                    Return DTA.getJmlBy_NoTransaksiLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.GetJmlBy_TglDariSampai(vTanggalDari, vTanggalSampai).HasValue = True Then
                    Return DTA.GetJmlBy_TglDariSampai(vTanggalDari, vTanggalSampai).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                If DTA.getJmlBy_SupplierCodeLike(Criteria + "%").HasValue = True Then
                    Return DTA.getJmlBy_SupplierCodeLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                If DTA.GetJmlBy_SupplierNameLike(Criteria + "%").HasValue = True Then
                    Return DTA.GetJmlBy_SupplierNameLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                If DTA.GetJmlBy_noGRNLike(Criteria + "%").HasValue = True Then
                    Return DTA.GetJmlBy_SupplierNameLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                If DTA.GetJmlBy_NoStockInPurchaseLike(Criteria + "%").HasValue = True Then
                    Return DTA.GetJmlBy_NoStockInPurchaseLike(Criteria + "%")
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mNoUrut = -1
            mPrimaryStockInPurchaseReturn = -1
            mPrimaryGRN = -1
            mTanggalTransaksi = Now
            mSubTotal = 0
            mTaxExclude = True
            mTax = 0
            mTaxAmount = 0
            mTotal = 0
            mReasonOfReturn = ""
            mRemarks = ""
            mSupplierDONumber = ""
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJurnalNumber_ByStockOutPurchaseReturn(ByVal vPrimaryStockOutPurchaseReturn As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_jurnalDataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_jurnalRow
        Dim TipePayment As EnumTipePayment = EnumTipePayment.Purchase_Return
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic
        Dim retval As String = ""

        Try
            DTA.Fill(TBL, vPrimaryStockOutPurchaseReturn, TipePayment, TipeBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Jurnal_Number
                Else
                    retval = retval + ", " + row.Jurnal_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetAOPNumber_ByPrimaryPR(ByVal vprimaryPR As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAOPNumber(vprimaryPR)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetAOPNumber(vprimaryPR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    'mengembalikan nilai true jika
    Public Function ApakahDariQuickPR(ByVal vPrimaryStockOurPurchaseReturn As Double) As Boolean
        Dim PrimaryPR As Double = -1
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            PrimaryPR = Me.GetPrimary_ByPrimaryStockOutPR(vPrimaryStockOurPurchaseReturn)
            If DTA.ScalarQuery_GetQuickPR(PrimaryPR).HasValue = True Then
                If DTA.ScalarQuery_GetQuickPR(PrimaryPR).Value = 1 Then
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

    Public Function GetPrimaryGRN(ByVal vPrimaryPurchaseReturn As Double) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            RetVal = DTA.ScalarQuery_GetPrimaryGRN(vPrimaryPurchaseReturn)

            If RetVal.HasValue = False Then
                Return -1
            Else
                Return RetVal.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
