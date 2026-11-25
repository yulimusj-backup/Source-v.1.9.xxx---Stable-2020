Option Explicit On
Option Strict Off

Public Class clsTransaksiStockInOut
    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTipeTransaksi As EnumTipeStockInOut = EnumTipeStockInOut.StockIn_Adjust
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    Private mPrimaryPOPurchase As Double = -1
    Private mPrimaryDOSales As Double = -1
    Private mPrimaryCust As Double = -1
    Private mPrimarySupplier As Double = -1
    Private mImporDariPackingList As Boolean = False
    Private mPrimaryGIT As Double = -1
    Private mSupplierDoNo As String = ""
    Private mPrimaryAsset As Double = -1
    Private mInOutAdjust As EnumInOutAdjust = EnumInOutAdjust.NoTipe
    Private mPrimaryAdjType As Double = -1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim SudahHapusAntrian As Boolean = True

    'elemen yang digunakan pada automatic SR transfer
    Public Enum EnumTipePRBalance As Integer
        Automatic
        Manual
        Manual_Transfer '--> tidak digunakan lagi, hanya pada cmsi saja
        Automatic_SR
        Manual_SR
    End Enum

    'elemen yang digunakan pada module asset
    Public Enum EnumInOutAdjust As Integer
        NoTipe
        OutAdjustMaterialUsageAssetAcquisition
        OutAdjustMaterialUsageAssetRevaluation
        InOutCustom
    End Enum

    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum
    Public Enum Enum_Tipe_Setting
        None
        Debet
        Kredit
    End Enum
    Public Enum EnumTipe_PurchaseReturn As Integer
        Revisi_AOP
        Nota_Retur
    End Enum

    Public Property InOutAdjust() As EnumInOutAdjust
        Get
            Return mInOutAdjust
        End Get
        Set(ByVal value As EnumInOutAdjust)
            mInOutAdjust = value
        End Set
    End Property

    Public Property PrimaryAsset() As Double
        Get
            Return mPrimaryAsset
        End Get
        Set(ByVal value As Double)
            mPrimaryAsset = value
        End Set
    End Property

    Public Property SupplierDoNo() As String
        Get
            Return mSupplierDoNo
        End Get
        Set(ByVal value As String)
            mSupplierDoNo = value
        End Set
    End Property

    Public Property PrimaryGIT() As Double
        Get
            Return Me.mPrimaryGIT
        End Get
        Set(ByVal value As Double)
            mPrimaryGIT = value
        End Set
    End Property

    Public Property ImporDariPackingList() As Boolean
        Get
            Return Me.mImporDariPackingList
        End Get
        Set(ByVal value As Boolean)
            mImporDariPackingList = value
        End Set
    End Property

    Public Property PrimarySupplier() As Double
        Get
            Return Me.mPrimarySupplier
        End Get
        Set(ByVal value As Double)
            mPrimarySupplier = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return Me.mPrimaryCust
        End Get
        Set(ByVal value As Double)
            mPrimaryCust = value
        End Set
    End Property

    Public Property PrimaryDOSales() As Double
        Get
            Return Me.mPrimaryDOSales
        End Get
        Set(ByVal value As Double)
            mPrimaryDOSales = value
        End Set
    End Property

    Public Property PrimaryPOPurchase() As Double
        Get
            Return Me.mPrimaryPOPurchase
        End Get
        Set(ByVal value As Double)
            mPrimaryPOPurchase = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property TipeTransaksi() As EnumTipeStockInOut
        Get
            Return Me.mTipeTransaksi
        End Get
        Set(ByVal value As EnumTipeStockInOut)
            mTipeTransaksi = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
        End Set
    End Property

    Public Enum EnumTipeStockInOut As Integer
        StockIn_Purchase
        StockIn_SalesReturn
        StockIn_Adjust
        StockOut_Sales
        StockOut_PurchaseReturn
        StockOut_Adjust
        StockIn_Process
        StockIn_BackProcess
        StockOut_Process
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
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

    Public Property PrimaryAdjType() As Double
        Get
            Return Me.mPrimaryAdjType
        End Get
        Set(ByVal value As Double)
            mPrimaryAdjType = value
        End Set
    End Property

    Public Function GetNoBaru(ByVal vJenisTransaksi As EnumTipeStockInOut, ByVal vTanggalTransaksi As Date, _
    Optional ByVal vJenisInOutAdjust As EnumInOutAdjust = EnumInOutAdjust.NoTipe) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, InOut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'in/out & Tipe Transaksi
            InOut = ""
            If vJenisTransaksi = EnumTipeStockInOut.StockIn_Adjust Or vJenisTransaksi = EnumTipeStockInOut.StockIn_BackProcess Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_Process Or vJenisTransaksi = EnumTipeStockInOut.StockIn_Purchase Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                InOut = "IN"
            Else
                If vJenisTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                    If vJenisInOutAdjust = EnumInOutAdjust.NoTipe Then
                        InOut = "OUT"
                    ElseIf vJenisInOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                        InOut = "ASSET-ACQUISITION"
                    ElseIf vJenisInOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                        InOut = "ASSET-REVALUATION"
                    End If
                Else
                    InOut = "OUT"
                End If
            End If

            'cari no urut
            If InOut = "IN" Then
                If DTA.ScalarQuery_GetNoUrutIn(TanggalEnd, TanggalStart).HasValue = True Then
                    NoUrut = CStr(DTA.ScalarQuery_GetNoUrutIn(TanggalEnd, TanggalStart).Value + 1)
                    If NoUrut <= 0 Then
                        NoUrut = 1
                    End If
                Else
                    NoUrut = 1
                End If
            Else
                If DTA.ScalarQuery_GetNoUrutOut(TanggalEnd, TanggalStart).HasValue = True Then
                    NoUrut = CStr(DTA.ScalarQuery_GetNoUrutOut(TanggalEnd, TanggalStart).Value + 1)
                    If NoUrut <= 0 Then
                        NoUrut = 1
                    End If
                Else
                    NoUrut = 1
                End If
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
            Return NoUrut + "/" + InOut + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    '11 Juli 2019 /untuk kebutuhan numbering di quick invoice terhadap numbering out sales
    Public Function GetNoBaru_20190711(ByVal vJenisTransaksi As EnumTipeStockInOut, ByVal vTanggalTransaksi As Date, _
    Optional ByVal vJenisInOutAdjust As EnumInOutAdjust = EnumInOutAdjust.NoTipe, Optional ByVal vUrutan As Integer = -1) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, InOut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'in/out & Tipe Transaksi
            InOut = ""
            If vJenisTransaksi = EnumTipeStockInOut.StockIn_Adjust Or vJenisTransaksi = EnumTipeStockInOut.StockIn_BackProcess Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_Process Or vJenisTransaksi = EnumTipeStockInOut.StockIn_Purchase Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                InOut = "IN"
            Else
                If vJenisTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                    If vJenisInOutAdjust = EnumInOutAdjust.NoTipe Then
                        InOut = "OUT"
                    ElseIf vJenisInOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                        InOut = "ASSET-ACQUISITION"
                    ElseIf vJenisInOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                        InOut = "ASSET-REVALUATION"
                    End If
                Else
                    InOut = "OUT"
                End If
            End If

            'cari no urut
            If vUrutan = -1 Then
                If InOut = "IN" Then
                    If DTA.ScalarQuery_GetNoUrutIn(TanggalEnd, TanggalStart).HasValue = True Then
                        NoUrut = CStr(DTA.ScalarQuery_GetNoUrutIn(TanggalEnd, TanggalStart).Value + 1)
                        If NoUrut <= 0 Then
                            NoUrut = 1
                        End If
                    Else
                        NoUrut = 1
                    End If
                Else
                    If DTA.ScalarQuery_GetNoUrutOut(TanggalEnd, TanggalStart).HasValue = True Then
                        NoUrut = CStr(DTA.ScalarQuery_GetNoUrutOut(TanggalEnd, TanggalStart).Value + 1)
                        If NoUrut <= 0 Then
                            NoUrut = 1
                        End If
                    Else
                        NoUrut = 1
                    End If
                End If
            Else
                If vUrutan < 1 Then
                    NoUrut = 0
                Else
                    NoUrut = vUrutan
                End If
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
            Return NoUrut + "/" + InOut + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail2_(ByVal vPrimaryTransaksi As Double, _
       ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinoutDataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinoutRow
        Dim ObjItem As New clsTransaksiStockInOutItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mImporDariPackingList = .Impor_Dari_Packing_List
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mPrimaryDOSales = .Primary_DO_Sales
                    Me.mPrimaryGIT = .Primary_GIT
                    Me.mPrimaryPOPurchase = .Primary_PO_Purchase
                    Me.mPrimarySupplier = .Primary_Supplier
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mTipeTransaksi = .Tipe_Transaksi

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mSupplierDoNo = ._Supplier_s_DO_Number
                End With
                'salin data item 
                ObjItem.IsiDGV2_(vDgvItem, Me.mPrimary, Me.mTipeTransaksi)
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vTbl As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinoutDataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinoutRow
        Dim ObjItem As New clsTransaksiStockInOutItem
        Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mImporDariPackingList = .Impor_Dari_Packing_List
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mPrimaryDOSales = .Primary_DO_Sales
                    Me.mPrimaryGIT = .Primary_GIT
                    Me.mPrimaryPOPurchase = .Primary_PO_Purchase
                    Me.mPrimarySupplier = .Primary_Supplier
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mTipeTransaksi = .Tipe_Transaksi
                    Me.mPrimaryAdjType = .PrimaryAdjType

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mSupplierDoNo = ._Supplier_s_DO_Number
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, Me.mTipeTransaksi, vTbl)

                Dim JenisTransAsalCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal
                If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Sales
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_SalesReturn
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Adjust
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                    JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Adjust
                End If

                ObjCustInOut.GetStockInOutItem_Custom(vTbl, vPrimaryTransaksi, True, JenisTransAsalCustom)

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

    Public Function GetCreatedBy(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True, _
    Optional ByVal vStrictEdit As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim ObjPR As New ClsTransaksiPurchaseReturn
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable10TableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable10DataTable
        Dim row As dseTransaksiStockInOut.DataTable10Row
        Dim ProductCode As String = ""
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockInOut
        Try
            If Me.GetNoTransaksi(Me.mPrimary) = "" Then
                MsgBox("Data not found, other user may have delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'msgbox("belum")
            'jika sudah ada product yang baru saja berubah kontrolnya sehingga tanggal minimum juga berubah
            'transaksi yang tanggal kontrolnya lebih kecil dari tanggal minimum tidak bisa diedit
            DTA.Fill(TBL, Me.mPrimary)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If ProductCode = "" Then
                        ProductCode = row.Product_Code
                    Else
                        ProductCode = ProductCode + ", " + row.Product_Code
                    End If
                Next

                MsgBox("Code no " + ProductCode + " have modified it's control state. You cannot edit this data", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'Suatu transaksi in/out tidak bisa diedit jika:
            '0. cek yang berkaitan dengan melakukan proses edit

            '1. sudah ada transaksi lanjutannya.
            '2. transaksi sudah dikunci oleh user
            '3. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            '4. ?
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be editted by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            If Me.TipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                'transaksi lanjutan tidak ada

            ElseIf Me.TipeTransaksi = EnumTipeStockInOut.StockIn_BackProcess Then
                'transaksi lanjutnya tidak ada

            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Process Then
                'transaksi lanjutnya tidak ada

            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                'transaksi lanjutnya adalah AOP
                Dim objgrn As New ClsTransaksiGRN
                'kalo sudah ada AOP yang merujuk ke GRN yang dirujuk oleg stock in purchase
                Dim NoAOP As String = objgrn.GetNoAOP(Me.mPrimary)
                If NoAOP <> "" Then
                    MsgBox("You cannot edit this data, AOP: " + NoAOP + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'kalo sudah ada stock out purchase return yang merujuk pada transaksi ini
                'maka data stock in purchase tidak boleh di hapus
                'kalo sudah ada AOP yang merujuk ke GRN yang dirujuk oleg stock in purchase
                NoAOP = objgrn.GetNoPurchaseReturn(Me.mPrimary)
                If NoAOP <> "" Then
                    MsgBox("You cannot edit this data, purchase return: " + NoAOP + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'maka tidak bisa di edit
                objgrn = Nothing

                If ObjCustom.ApakahBolehEdit_CustomStockInPurchase(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                'cek apakah dibuat berdasarkan pick up order
                Dim DTAPU As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                Dim DummyNoPickUpOrder As String = ""
                Dim RetVal As Boolean = True

                DummyNoPickUpOrder = DTAPU.ScalarQuery_GetPirckUpOrderNo(Me.mPrimary)
                If String.IsNullOrEmpty(DummyNoPickUpOrder) = False Then
                    RetVal = False
                End If
                DTAPU.Dispose()

                If RetVal = False Then
                    If TampilkanPesan = True Then
                        MsgBox("This data is refered by pick up order: " + DummyNoPickUpOrder + _
                        ", please edit this data from pick up order form.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If

                If ObjCustom.ApakahBolehEdit_CustomStockOutAdjust(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Process Then
                'transaksi lanjutnya tidak ada

            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                TipePR = ObjPR.GetTipePR_ByPrimaryStockOut(Me.mPrimary)
                If TipePR = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP Then
                    Dim NoAOP As String = ""
                    'jika purchase return yang dirujuk sudah dibuatkan AOP, maka stock out purchase return tidak bisa dihapuskan
                    NoAOP = Me.GetNoAOP_ByPrimaryStockOutPurchaseReturn(Me.mPrimary)
                    If NoAOP <> "" Then
                        If TampilkanPesan = True Then
                            MsgBox("You cannot edit this data, AOP: " + NoAOP + _
                            " is refering to this transaction", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                Else
                    'jika sudah ada purchase return receivement untuk purchase return ini, 
                    'maka stock out tidak bisa diedit
                    Dim Jurnal As String = ""
                    Jurnal = ObjPR.GetJurnalNumber_ByStockOutPurchaseReturn(Me.mPrimary)
                    If Jurnal <> "" Then
                        If TampilkanPesan = True Then
                            MsgBox("You cannot edit this data, purchase receivement jurnal: " + Jurnal + _
                            " is refering to this transaction", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                'transaksi lanjutnya adalah invoice yang merujuk pada DO yang dirujuk oleh 
                'transaksi stock out ini
                Dim ObjDO1 As New ClsTransaksiDO
                If ObjDO1.ApakahSudahAdaInvoice(Me.mPrimaryDOSales) = True Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot edit this data, because delivery order refered by this transaction is already invoiced.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    ObjDO1 = Nothing
                    Return False
                ElseIf ObjDO1.ApakahSudahAdaSalesReturn(Me.mPrimaryDOSales) = True Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot edit this data, because there are sales return refering to this transaction.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    ObjDO1 = Nothing
                    Return False
                End If
                ObjDO1 = Nothing

                If ObjCustom.ApakahBolehEdit_CustomStockOutSales(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                'tidak bisa dihapus jika:
                '1. sudah ada sales return yang merujuk ke transaksi ini
                Dim ObjSR As New ClsTransaksiSalesReturn
                Dim NoSR As String = ObjSR.GetNoSalesReturn(Me.mPrimary)
                If NoSR <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot edit this data, sales return: " + NoSR + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
                ObjSR = Nothing
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            ObjPR = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True, _
    Optional ByVal vStrictEdit As Boolean = True, Optional ByVal vFromPickUpOrder As Boolean = False, Optional ByVal vCekStockMinus As Boolean = True) As Boolean
        Dim InvoiceNo As String = ""
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim NoAOP As String = ""
        Dim ObjPR As New ClsTransaksiPurchaseReturn
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable10TableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable10DataTable
        Dim row As dseTransaksiStockInOut.DataTable10Row
        Dim ProductCode As String = ""
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockInOut
        Try
            If Me.GetNoTransaksi(Me.mPrimary) = "" Then
                MsgBox("Data not found, other user may have delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'msgbox("belum")
            'jika sudah ada product yang baru saja berubah kontrolnya sehingga tanggal minimum juga berubah
            'transaksi yang tanggal kontrolnya lebih kecil dari tanggal minimum tidak bisa diedit
            DTA.Fill(TBL, Me.mPrimary)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If ProductCode = "" Then
                        ProductCode = row.Product_Code
                    Else
                        ProductCode = ProductCode + ", " + row.Product_Code
                    End If
                Next

                MsgBox("Code no " + ProductCode + " have modified it's control state. You cannot delete this data", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If
            'transaksi tidak bisa dihapus bila:
            '0. strict deleting
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be deleted by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            '1. transaksi sudah dikunci oleh user
            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            '3. jika menghapus transaksi ini akan mengakibatkan data product menjadi minus
            If vCekStockMinus = True Then
                Dim ObjItem As New clsTransaksiStockInOutItem
                If ObjItem.ApakahHapusMengakibatkanMinus(Me.mPrimary, Me.mTipeTransaksi, TampilkanPesan, Me.vPrimary_Antrian, _
                Me.DTAInOUT, Me.SudahHapusAntrian) = True Then
                    ObjItem = Nothing
                    Return False
                End If
                ObjItem = Nothing
            End If

            '4. Jika sudah ada transaksi lanjutan
            If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                'transaksi lanjutnya adalah invoice yang merujuk pada DO yang dirujuk oleh 
                'transaksi stock out ini
                Dim ObjDO As New ClsTransaksiDO
                If ObjDO.ApakahSudahAdaInvoice(Me.mPrimaryDOSales) = True Then
                    If TampilkanPesan = True Then
                        InvoiceNo = ObjDO.GetInvoiceNo(Me.mPrimaryDOSales)
                        MsgBox("You cannot delete this data, because delivery order refered by this transaction is already invoiced. The invoice number is " + InvoiceNo, _
                        MsgBoxStyle.Information, "Information")
                    End If

                    ObjDO = Nothing
                    Return False
                ElseIf ObjDO.ApakahSudahAdaSalesReturn(Me.mPrimaryDOSales) = True Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot delete this data, because there are sales return refering to the delivery order refered by this transaction.", _
                        MsgBoxStyle.Information, "Information")
                    End If

                    ObjDO = Nothing
                    Return False
                End If
                ObjDO = Nothing
                'jika delivery order yang dirujuk oleh stock out sales sudah ada return
                Dim NoTransaksiStockInSalesReturn As String = Me.GetStockInSalesReturnNo(Me.mPrimaryDOSales)
                If NoTransaksiStockInSalesReturn <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot delete this data, because the stock in sales return:" + NoTransaksiStockInSalesReturn _
                        + " refer to the delivery order refered by this transaction.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If

                If ObjCustom.ApakahBolehHapus_CustomStockOutSales(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                'tidak bisa dihapus jika:
                '1. sudah ada sales return yang merujuk ke transaksi ini
                Dim ObjSR As New ClsTransaksiSalesReturn
                Dim NoSR As String = ObjSR.GetNoSalesReturn(Me.mPrimary)
                If NoSR <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot delete this data, sales return: " + NoSR + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
                ObjSR = Nothing
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                Dim objgrn As New ClsTransaksiGRN
                'kalo sudah ada AOP yang merujuk ke GRN yang dirujuk oleg stock in purchase
                NoAOP = objgrn.GetNoAOP(Me.mPrimary)
                If NoAOP <> "" Then
                    MsgBox("You cannot delete this data, AOP: " + NoAOP + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'kalo sudah ada stock out purchase return yang merujuk pada transaksi ini
                'maka data stock in purchase tidak boleh di hapus
                'kalo sudah ada AOP yang merujuk ke GRN yang dirujuk oleg stock in purchase
                NoAOP = objgrn.GetNoPurchaseReturn(Me.mPrimary)
                If NoAOP <> "" Then
                    MsgBox("You cannot delete this data, purchase return: " + NoAOP + _
                        " is refering to this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'maka tidak bisa hapus
                objgrn = Nothing

                If ObjCustom.ApakahBolehHapus_CustomStockInPurchase(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                TipePR = ObjPR.GetTipePR_ByPrimaryStockOut(Me.mPrimary)
                If TipePR = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP Then
                    NoAOP = ""
                    'jika purchase return yang dirujuk sudah dibuatkan AOP, maka stock out purchase return tidak bisa dihapuskan
                    NoAOP = Me.GetNoAOP_ByPrimaryStockOutPurchaseReturn(Me.mPrimary)
                    If NoAOP <> "" Then
                        If TampilkanPesan = True Then
                            MsgBox("You cannot edit this data, AOP: " + NoAOP + _
                            " is refering to this transaction", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                Else
                    'jika sudah ada purchase return receivement untuk purchase return ini, 
                    'maka stock out tidak bisa diedit
                    Dim Jurnal As String = ""
                    Jurnal = ObjPR.GetJurnalNumber_ByStockOutPurchaseReturn(Me.mPrimary)
                    If Jurnal <> "" Then
                        If TampilkanPesan = True Then
                            MsgBox("You cannot delete this data, purchase receivement jurnal: " + Jurnal + _
                            " is refering to this transaction", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                If vFromPickUpOrder = False Then
                    'cek apakah dibuat berdasarkan pick up order
                    Dim DTAPU As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                    Dim DummyNoPickUpOrder As String = ""
                    Dim RetVal As Boolean = True

                    DummyNoPickUpOrder = DTAPU.ScalarQuery_GetPirckUpOrderNo(Me.mPrimary)
                    If String.IsNullOrEmpty(DummyNoPickUpOrder) = False Then
                        RetVal = False
                    End If
                    DTAPU.Dispose()

                    If RetVal = False Then
                        If TampilkanPesan = True Then
                            MsgBox("This data is refered by pick up order: " + DummyNoPickUpOrder + _
                            ", please edit this data from pick up order form.", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If

                If ObjCustom.ApakahBolehHapus_CustomStockOutAdjust(Me.mPrimary, TampilkanPesan) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            ObjPR = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan no AOP yang merujuk stock out purchase return ybs, dengan memberikan primary stock 
    ''' out purchase return yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryStockOut"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoAOP_ByPrimaryStockOutPurchaseReturn(ByVal vPrimaryStockOut As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable9TableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable9DataTable
        Dim row As dseTransaksiStockInOut.DataTable9Row
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP
        Dim retVal As String = ""

        Try
            DTA.Fill(TBL, vPrimaryStockOut, TipePR)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If retVal = "" Then
                        retVal = row.AOP_Number
                    Else
                        retVal = retVal + ", " + row.AOP_Number
                    End If
                Next

                Return retVal
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahBeginingBalance(ByVal vPrimary As Double, ByVal vConString As String) As Boolean
        Dim Dummy As Nullable(Of Integer)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try

            ''@rosadinnn
            If String.IsNullOrEmpty(vConString) = False Then
                DTA.Connection.ConnectionString = vConString
            End If

            Dummy = DTA.ScalarQuery_GetIsBeginingBalance(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBeginingBalance(ByVal vPrimary As Double) As Boolean
        Dim Dummy As Nullable(Of Integer)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            Dummy = DTA.ScalarQuery_GetIsBeginingBalance(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
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

    Public Overridable Function ApakahDataBenar_Default(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, _
    ByVal vTipeTransaksiStockInOut As EnumTipeStockInOut, Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New dseTransaksiStockInOut.tblmaster_productDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim RetVal As Boolean = True
        Dim DummyMsg As String = ""
        Dim QtyDummy As Decimal = 0

        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            '2. Jika semua transaksi inventory sedang dikunci

            '3. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif
            'yang disimulasikan hanya item goods saja
            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.FillBy_Primary_Simulasi(TBL, Me.mPrimary, CInt(TipeItemGoods))
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            DTAService.Fill(TBLService, TipeItemService)
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(15).Value
                'jika service maka tidak masuk
                If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                    BatchNo = rowDGV.Cells(9).Value
                    PrimaryLocation = rowDGV.Cells(16).Value
                    If IsNumeric(rowDGV.Cells(7).Value) = True Then
                        Sec_Qty = rowDGV.Cells(7).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(10).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(10).Value)
                    End If
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                End If
            Next

            If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Or Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Process Or Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                'simulasikan untuk stock in 
                'simulasi hapus data item lama
                If TBL.Rows.Count > 0 Then
                    For Each rowInOutLama In TBL.Rows
                        With rowInOutLama
                            Me.Koreksi4HapusIn(TBLMasterProductQty, _
                            .Primary_Product, .Qty, .Secondary_Qty, _
                            .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                        End With
                    Next
                End If
                'simulasi simpan item baru
                For Each rowDGV In vDGVItem.Rows
                    BatchNo = rowDGV.Cells(9).Value
                    PrimaryLocation = rowDGV.Cells(16).Value
                    If IsNumeric(rowDGV.Cells(7).Value) = True Then
                        Sec_Qty = rowDGV.Cells(7).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(10).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(10).Value)
                    End If
                    PrimaryProduct = rowDGV.Cells(15).Value
                    Qty = rowDGV.Cells(5).Value

                    Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                    PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                    BatchNo, -1, ExpiryDate)
                Next
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Process Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                'simulasikan untuk stock out
                'simulasi hapus data item lama
                If TBL.Rows.Count > 0 Then
                    For Each rowInOutLama In TBL.Rows
                        With rowInOutLama
                            Me.Koreksi4HapusOut(TBLMasterProductQty, _
                            .Primary_Product, .Qty, .Secondary_Qty, _
                            .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                        End With
                    Next
                End If
                'simulasi simpan item baru
                For Each rowDGV In vDGVItem.Rows
                    BatchNo = rowDGV.Cells(9).Value
                    PrimaryLocation = rowDGV.Cells(16).Value
                    If IsNumeric(rowDGV.Cells(7).Value) = True Then
                        Sec_Qty = rowDGV.Cells(7).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(10).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(10).Value)
                    End If
                    PrimaryProduct = rowDGV.Cells(15).Value
                    Qty = rowDGV.Cells(5).Value

                    Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                    PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                    BatchNo, ExpiryDate)
                Next
            End If

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While

                If Me.SudahHapusAntrian = False Then
                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                        MsgBox("Logic error(1), contact your provider when you see this message", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.SudahHapusAntrian = True
                End If

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            If vTipeTransaksiStockInOut = EnumTipeStockInOut.StockOut_Adjust Then
                If vPrimaryPickUpOrder <> -1 Then
                    Dim DTAPickUP As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_item1TableAdapter
                    Dim TBLPickUp As New DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1DataTable
                    Dim rowPickUp As DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1Row
                    Dim vFrm As New FrmTransaksiStockOutAdj

                    ListPrimaryProduct.Clear()

                    DTAPickUP.Fill(TBLPickUp, vPrimaryPickUpOrder)
                    For Each rowDGV In vDGVItem.Rows
                        TBLPickUp.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value.ToString
                        QtyDummy = rowDGV.Cells(vFrm.ColQty.Index).Value
                        While QtyDummy > 0 And TBLPickUp.DefaultView.Count > 0
                            rowPickUp = TBLPickUp.DefaultView(0).Row
                            If rowPickUp.qty <= QtyDummy Then
                                'qty pick up <= qty dgv
                                QtyDummy -= rowPickUp.qty
                                TBLPickUp.Rows.Remove(rowPickUp)
                            Else
                                rowPickUp.qty -= QtyDummy
                                QtyDummy = 0
                            End If
                        End While
                        If QtyDummy > 0 Then
                            'qty dg ada yg sisa dan tidak ada di item pick up order
                            If ListPrimaryProduct.Contains(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value) = False Then
                                ListPrimaryProduct.Enqueue(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                                If DummyMsg = "" Then
                                    DummyMsg = rowDGV.Cells(vFrm.ColCodeNo.Index).Value
                                Else
                                    DummyMsg = DummyMsg + ", " + rowDGV.Cells(vFrm.ColCodeNo.Index).Value
                                End If
                            End If
                        End If
                    Next
                    If TBLPickUp.Rows.Count > 0 Then
                        'item pick up order ada yang tidak di masukkan ke dalam dgv
                        For Each rowPickUp In TBLPickUp.Rows
                            If ListPrimaryProduct.Contains(rowPickUp.primary_code_no) = False Then
                                ListPrimaryProduct.Enqueue(rowPickUp.primary_code_no)
                                If DummyMsg = "" Then
                                    DummyMsg = rowPickUp.Product_Code
                                Else
                                    DummyMsg = DummyMsg + ", " + rowPickUp.Product_Code
                                End If
                            End If
                        Next
                    End If
                    DTAPickUP.Dispose()
                    TBLPickUp.Dispose()
                    vFrm.Dispose()

                    If ListPrimaryProduct.Count > 0 Then
                        If Me.SudahHapusAntrian = False Then
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Me.SudahHapusAntrian = True
                        End If

                        MsgBox("Inconsistent data with pick up order for the following product: " + DummyMsg + ". Please fix first.", _
                        MsgBoxStyle.Information, "Information")
                        RetVal = False
                    Else
                        RetVal = True
                    End If

                    If RetVal = True Then
                        Dim DTAPickUp1 As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                        Dim PrimaryStockOutAdjustDummy As Double = -1
                        Dim ObjPR As New ClsTransaksiPR_PickUpOrder

                        If DTAPickUp1.ScalarQuery_GetPrimaryStockOutAdjust(vPrimaryPickUpOrder).HasValue = True Then
                            PrimaryStockOutAdjustDummy = DTAPickUp1.ScalarQuery_GetPrimaryStockOutAdjust(vPrimaryPickUpOrder)
                        Else
                            PrimaryStockOutAdjustDummy = -1
                        End If
                        If vBaru = True Then
                            'pastikan pick up order tidak memiliki stock out adjust yg merujuk, jika baru
                            If PrimaryStockOutAdjustDummy <> -1 Then
                                If Me.SudahHapusAntrian = False Then
                                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                        MsgBox("Logic error(1), contact your provider when you see this message", _
                                        MsgBoxStyle.Information, "Information")
                                    End If
                                    Me.SudahHapusAntrian = True
                                End If

                                MsgBox("The pick up order: " + ObjPR.GetNoTransaksi(vPrimaryPickUpOrder) + " is refered by stock out adjust: " + Me.GetNoTransaksi(PrimaryStockOutAdjustDummy) + _
                                ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                                RetVal = False
                            End If
                        Else
                            'pastikan pick up order merujuk pada primary stock out ybs, jika simpan perubahan
                            If PrimaryStockOutAdjustDummy <> Me.mPrimary Then
                                If Me.SudahHapusAntrian = False Then
                                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                        MsgBox("Logic error(1), contact your provider when you see this message", _
                                        MsgBoxStyle.Information, "Information")
                                    End If
                                    Me.SudahHapusAntrian = True
                                End If

                                MsgBox("The pick up order: " + ObjPR.GetNoTransaksi(vPrimaryPickUpOrder) + " is refered by stock out adjust: " + Me.GetNoTransaksi(PrimaryStockOutAdjustDummy) + _
                                ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                                RetVal = False
                            End If
                        End If
                        DTAPickUp1.Dispose()
                        ObjPR = Nothing
                    End If

                    If RetVal = False Then
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
            ObjProduct = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar_StockInSalesReturn(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New dseTransaksiStockInOut.tblmaster_productDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim RetVal As Boolean = True
        Dim DummyMsg As String = ""
        Dim QtyDummy As Decimal = 0
        Dim FrmRef As New FrmTransaksiStockInSalesReturn2

        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            '2. Jika semua transaksi inventory sedang dikunci

            '3. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif
            'yang disimulasikan hanya item goods saja
            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.FillBy_Primary_Simulasi(TBL, Me.mPrimary, CInt(TipeItemGoods))
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            DTAService.Fill(TBLService, TipeItemService)
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(FrmRef.ColPrimaryCodeNo.Index).Value
                'jika service maka tidak masuk
                If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                    BatchNo = rowDGV.Cells(FrmRef.ColBatch.Index).Value
                    PrimaryLocation = rowDGV.Cells(FrmRef.ColPrimaryLocation.Index).Value
                    Sec_Qty = "-1"
                    If IsDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value)
                    End If
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                End If
            Next

            'simulasikan untuk stock in 
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusIn(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(FrmRef.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(FrmRef.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(7).Value) = True Then
                    Sec_Qty = rowDGV.Cells(7).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(FrmRef.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(5).Value

                Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, -1, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While

                If Me.SudahHapusAntrian = False Then
                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                        MsgBox("Logic error(1), contact your provider when you see this message", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.SudahHapusAntrian = True
                End If

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
            ObjProduct = Nothing
            FrmRef = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar_StockOutSales(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New dseTransaksiStockInOut.tblmaster_productDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim RetVal As Boolean = True
        Dim DummyMsg As String = ""
        Dim QtyDummy As Decimal = 0
        Dim FrmRef As New FrmTransaksiStockOutSales
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting

        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            '2. Jika semua transaksi inventory sedang dikunci

            '3. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif
            'yang disimulasikan hanya item goods saja
            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.FillBy_Primary_Simulasi(TBL, Me.mPrimary, CInt(TipeItemGoods))
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            DTAService.Fill(TBLService, TipeItemService)
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(FrmRef.ColPrimaryCodeNo.Index).Value
                'jika service maka tidak masuk
                If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                    BatchNo = rowDGV.Cells(FrmRef.ColBatch.Index).Value
                    PrimaryLocation = rowDGV.Cells(FrmRef.ColPrimaryLocation.Index).Value
                    Sec_Qty = "-1"
                    If IsDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value)
                    End If
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                End If
            Next

            'simulasikan untuk stock out
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusOut(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If
            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(FrmRef.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(FrmRef.ColPrimaryLocation.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(FrmRef.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(FrmRef.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(FrmRef.ColQty.Index).Value

                Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        If ListPrimaryProduct.Contains(rowProduct.Primary_Product) = False Then
                            ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                        End If
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                For Each PrimaryProduct In ListPrimaryProduct
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(PrimaryProduct)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(PrimaryProduct)
                    End If
                Next

                If Me.SudahHapusAntrian = False Then
                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                        MsgBox("Logic error(1), contact your provider when you see this message", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.SudahHapusAntrian = True
                End If

                Dim wd As New FrmTransaksiStockOutSales_QtyMinus
                wd.FrmRef = FrmRef
                wd.TBLMasterProductQty = TBLMasterProductQty
                wd.ListPrimaryProduct = ListPrimaryProduct
                ' wd.ObjProduct = ObjProduct
                wd.DGVItem = vDGVItem 
                wd.ShowDialog()

                Return False
            End If

            'jika boleh override qty pada stock out sales 
            'disini harus cek ulang sales order balance multinya
            'disini, nilai sales order balance, harus consider upper tolerance
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Stock_Out_Sales_AllowOverrideDO) = "1" Then
                'ambil data kombinasi primary so dan primary product, berdasarkan do

                'ambil data kombinasi primary so dan primary product, berdasarkan tabel stock in out item di db

                'ambil data kombinasi primary so dan primary product, berdasarkan dgv

            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
            ObjProduct = Nothing
            FrmRef = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Overridable Function ApakahDataBenar_StockInPurchase(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView) As Boolean
        Dim vFrmRef As New FrmTransaksiStockInPurchase2

        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New dseTransaksiStockInOut.tblmaster_productDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            '2. Jika semua transaksi inventory sedang dikunci

            '3. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif
            'yang disimulasikan hanya item goods saja
            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.FillBy_Primary_Simulasi(TBL, Me.mPrimary, CInt(TipeItemGoods))
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            DTAService.Fill(TBLService, TipeItemService)
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(vFrmRef.ColPrimaryCodeNo.Index).Value
                'jika service maka tidak masuk
                If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                    BatchNo = rowDGV.Cells(vFrmRef.ColBatch.Index).Value
                    PrimaryLocation = rowDGV.Cells(vFrmRef.ColPrimaryLocation.Index).Value
                    If IsNumeric(rowDGV.Cells(vFrmRef.ColSecondaryQty.Index).Value) = True Then
                        Sec_Qty = rowDGV.Cells(vFrmRef.ColSecondaryQty.Index).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(vFrmRef.ColExpiredDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(vFrmRef.ColExpiredDate.Index).Value)
                    End If
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                End If
            Next

            'simulasikan untuk stock in 
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusIn(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                    End With
                Next
            End If
            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(vFrmRef.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrmRef.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(vFrmRef.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = rowDGV.Cells(vFrmRef.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(vFrmRef.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrmRef.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrmRef.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(vFrmRef.ColQty.Index).Value

                Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, -1, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While

                If Me.SudahHapusAntrian = False Then
                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                        MsgBox("Logic error(1), contact your provider when you see this message", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.SudahHapusAntrian = True
                End If

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
            ObjProduct = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar_StockOutPurchaseReturn(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New dseTransaksiStockInOut.tblmaster_productDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Dim frmRef As New FrmTransaksiStockOutPurchaseReturn2

        Try
            'cocokan data purchase return-nya


            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.FillBy_Primary_Simulasi(TBL, Me.mPrimary, CInt(TipeItemGoods))
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            DTAService.Fill(TBLService, TipeItemService)
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(frmRef.ColPrimaryCodeNo.Index).Value
                'jika service maka tidak masuk
                If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                    BatchNo = rowDGV.Cells(frmRef.ColBatch.Index).Value
                    PrimaryLocation = rowDGV.Cells(frmRef.ColPrimaryLocation.Index).Value
                    If IsNumeric(rowDGV.Cells(frmRef.ColSecondaryQty.Index).Value) = True Then
                        Sec_Qty = rowDGV.Cells(frmRef.ColSecondaryQty.Index).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(frmRef.ColExpiredDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(frmRef.ColExpiredDate.Index).Value)
                    End If
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                End If
            Next

            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusOut(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If
            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(frmRef.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(frmRef.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(frmRef.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = rowDGV.Cells(frmRef.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(frmRef.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(frmRef.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(frmRef.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(frmRef.ColQty.Index).Value

                Me.Koreksi4SimpanOut(TBLMasterProductQty, PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While

                If Me.SudahHapusAntrian = False Then
                    If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                        MsgBox("Logic error(1), contact your provider when you see this message", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.SudahHapusAntrian = True
                End If

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If



            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
            ObjProduct = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
            ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
            Optional ByVal vAskEmailNotification As Boolean = True, Optional ByVal vSetPOB As Boolean = True, _
            Optional ByVal vPrimaryPickUpOrder As Double = -1, Optional ByVal vCekDataBenar As Boolean = True, _
            Optional ByVal vKoreksiProdQty As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean

        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim ObjBooking As New ClsTransaksiBookingInv

        Dim ObjItem As New clsTransaksiStockInOutItem
        Dim CustName, SuppName, CustCode As String
        Dim rowDGV As DataGridViewRow
        Dim BeginingBalance As Boolean = False
        Dim Lanjut As Boolean = False
        Dim PrimaryPRLama As Double = -1

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut
        Dim ListProductCustom As New List(Of Double)
        Dim JenisTransAsalCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal
        Dim JenisStockCardCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard
        Dim DescSC As String = ""

        Try
            If Me.mPrimaryCust <> -1 Then
                Dim ObjCust As New Core_Module.clsMasterCustomer
                CustName = ObjCust.GetCustName(Me.mPrimaryCust)
                CustCode = ObjCust.GetCustomerCode(Me.mPrimaryCust)
                ObjCust = Nothing
            Else
                CustName = ""
                CustCode = ""
            End If
            If Me.mPrimarySupplier <> -1 Then
                Dim ObjSupp As New Core_Module.clsMasterSupplier
                SuppName = ObjSupp.GetSupplierName(Me.mPrimarySupplier)
                ObjSupp = Nothing
            Else
                SuppName = ""
            End If

            If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                DescSC = "Purchase from " + SuppName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                DescSC = "Purchase return to " + SuppName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                DescSC = "Sales to " + CustName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Sales
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                DescSC = "Sales return from " + CustName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_SalesReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                DescSC = "Stock in adjust"
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                DescSC = "Stock out adjust"
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            End If

            If vCekDataBenar = True Then
                If Me.TipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                    If Me.ApakahDataBenar_StockInPurchase(True, vdgvItem) = True Then
                        Lanjut = True
                    End If
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    If Me.ApakahDataBenar_StockOutPurchaseReturn(True, vdgvItem) = True Then
                        Lanjut = True
                    End If
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                    If Me.ApakahDataBenar_StockOutSales(True, vdgvItem) = True Then
                        Lanjut = True
                    End If
                ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                    If Me.ApakahDataBenar_StockInSalesReturn(True, vdgvItem) = True Then
                        Lanjut = True
                    End If
                Else
                    If Me.ApakahDataBenar_Default(True, vdgvItem, Me.TipeTransaksi, vPrimaryPickUpOrder) = True Then
                        Lanjut = True
                    End If
                End If

                '################################CUSTOM INOUT#############################################
                If ObjCustInOut.ApakahDataBenar_CustomStockInOut(Me.mPrimary, True, JenisStockCardCustom, _
                tblProdQty, vTblItemCustom, ListProductCustom, JenisTransAsalCustom) = False Then
                    Return False
                End If
                '################################CUSTOM INOUT#############################################
            Else
                Lanjut = True
            End If

            If Lanjut = True Then
                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA.Insert(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTipeTransaksi, _
                   Me.mTanggalTransaksi, Me.mPrimaryPOPurchase, Me.mPrimaryDOSales, Me.mPrimaryCust, _
                   Me.mPrimarySupplier, Me.mImporDariPackingList, Me.mPrimaryGIT, Me.mRemarks, CreatedBy_, _
                   CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mSupplierDoNo, Me.mPrimaryAdjType) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    If vCekDataBenar = True Then
                        'berkaitan penomoran yang sama untuk bulan yang berbeda
                        Me.CekNomor(DTA, Me.mTipeTransaksi)
                    End If

                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                        Dim ijk As New FrmTransaksiStockInAdj
                        rowDGV = vdgvItem.Rows(0)
                        If CDec(rowDGV.Cells(ijk.ColAmount.Index).Value) <> 0 Then
                            DTA.UpdateQuery_SetBeginingBalance(1, Me.mPrimary)
                            BeginingBalance = True
                        End If
                        ijk = Nothing
                    End If

                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                        'notification email
                        Dim ObjDO As New ClsTransaksiDO
                        Dim ObjItemDO As New ClsTransaksiDOItem
                        'If vAskEmailNotification = True Then
                        '    If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        '        Me.SendStockOutSalesReadyNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
                        '        Me.mModifiedBy, CustCode, CustName)
                        '    End If
                        'End If
                        'update data primary stock out
                        ObjDO.SetPrimaryStockOut(Me.mPrimaryDOSales, Me.mPrimary)
                        'update data item delivery order sesuai dengan yang dituliskan pada bon out 
                        Dim DTA_DO_Item As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltrans_do_itemTableAdapter
                        Dim TBL_DO_Item As New dseTransaksiStockInOut_CoverMinus.tbltrans_do_itemDataTable
                        DTA_DO_Item.Fill(TBL_DO_Item, Me.mPrimaryDOSales)

                        ObjItemDO.Hapus(Me.mPrimaryDOSales)
                        ObjItemDO.Simpan_DariFormStockInOut(vdgvItem, Me.mPrimaryDOSales, TBL_DO_Item, vSetPOB)
                        ObjDO = Nothing
                        ObjItemDO = Nothing
                        DTA_DO_Item.Dispose()
                        TBL_DO_Item.Dispose()
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                        Dim ObjDO As New ClsTransaksiDO
                        Dim ObjCust As New Core_Module.clsMasterCustomer
                        'notification email
                        'If vAskEmailNotification = True Then
                        '    If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        '        Me.SendStockInSalesReturnNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
                        '        Me.mModifiedBy, ObjDO.GetUserModifiedBy(Me.mPrimaryDOSales), ObjCust.GetCustName(Me.mPrimaryCust))
                        '    End If
                        'End If
                        ObjDO = Nothing
                        ObjCust = Nothing
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        'simpan GRN
                        Dim NoGRN As String = ""
                        Me.SimpanGRN(vdgvItem, vUserCode, NoGRN)
                        'kirim email GRN
                        'If vAskEmailNotification = True Then
                        '    If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        '        Me.SendStockInPurchaseNotificationEmail(Me.mNoTransaksi, NoGRN, vUserCode)
                        '    End If
                        'End If
                    End If

                    'simpan data item

                    ''@rosadinnn
                    ''@MUaT
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        ObjItem.Simpan_StockInPurchase(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, vSetPOB, BeginingBalance, vKoreksiProdQty)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        ObjItem.Simpan_StockOutPurchaseReturn(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, vSetPOB, BeginingBalance, vKoreksiProdQty)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                        ObjItem.Simpan_StockOutSales(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, vSetPOB, BeginingBalance, vKoreksiProdQty)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                        ObjItem.Simpan_StockInSalesReturn(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, vSetPOB, BeginingBalance, vKoreksiProdQty)
                    Else
                        ObjItem.Simpan(vdgvItem, Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, vSetPOB, BeginingBalance, vKoreksiProdQty)
                    End If

                    '################################CUSTOM INOUT#############################################
                    Dim GetUlangProdQty As Boolean = False
                    If vCekDataBenar = False Then
                        GetUlangProdQty = True
                    End If
                    ObjCustInOut.Simpan_CustomStockInOut(Me.mPrimary, Me.mNoTransaksi, Me.mTanggalTransaksi, DescSC, _
                    JenisTransAsalCustom, JenisStockCardCustom, vTblItemCustom, tblProdQty, ListProductCustom, GetUlangProdQty)
                    '################################CUSTOM INOUT#############################################

                    'berkaitan dengan proses booking inventory
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        If Me.mPrimaryGIT = -1 Then
                            'jika langsung dari stock in purchase maka simpan data booking inventory jika ada
                            ObjBooking.CreateBookingInventory_TanpaGIT(Me.mPrimary, vUserCode)
                        Else
                            'jika dari packing list, cek apakah ada booking git, jika ada maka 
                            'ubah data booking git menjadi booking inventory
                            ObjBooking.CreateBookingInventory_DenganGIT(Me.mPrimary, Me.mPrimaryGIT, vUserCode)
                        End If
                    ElseIf Me.TipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                        'berkaitan dengan pick up order
                        If vPrimaryPickUpOrder <> -1 Then
                            Dim DTA_PickUp As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                            DTA_PickUp.Connection.ConnectionString = Core_Module.GetConstring
                            DTA_PickUp.UpdateQuery_SetPrimaryStockOutAdjust(Me.mPrimary, vPrimaryPickUpOrder)
                            DTA_PickUp.Dispose()
                        End If
                    End If

                    'simpan data jurnal
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        Dim ObjPR As New ClsTransaksiPurchaseReturn
                        ObjPR.SetPrimaryStockOutPurchaseReturn(Me.mPrimaryPOPurchase, Me.mPrimary)
                        Me.TransferJurnal(Me.mPrimarySupplier, Me.mPrimary, PrimaryPRLama)
                        ObjPR = Nothing
                    End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vdgvItem, vUserCode, vTblItemCustom, vAskEmailNotification, vSetPOB, vPrimaryPickUpOrder, vCekDataBenar, vKoreksiProdQty, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjBooking = Nothing
        End Try
    End Function

    Public Sub TransferJurnal(ByVal vPrimarySupplier As Double, ByVal vPrimaryStockOutPurchase As Double, ByVal vPrimaryPRLama As Double)
        Dim DTA As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnDataTable
        Dim row As DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnRow
        Dim Tipe_PR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur
        Dim PR As Decimal
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTA.Fill(TBL, vPrimaryStockOutPurchase)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Tipe_PR = row.Tipe_PR
                If Tipe_PR = EnumTipe_PurchaseReturn.Nota_Retur Then
                    If row.Primary_Currency = 1 Then
                        PR = row.Sub_Total
                    Else
                        PR = Math.Round(row.Sub_Total * row.Tax_Exchange_Rate, ObjCur.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero)
                    End If
                    With row
                        Me.TransferJurnal_PR(.Primary, vPrimarySupplier, _
                        PR, (row.Tax_Amount * row.Tax_Exchange_Rate), (row.Dutystamp * row.Tax_Exchange_Rate), _
                        row.Urutan, .Tanggal_Transaksi, .No_Transaksi, .Created_By, _
                        .Payment_Term, row.Sub_Total, row.Tax_Exchange_Rate, _
                        row.Primary_Currency, row.primary_ap_account, row.Primary_Curr_Tax, row.Tax_Amount, row.Dutystamp, vPrimaryPRLama)
                    End With
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub CekNomor(ByVal DTA As dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter, _
    ByVal vJenisTransaksi As EnumTipeStockInOut)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Double = Me.mNoUrut
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1
        Dim InOUT As String = ""

        Try
            If vJenisTransaksi = EnumTipeStockInOut.StockIn_Adjust Or vJenisTransaksi = EnumTipeStockInOut.StockIn_BackProcess Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_Process Or vJenisTransaksi = EnumTipeStockInOut.StockIn_Purchase Or _
            vJenisTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                InOUT = "IN"
            Else
                InOUT = "OUT"
            End If

            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If InOUT = "IN" Then
                    If DTA.ScalarQuery_GetJmlInBuUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                        JmlTrans = DTA.ScalarQuery_GetJmlInBuUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).Value
                    Else
                        JmlTrans = 1
                    End If
                ElseIf InOUT = "OUT" Then
                    If DTA.ScalarQuery_GetJmlOutByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                        JmlTrans = DTA.ScalarQuery_GetJmlOutByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).Value
                    Else
                        JmlTrans = 1
                    End If
                End If

                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If InOUT = "IN" Then
                        If DTA.ScalarQuery_GetPrimaryMaxInByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                            PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxInByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari)
                        Else
                            PrimaryMax = -1
                        End If
                    ElseIf InOUT = "OUT" Then
                        If DTA.ScalarQuery_GetPrimaryMaxOutByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                            PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxOutByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari)
                        Else
                            PrimaryMax = -1
                        End If
                    End If

                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(vJenisTransaksi, Me.mTanggalTransaksi)
                        Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mNoTransaksi, Me.mNoUrut, Me.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vAskEmailNotification As Boolean = True, Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            Me.SudahHapusAntrian = False

                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vTblItemCustom, vAskEmailNotification, True, vPrimaryPickUpOrder)
                            'hapus data antri untuk transaksi ini
                            If Me.SudahHapusAntrian = False Then
                                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                    MsgBox("Logic error(1), contact your provider when you see this message", _
                                    MsgBoxStyle.Information, "Information")
                                End If
                                Me.SudahHapusAntrian = True
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Private Sub SimpanGRN(ByVal vDGV As DataGridView, ByVal vUserCode As String, ByRef vNoGRN As String)
        Dim ObjGRN As New ClsTransaksiGRN
        Try
            With ObjGRN
                .NoTransaksi = ObjGRN.GetNoBaru(Me.TanggalTransaksi)
                vNoGRN = .NoTransaksi
                .TanggalTransaksi = Me.TanggalTransaksi
                .primaryGIT = Me.PrimaryGIT
                .primaryStockInPurchase = Me.Primary
                .primarySupplier = Me.PrimarySupplier
                .Remark = Me.Remarks
                .SupplierDONumber = Me.SupplierDoNo
            End With
            ObjGRN.SimpanBaruLewatTransaksiStockInPurchase(vDGV, vUserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjGRN = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengirimkan email kepada user yang menginput delivery order
    ''' </summary>
    ''' <param name="DONo"></param>
    ''' <param name="User"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendStockInSalesReturnNotificationEmail(ByVal StockInNo As String, ByVal DONo As String, _
        ByVal User As String, ByVal UserTujuan As String, ByVal CustomerName As String) As Boolean
        Dim ObjEmail As New E_mail_Module.ClsTransEmail
        Dim ListUser As New List(Of String)

        Try
            With ObjEmail
                .Content = "There goods returned from our dear customer: " + CustomerName + ". The stock in sales return number is: " + _
                StockInNo + "." + vbCrLf + "Please confirm to our customer for futher information." + vbCrLf + _
                vbCrLf + vbCrLf + "Regards, " + vbCrLf _
                + vbCrLf + CStr(User)
                .CreatedBy = User
                .TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                .Subject = "Stock In:  " + CStr(StockInNo) + " - Returned by customer."
            End With

            ListUser.Add(UserTujuan)

            If ObjEmail.SendEmail(ListUser) = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SendStockInPurchaseNotificationEmail(ByVal vStockInNo As String, _
    ByVal vGRNNo As String, ByVal vUser As String) As Boolean
        Dim ObjEmail As New E_mail_Module.ClsTransEmail
        Dim ListUser As New List(Of String)
        Dim DTA As New E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        Dim TBL As New E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoDataTable
        Dim row As E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoRow

        Try
            With ObjEmail
                .Content = "The goods for stock in purchase:  " + CStr(vStockInNo) + " is ready. The GRN no is: " + vGRNNo + "." + vbCrLf + _
                vbCrLf + "Regards, " + vbCrLf _
                + vbCrLf + CStr(vUser)
                .CreatedBy = vUser
                .TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                .Subject = "Stock In Purchase:  " + CStr(vStockInNo) + " - Ready."
            End With

            DTA.FillBy_PrimaryEmail(TBL, 3)
            For Each row In TBL.Rows
                ListUser.Add(row.User_Code)
            Next

            If ObjEmail.SendEmail(ListUser) = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function SendStockOutSalesReadyNotificationEmail(ByVal StockOutNo As String, ByVal DONo As String, _
    ByVal User As String, ByVal vCustomerCode As String, ByVal vCustomerName As String) As Boolean
        Dim ObjEmail As New E_mail_Module.ClsTransEmail
        Dim ListUser As New List(Of String)
        Dim DTA As New E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        Dim TBL As New E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoDataTable
        Dim row As E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoRow
        Dim StatusKwitansi, StatusInv As String
        Dim DTA_DO As New DseTransaksiDOTableAdapters.tbltrans_doTableAdapter
        Dim TBL_DO As New DseTransaksiDO.tbltrans_doDataTable
        Dim row_DO As DseTransaksiDO.tbltrans_doRow

        Try
            With ObjEmail
                If DTA_DO.Fill(TBL_DO, DONo) > 0 Then
                    row_DO = TBL_DO.Rows(0)
                    If row_DO.Cetak_Kwitansi = True Then
                        StatusKwitansi = "Kwitansi = YES"
                    Else
                        StatusKwitansi = "Kwitansi = NO"
                    End If
                    If row_DO.Cetak_Invoice_N_Faktur_Pajak = True Then
                        StatusInv = "Invoice = YES"
                    Else
                        StatusInv = "Invoice = NO"
                    End If
                Else
                    StatusKwitansi = "Kwitansi = NO"
                    StatusInv = "Invoice = NO"
                End If

                .Content = "The goods for " + vCustomerCode + "-" + vCustomerName + " is ready. The stock out number is:  " _
                + CStr(StockOutNo) + ". The delivery order no is: " + DONo + "." + vbCrLf + _
                "Please prepare all documents for the shipment." + vbCrLf + StatusKwitansi + vbCrLf + StatusInv + vbCrLf + _
                vbCrLf + vbCrLf + "Regards, " + vbCrLf _
                + vbCrLf + CStr(User)
                .CreatedBy = User
                .TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                .Subject = "Stock Out:  " + CStr(StockOutNo) + " - " + vCustomerCode + " - Ready to send."
            End With

            DTA.FillBy_PrimaryEmail(TBL, 2)
            For Each row In TBL.Rows
                ListUser.Add(row.User_Code)
            Next

            If ObjEmail.SendEmail(ListUser) = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_DO.Dispose()
            TBL_DO.Dispose()
        End Try
    End Function

    Private Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim ObjItem As New clsTransaksiStockInOutItem
        Dim CustName, SuppName, CustCode As String
        CustName = ""
        CustCode = ""
        Dim ObjBooking As New ClsTransaksiBookingInv
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim ListPrimaryPO As New List(Of Double)
        Dim ListPrimaryPO2 As New List(Of Double)
        Dim PrimaryPO As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim BeginingBalance As Boolean = False

        Dim lanjut As Boolean = False
        Dim PrimaryPRLama As Double = -1

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut
        Dim ListProductCustom As New List(Of Double)
        Dim JenisTransAsalCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal
        Dim JenisStockCardCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard
        Dim DescSC As String = ""

        Try
            If Me.TipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Dim DTA_PR As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_purchasereturnTableAdapter
                Dim TBL_PR As New DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnDataTable
                Dim row_PR As DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnRow
                DTA_PR.Fill(TBL_PR, Me.mPrimary)
                If TBL_PR.Rows.Count = 0 Then
                    MsgBox("Logic error: purchase return not found. Please contact your vendir if you see this message.", MsgBoxStyle.Information, "Information")
                Else
                    row_PR = TBL_PR.Rows(0)
                    PrimaryPRLama = row_PR.Primary
                End If
            End If

            ''@rosadinnn
            ''@MUaT

            If Me.TipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                If Me.ApakahDataBenar_StockInPurchase(False, dgvItem) = True Then
                    lanjut = True
                End If
            ElseIf Me.TipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                If Me.ApakahDataBenar_StockOutPurchaseReturn(False, dgvItem) = True Then
                    lanjut = True
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                If Me.ApakahDataBenar_StockOutSales(False, dgvItem) = True Then
                    lanjut = True
                End If
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                If Me.ApakahDataBenar_StockInSalesReturn(False, dgvItem) = True Then
                    lanjut = True
                End If
            Else
                If Me.ApakahDataBenar_Default(False, dgvItem, Me.TipeTransaksi, vPrimaryPickUpOrder) = True Then
                    lanjut = True
                End If
            End If

            'simpan data item baru
            If Me.mPrimaryCust <> -1 Then
                Dim ObjCust As New Core_Module.clsMasterCustomer
                CustName = ObjCust.GetCustName(Me.mPrimaryCust)
                CustCode = ObjCust.GetCustomerCode(Me.mPrimaryCust)
                ObjCust = Nothing
            Else
                CustName = ""
                CustCode = ""
            End If
            If Me.mPrimarySupplier <> -1 Then
                Dim ObjSupp As New Core_Module.clsMasterSupplier
                SuppName = ObjSupp.GetSupplierName(Me.mPrimarySupplier)
                ObjSupp = Nothing
            Else
                SuppName = ""
            End If

            '################################CUSTOM INOUT#############################################
            If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                DescSC = "Purchase from " + SuppName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                DescSC = "Purchase return to " + SuppName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                DescSC = "Sales to " + CustName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Sales
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                DescSC = "Sales return from " + CustName
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_SalesReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                DescSC = "Stock in adjust"
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                DescSC = "Stock out adjust"
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            End If

            If ObjCustInOut.ApakahDataBenar_CustomStockInOut(Me.mPrimary, False, JenisStockCardCustom, _
            tblProdQty, vTblItemCustom, ListProductCustom, JenisTransAsalCustom) = False Then
                Return False
            End If
            '################################CUSTOM INOUT#############################################

            If lanjut = True Then
                If DTA.Update(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTipeTransaksi, Me.mTanggalTransaksi, _
                Me.mPrimaryPOPurchase, Me.mPrimaryDOSales, Me.mPrimaryCust, Me.mPrimarySupplier, _
                Me.mImporDariPackingList, Me.mPrimaryGIT, Me.mRemarks, vUserCode, Now, Now, Me.SupplierDoNo, Me.mPrimaryAdjType, Me.mPrimary, _
                Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                        Dim ijk As New FrmTransaksiStockInAdj
                        rowDGV = dgvItem.Rows(0)
                        If CDec(rowDGV.Cells(ijk.ColAmount.Index).Value) <> 0 Then
                            DTA.UpdateQuery_SetBeginingBalance(1, Me.mPrimary)
                        End If
                        ijk = Nothing
                    End If

                    'email pemberitahuan
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                        Dim ObjDO As New ClsTransaksiDO
                        Dim ObjitemDO As New ClsTransaksiDOItem
                        'If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        '    Me.SendStockOutSalesReadyNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
                        '    Me.mModifiedBy, CustCode, CustName)
                        'End If
                        'update data primary stock out sales baru
                        ObjDO.SetPrimaryStockOut(Me.mPrimaryDOSales, Me.mPrimary)
                        'update data item delivery order sesuai dengan yang dituliskan pada bon out 
                        Dim DTA_DO_Item As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltrans_do_itemTableAdapter
                        Dim TBL_DO_Item As New dseTransaksiStockInOut_CoverMinus.tbltrans_do_itemDataTable
                        DTA_DO_Item.Fill(TBL_DO_Item, Me.mPrimaryDOSales)
                        'update data item delivery order sesuai dengan yang dituliskan pada bon out
                        ObjitemDO.Hapus(Me.mPrimaryDOSales)
                        ObjitemDO.Simpan_DariFormStockInOut(dgvItem, Me.mPrimaryDOSales, TBL_DO_Item)
                        ObjDO = Nothing
                        ObjitemDO = Nothing
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                        Dim ObjDO As New ClsTransaksiDO
                        Dim ObjLS As New Core_Module.clsMasterListedSoftware
                        Dim ObjCust As New Core_Module.clsMasterCustomer
                        If ObjLS.IsSalesInstalled = True Then
                            'notification email
                            'If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                            '    Me.SendStockInSalesReturnNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
                            '    Me.mModifiedBy, ObjDO.GetUserModifiedBy(Me.mPrimaryDOSales), ObjCust.GetCustName(Me.mPrimaryCust))
                            'End If
                        End If
                        ObjLS = Nothing
                        ObjDO = Nothing
                        ObjCust = Nothing
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        Dim objGRN As New ClsTransaksiGRN
                        Dim ObjPR As New ClsTransaksiPurchaseReturn
                        Dim PrimaryGRN As Double = -1
                        Dim PrimaryGRNBaru As Double = -1

                        Dim NoGRN As String = ""

                        PrimaryGRN = objGRN.GetPrimary_ByPrimaryStockInPurchase(Me.mPrimary)
                        'hapus dulu GRN
                        objGRN.HapusDariStockInPurchase(Me.mPrimary)
                        'simpan GRN
                        Me.SimpanGRN(dgvItem, vUserCode, NoGRN)

                        'kirim email
                        'If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                        '    Me.SendStockInPurchaseNotificationEmail(Me.mNoTransaksi, NoGRN, vUserCode)
                        'End If

                        'update primary grn ke dalam semua transaksi purchase return yang merujuk
                        PrimaryGRNBaru = objGRN.GetPrimary_ByPrimaryStockInPurchase(Me.mPrimary)
                        ObjPR.SetPrimaryGRN(PrimaryGRN, PrimaryGRNBaru)

                        objGRN = Nothing
                        ObjPR = Nothing
                    End If

                    'ambil data primary po, untuk digunakan untuk perhitungan order balance
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        Me.GetPrimaryPO(Me.mPrimary, ListPrimaryPO)
                    End If
                    BeginingBalance = Me.ApakahBeginingBalance(Me.mPrimary)

                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me.mTipeTransaksi, True, BeginingBalance)

                    ''################################CUSTOM INOUT#############################################
                    ObjCustInOut.Hapus_CustomStockInOut(Me.mPrimary, JenisTransAsalCustom, JenisStockCardCustom, _
                    tblProdQty, ListProductCustom, False)
                    ''################################CUSTOM INOUT#############################################

                    'untuk simpan item disini, tidak perlu set pob, karena perhitungannya akan dilakukan disini
                    ''@rosadinnn
                    ''@MUaT
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        ObjItem.Simpan_StockInPurchase(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, False, BeginingBalance)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        ObjItem.Simpan_StockOutPurchaseReturn(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, False, BeginingBalance)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                        ObjItem.Simpan_StockOutSales(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, False, BeginingBalance)
                        ObjCustInOut.SimpanPerubahanCustom_StockOutSales(Me.mPrimary, Me.mNoTransaksi)
                    ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                        ObjItem.Simpan_StockInSalesReturn(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, False, BeginingBalance)
                    Else
                        ObjItem.Simpan(dgvItem, Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mPrimary, _
                            Me.mNoTransaksi, SuppName, CustName, False, BeginingBalance)
                    End If

                    '################################CUSTOM INOUT#############################################
                    ObjCustInOut.Simpan_CustomStockInOut(Me.mPrimary, Me.mNoTransaksi, Me.mTanggalTransaksi, DescSC, _
                    JenisTransAsalCustom, JenisStockCardCustom, vTblItemCustom, tblProdQty, ListProductCustom)
                    '################################CUSTOM INOUT#############################################

                    'hitung order balance
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        Me.GetPrimaryPO(Me.mPrimary, ListPrimaryPO2)
                        For Each PrimaryPO In ListPrimaryPO2
                            If ListPrimaryPO.Contains(PrimaryPO) = False Then
                                ListPrimaryPO.Add(PrimaryPO)
                            End If
                        Next
                        ObjPOItem.SetOrderBalance(ListPrimaryPO)
                    End If


                    'berkaitan dengan proses booking inventory
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        If Me.mPrimaryGIT = -1 Then
                            'jika langsung dari stock in purchase maka simpan data booking inventory jika ada
                            ObjBooking.CreateBookingInventory_TanpaGIT(Me.mPrimary, vUserCode)
                        Else
                            'jika dari packing list, cek apakah ada booking git, jika ada maka 
                            'ubah data booking git menjadi booking inventory
                            ObjBooking.CreateBookingInventory_DenganGIT(Me.mPrimary, Me.mPrimaryGIT, vUserCode)
                        End If
                        ObjCustInOut.SimpanPerubahanCustom_StockInPurchase(Me.mPrimary, Me.mNoTransaksi)
                    ElseIf Me.TipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                        'berkaitan dengan pick up order
                        If vPrimaryPickUpOrder <> -1 Then
                            Dim DTA_PickUp As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                            DTA_PickUp.Connection.ConnectionString = Core_Module.GetConstring
                            DTA_PickUp.UpdateQuery_SetPrimaryStockOutAdjust(Me.mPrimary, vPrimaryPickUpOrder)
                            DTA_PickUp.Dispose()
                        End If
                        ObjCustInOut.SimpanPerubahanCustom_StockOutAdjust(Me.mPrimary, Me.mNoTransaksi)
                    End If

                    'simpan data jurnal
                    If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        Dim ObjPR As New ClsTransaksiPurchaseReturn
                        ObjPR.SetPrimaryStockOutPurchaseReturn_Minus_Satu(Me.mPrimary)
                        ObjPR.SetPrimaryStockOutPurchaseReturn(Me.mPrimaryPOPurchase, Me.mPrimary)
                        Me.TransferJurnal(Me.mPrimarySupplier, Me.mPrimary, PrimaryPRLama)
                        ObjPR = Nothing
                    End If

                    'nilai balikan
                    Return True
                Else
                    'cek apakah timestampnya beda, jika ya maka minta user untuk reload data
                    If DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value <> Me.mTimestamp Then
                        If Me.SudahHapusAntrian = False Then
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Me.SudahHapusAntrian = True
                        End If

                        MsgBox("Other user have modified this transaction, please reload from database.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan1(dgvItem, vUserCode, vTblItemCustom)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjBooking = Nothing
            ObjPOItem = Nothing
            ListPrimaryPO = Nothing
            ListPrimaryPO2 = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            Me.SudahHapusAntrian = False

                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vdgvItem, vUserCode, vTblItemCustom, vPrimaryPickUpOrder)
                            'hapus data antri untuk transaksi ini
                            If Me.SudahHapusAntrian = False Then
                                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                    MsgBox("Logic error(1), contact your provider when you see this message", _
                                    MsgBoxStyle.Information, "Information")
                                End If
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Threading.Thread.Sleep(1000)
        Finally
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, _
    ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_Hapus, vUser)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            Me.SudahHapusAntrian = False

                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUser, vTblItemCustom, True, vPrimaryPickUpOrder)
                            'hapus data antri untuk transaksi ini
                            If Me.SudahHapusAntrian = False Then
                                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                    MsgBox("Logic error(1), contact your provider when you see this message", _
                                    MsgBoxStyle.Information, "Information")
                                End If
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub GetPrimaryPO(ByVal vPrimaryStockInPurchase As Double, ByVal vListPrimaryPO As List(Of Double))
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_item4TableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinout_item4DataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinout_item4Row

        Try
            vListPrimaryPO.Clear()
            DTA.Fill(TBL, vPrimaryStockInPurchase)
            For Each row In TBL.Rows
                vListPrimaryPO.Add(row.Primary_PO_Purchase)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Function GetPrimaryGIT(ByVal vPrimaryStockInPurchase As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryGIT(vPrimaryStockInPurchase).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryGIT(vPrimaryStockInPurchase).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeTransaksi(ByVal vPrimary As Double) As EnumTipeStockInOut
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTipeTransaksi(vPrimary)

            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return EnumTipeStockInOut.StockIn_Adjust
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' jika cek minus "False", maka fungsi tidak akan melakukan pengecekan minus qty
    ''' </summary>
    ''' <param name="vPrimaryTransaksi"></param>
    ''' <param name="vUser"></param>
    ''' <param name="vCekMinus"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, _
    ByVal vTblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vCekMinus As Boolean = True, Optional ByVal vPrimaryPickUpOrder As Double = -1) As Boolean
        Dim Obj As New clsTransaksiStockInOutItem
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTransaksi As EnumTipeStockInOut = EnumTipeStockInOut.StockIn_Adjust
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)
        Dim PrimaryGIT As Double
        Dim Lanjut As Boolean = True
        Dim BeginingBalance As Boolean = False

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut
        Dim ListProductCustom As New List(Of Double)
        Dim JenisTransAsalCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal
        Dim JenisStockCardCustom As CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard

        Try
            If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Sales
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_SalesReturn
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In
            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                JenisTransAsalCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockOut_Adjust
                JenisStockCardCustom = CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_Out
            End If

            If vCekMinus = False Then
                Lanjut = True
            Else
                Dim ObjItem As New clsTransaksiStockInOutItem
                If ObjItem.ApakahHapusMengakibatkanMinus(vPrimaryTransaksi, Me.GetTipeTransaksi(vPrimaryTransaksi), True, Me.vPrimary_Antrian, _
                Me.DTAInOUT, Me.SudahHapusAntrian) = True Then
                    ObjItem = Nothing
                    Return False
                End If
                ObjItem = Nothing

                '################################CUSTOM INOUT#############################################
                If ObjCustInOut.ApakahBolehHapus_CustomStockInOut(vPrimaryTransaksi, True, JenisStockCardCustom, _
                tblProdQty, ListProductCustom, JenisTransAsalCustom) = False Then
                    Return False
                End If
                '################################CUSTOM INOUT#############################################
            End If


            If Lanjut = True Then
                TipeTransaksi = DTA.ScalarQuery_GetTipeTransaksi(vPrimaryTransaksi).Value
                'update data delivery order
                If TipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                    Dim ObjDO As New ClsTransaksiDO
                    ObjDO.SetPrimaryStockOut(DTA.ScalarQuery_GetPrimaryDOSales(vPrimaryTransaksi).Value, -1)
                    ObjDO = Nothing
                End If

                If TipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Or TipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    'ambil list primary po
                    Me.GetPrimaryPO(vPrimaryTransaksi, vListPrimaryPO)
                End If

                'hapus data item 
                BeginingBalance = Me.ApakahBeginingBalance(vPrimaryTransaksi)
                Obj.Hapus(vPrimaryTransaksi, TipeTransaksi, True, BeginingBalance)

                '################################CUSTOM INOUT#############################################
                Dim GetUlangProdQty As Boolean = False
                If vCekMinus = False Then
                    GetUlangProdQty = True
                End If
                ObjCustInOut.Hapus_CustomStockInOut(vPrimaryTransaksi, JenisTransAsalCustom, JenisStockCardCustom, tblProdQty, ListProductCustom, True, GetUlangProdQty)
                '################################CUSTOM INOUT#############################################

                'hapus data jurnal
                If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    Me.Hapus_Jurnal_PR(vPrimaryTransaksi)
                End If

                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    'untuk stock in purchase, hapus juga data GRN dan hitung ulang nilai order balance
                    If TipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
                        'hapus data GRN
                        Dim objGRN As New ClsTransaksiGRN
                        objGRN.HapusDariStockInPurchase(vPrimaryTransaksi)
                        objGRN = Nothing
                        'set aktif kembali semua booking yang merujuk pada packing list ybs
                        PrimaryGIT = Me.GetPrimaryGIT(vPrimaryTransaksi)
                        If PrimaryGIT <> -1 Then
                            Dim ObjBooking As New ClsTransaksiBookingInv
                            ObjBooking.SetBookingGITAktif(PrimaryGIT)
                            ObjBooking = Nothing
                        End If
                        'set ulang nilai order balance
                        ObjPOItem.SetOrderBalance(vListPrimaryPO)
                    ElseIf TipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        'untuk purchase return, set nilai primary stock out purchase return 
                        'menjadi -1
                        'tidak perlu menset nilai POB, karena sudah di set ketika simpan transaksi purchase return
                        Dim ObjPR As New ClsTransaksiPurchaseReturn
                        ObjPR.SetPrimaryStockOutPurchaseReturn_Minus_Satu(vPrimaryTransaksi)
                        ObjPR = Nothing
                    ElseIf TipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Then
                        If vPrimaryPickUpOrder <> -1 Then
                            Dim DTA_PickUp As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
                            DTA_PickUp.Connection.ConnectionString = Core_Module.GetConstring
                            DTA_PickUp.UpdateQuery_SetPrimaryStockOutAdjust(-1, vPrimaryPickUpOrder)
                            DTA_PickUp.Dispose()
                        End If
                    End If

                    'nilai balikan
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
            ObjPOItem = Nothing
            'vListPrimaryPO = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, ByVal vConString As String, _
     Optional ByVal vCekMinus As Boolean = True) As Boolean
        Dim Obj As New clsTransaksiStockInOutItem
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTransaksi As EnumTipeStockInOut = EnumTipeStockInOut.StockIn_Adjust
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)
        Dim PrimaryGIT As Double
        Dim Lanjut As Boolean = True
        Dim BeginingBalance As Boolean = False

        Try
            ''@rosadinnn
            If String.IsNullOrEmpty(vConString) = False Then
                DTA.Connection.ConnectionString = vConString
            End If

            If Lanjut = True Then

                'hapus data item 
                BeginingBalance = Me.ApakahBeginingBalance(vPrimaryTransaksi, vConString)
                Obj.Hapus(vPrimaryTransaksi, TipeTransaksi, vConString, True, BeginingBalance)

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
            ObjPOItem = Nothing
        End Try
    End Function

    Public Sub Hapus_Jurnal_PR(ByVal vPrimaryStockOut As Double)
        Dim DTA As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnDataTable
        Dim row As DseTransaksiPurchaseReturn_TransfJurnal.tbltrans_purchasereturnRow
        Dim Tipe_PR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur

        Try
            DTA.Fill(TBL, vPrimaryStockOut)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Me.Hapus_Jurnal(row.Primary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Overridable Function HapusDariGRN(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Obj As New clsTransaksiStockInOutItem
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTransaksi As EnumTipeStockInOut

        Try
            TipeTransaksi = DTA.ScalarQuery_GetTipeTransaksi(vPrimaryTransaksi).Value
            'update data delivery order
            'hapus data item 
            Obj.Hapus(vPrimaryTransaksi, TipeTransaksi)
            'hapus data header, data item otomatis dihapuskan oleh server
            If DTA.Delete(vPrimaryTransaksi) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Sub GetListPrimaryCodeNo(ByVal vPrimaryStockInOut As Double, ByVal vListPrimary As List(Of Double))
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_item5TableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinout_item5DataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinout_item5Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryStockInOut)

            vListPrimary.Clear()
            For Each row In TBL.Rows
                vListPrimary.Add(row.Primary_Product)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCustomer(ByVal vPrimary As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCustomer(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCustomer(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryDO(ByVal vPrimaryTransaksiStockInSalesReturn As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            Return DTA.ScalarQuery_GetPrimaryDOSales(vPrimaryTransaksiStockInSalesReturn).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetExpiryDate(ByVal vPrimaryProduct As Double, ByVal vBatchNo _
    As String, ByVal vJenisTransaksi As Integer) As Date
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetExpiryDate(vJenisTransaksi, vPrimaryProduct, vBatchNo).HasValue = True Then
                Return DTA.ScalarQuery_GetExpiryDate(vJenisTransaksi, vPrimaryProduct, vBatchNo).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan no transaksi yang dirujuk oleh DO atau DO Loss yang diberikan sebagai parameter
    ''' </summary>
    ''' <param name="vPrimaryDOLoss"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(ByVal vPrimaryDOLoss As Double, _
    ByVal TipeTransaksi As EnumTipeStockInOut) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksiByPrimaryDO_PrimaryDOLoss(vPrimaryDOLoss, TipeTransaksi)) _
            = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoTransaksiByPrimaryDO_PrimaryDOLoss(vPrimaryDOLoss, TipeTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan pada transaksi GRN ketika menginputan pertama mengisikan item  dari Stock In Purchase
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub IsikanTBLStockInOut(ByVal vTBL As dseTransaksiStockInOut.tbltransaksistockinout_view1DataTable, ByVal vPrimaryStock As Double)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinout_view1TableAdapter
        Try
            DTA.FillBy_PrimaryStock(vTBL, vPrimaryStock)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub GetInfo(ByVal vPrimaryTransaksi As Double, ByRef vSupplierDONo As String, ByRef vPrimarySupplier As String, _
    ByRef vSupplierCodeNo As String, ByRef vSupplierName As String, ByRef vPrimaryPAckingList As String, _
    ByRef vPackingListNo As String)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinout_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksistockinout_viewDataTable
        Dim row As dseTransaksiStockInOut.tbltransaksistockinout_viewRow
        Try
            DTA.FillBy_PrimaryStockInOut(TBL, vPrimaryTransaksi)
            vSupplierCodeNo = ""
            vSupplierDONo = ""
            vSupplierName = ""
            vPackingListNo = ""
            vPrimarySupplier = -1
            vPrimaryPAckingList = -1
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vSupplierCodeNo = row.Supplier_Code
                vSupplierDONo = row._Supplier_s_DO_Number
                vSupplierName = row.Supplier_Name
                vPackingListNo = row.No_Packing_List
                vPrimarySupplier = row.Primary_Supplier
                vPrimaryPAckingList = row.Primary_GIT
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetTanggalTransaksi(ByVal vPrimaryTransaksi As Double, ByVal vConString As String) As Date
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            ''@rosadinnn
            If String.IsNullOrEmpty(vConString) = False Then
                DTA.Connection.ConnectionString = vConString
            End If

            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryTransaksi).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimaryTransaksi As Double) As Date
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryTransaksi).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi_ByNoTransaksi(ByVal vNoTransaksi As String) As Date
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            If DTA.ScalarQuery_GetTanggalTransaksi2(vNoTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi2(vNoTransaksi).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSupplierDONumber(ByVal vPrimaryStockIn As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierDONumber(vPrimaryStockIn)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetSupplierDONumber(vPrimaryStockIn)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan no transaksi stock in sales return number yang merujuk ke DO sales yang diberikan
    ''' primary-nya
    ''' </summary>
    ''' <param name="vPrimaryDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStockInSalesReturnNo(ByVal vPrimaryDO As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinoutDataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinoutRow
        Dim retval As String = ""
        Dim TipeTrans As EnumTipeStockInOut = EnumTipeStockInOut.StockIn_SalesReturn

        Try
            If DTA.FillBy_PrimaryDOSales(TBL, vPrimaryDO, TipeTrans) > 0 Then
                For Each row In TBL.Rows
                    If retval = "" Then
                        retval = row.No_Transaksi
                    Else
                        retval = ", " + row.No_Transaksi
                    End If
                Next

                Return retval
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan no transaksi stock in yang merujuk pada packing list yang diberikan primarynya
    ''' </summary>
    ''' <param name="vPrimaryPackingList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStockInNumberByPrimaryPackingList(ByVal vPrimaryPackingList As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksiByPrimaryGIT(vPrimaryPackingList)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoTransaksiByPrimaryGIT(vPrimaryPackingList)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vJenisTransaksi As EnumTipeStockInOut, ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#), _
    Optional ByVal vInOutAdjust As clsTransaksiStockInOut.EnumInOutAdjust = EnumInOutAdjust.NoTipe) As Long
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari, vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari, vJenisTransaksi, vInOutAdjust).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                If DTA.ScalarQuery_GetJmlBySupplierLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlBySupplierLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                If DTA.ScalarQuery_GetJmlBySupplierName(vJenisTransaksi, vInOutAdjust, Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlBySupplierName(vJenisTransaksi, vInOutAdjust, Criteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                If DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                If DTA.ScalarQuery_GetJmlByCustNameLIKE(vJenisTransaksi, Criteria + "%", vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByCustNameLIKE(vJenisTransaksi, Criteria + "%", vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoWorkOrder Then
                'msgbox("belum")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONumber Then
                If DTA.ScalarQuery_GetJmlByPONumberLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByPONumberLIKE(Criteria + "%", vJenisTransaksi, vInOutAdjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByPurchaseReturnNo Then
                Dim v As New Nullable(Of Long)
                v = DTA.ScalarQuery_GetJmlByPurchaseReturnNo(vJenisTransaksi, vInOutAdjust, Criteria + "%")
                If v.HasValue Then
                    Return v.Value
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

    'Public Overridable Function IsiDGV(ByVal vJenisTransaksi As EnumTipeStockInOut, ByVal Criteria As String, _
    'ByVal DGV As DataGridView, ByVal TBL As dseTransaksiStockInOut.tbltransaksi_stockinout1DataTable, _
    'ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    'Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
    'Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout1TableAdapter
    'Dim QueryResult As Double
    'Dim StartT As Date
    '   Try
    '       StartT = Now
    '       Windows.Forms.Cursor.Current = Cursors.WaitCursor'

    '       If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
    '          QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%", vJenisTransaksi)
    '     ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
    '        QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari, vJenisTransaksi)
    '   ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
    '      QueryResult = DTA.FillBy_SupplierLIKE(TBL, Criteria + "%", vJenisTransaksi)
    '  ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
    '      QueryResult = DTA.FillBy_CustCodeLIKE(TBL, vJenisTransaksi, Criteria + "%")
    '  ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
    '      QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, vJenisTransaksi, Criteria + "%")
    '  ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
    '      QueryResult = DTA.FillBy_ProductDescLIKE(TBL, vJenisTransaksi, Criteria + "%")
    '  End If

    'update dgv
    '        DGV.DataSource = TBL.DefaultView
    '        DGV.Refresh()
    '        If DGV.ColumnCount > 0 Then
    ''             DGV.Columns(0).Visible = False
    '             DGV.Columns(1).HeaderText = "Transaction No"
    ' Dim ij As New DataGridViewCellStyle
    '             ij.Format = "dd MMM yyyy"
    '             DGV.Columns(2).DefaultCellStyle = ij
    '         End If

    '    Dim TS As TimeSpan = Now - StartT
    '            If QueryResult > 0 Then
    '                If TS.TotalSeconds > 1 Then
    '                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds / 10000) + " seconds"
    '                Else
    '                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / 10000, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
    '                End If
    '            Else
    '                tsl.Text = "No result found"
    '            End If
    '
    '         Return True
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         DTA.Dispose()
    '         Windows.Forms.Cursor.Current = Cursors.Arrow
    '     End Try
    ' End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mImporDariPackingList = False
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mPrimaryCust = -1
            Me.mPrimaryDOSales = -1
            Me.mPrimaryGIT = -1
            Me.mPrimaryPOPurchase = -1
            Me.mPrimarySupplier = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Adjust
            Me.mSupplierDoNo = ""
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Koreksi4HapusOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty _
    As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, _
    ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, Optional ByVal _
    vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "", _
    Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                    vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, ByVal vPrimaryProduct As _
    Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, _
    Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, _
    Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, _
    Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hapus_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " + _
            vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Decimal
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Qty
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.Qty = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Insert_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            row = TBL.Newtblmaster_productqtyRow

            row.Primary = Me.GetPrimaryMax_TBLQtySimulasi(TBL) + 1
            row.Primary_Product = vPrimaryProduct
            row.Primary_Storage_Area = vPrimaryLocation
            row.Qty_Secondary = vSecQty
            row.Lot_No = vBatchNo
            'expiry date diabaikan dalam simulasi karena tidak berpengaruh ke databse
            'yang sebenarnya
            row.Expiry_Date = #1/1/1950#
            row.Qty = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_TBLQtySimulasi(ByVal TBL As _
    Core_Module.dseMasterProduct.tblmaster_productqtyDataTable) As Double
        Dim Sort As String = "[Primary] DESC"
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = Sort
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Primary
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetPrimaryTransaksiQuick(ByVal vPrimaryTransaksiQuick As Double, ByVal vPrimaryStockInOut As Double)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            DTA.UpdateQuery_SetPrimaryTransaksiQuick(vPrimaryTransaksiQuick, vPrimaryStockInOut)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Function ApakahTransferOtomatis(ByVal vTanggalPR As Date) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggalPR.Date Then
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

    Private Sub TransferJurnal_PR(ByVal vPrimaryPR As Double, ByVal vPrimarySupplier As Double, ByVal vPurchaseReturn As Decimal, ByVal vVAT As Decimal, _
               ByVal vDutystamp As Decimal, ByVal vUrutanPR As Double, ByVal vTanggalPR As Date, ByVal vPRNumber As String, _
               ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vPurchaseReturnOriginalAmount As Decimal, ByVal vExchangeRate As Decimal, _
               ByVal vPrimaryCurrency As Double, ByVal vPrimaryAccount As Double, _
               ByVal vPrimaryCurrTax As Double, ByVal vVatOriginal As Decimal, ByVal vDutyStampOriginal As Decimal, ByVal vPrimaryPRLama As Double)

        Dim DTA_Balance As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim TipeBalance As EnumTipePRBalance = EnumTipePRBalance.Automatic_SR
        Dim DTA_Jurnal As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseReturnJournal
        Dim RefNumber As String = CStr(vUrutanPR)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Decimal
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim ObjPR As New ClsTransaksiPurchaseReturn
        Dim MonthStr As String = ""
        Dim PrimaryAkunAP, PrimaryAkunAP_BaseCurrency, PrimaryAkunAP_BaseIDR As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjMulti As New ClsMultiAccount

        Dim TBLInv As New DseTransaksiPurchaseReturn_TransfJurnal.DataTable1DataTable
        Dim rowInv As DseTransaksiPurchaseReturn_TransfJurnal.DataTable1Row
        Dim DTAInv As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.DataTable1TableAdapter

        Dim ListAkunInv As New List(Of Double)
        Dim PrimaryAkun As Double = -1
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RoundingIDR As Integer = 0
        Dim JmlReturnInv As Decimal = 0
        Dim IndexRowInv As Integer = 0

        Try
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTA_Balance.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTA_JurnalItem.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryPRLama)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryPRLama, TipeBalance)

            'hapuskan kartu purchase return jurnal
            ObjKartu.Hapus(vPrimaryPRLama, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalPR) = True Then
                'buat jurnal baru
                If RefNumber.Length < 6 Then
                    While RefNumber.Length < 6
                        RefNumber = "0" + RefNumber
                    End While
                End If
                MonthStr = vTanggalPR.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                'data header
                DTA_Jurnal.Insert("PRJ-" + RefNumber + "/" + MonthStr + "/" + vTanggalPR.Year.ToString, vTanggalPR, _
                vUrutanPR, TipeJurnal, vPrimaryPR, vPrimarySupplier, "PR: " + vPRNumber _
                + " - " + ObjSupplier.GetSupplierCode(vPrimarySupplier) + " - " + ObjSupplier.GetSupplierName(vPrimarySupplier), vUserCode, _
                Now, vUserCode, Now, Now)
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryPR)

                'data item jurnal
                'simpan data ap, debet
                PrimaryAkunAP = vPrimaryAccount
                PrimaryAkunAP_BaseIDR = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                PrimaryAkunAP_BaseCurrency = ObjMulti.GetPrimaryAkunAPBase(vPrimaryCurrency)

                '--------------------------------------------------------------------------------------------
                'simpan data akun AP
                Debet = 0
                Credit = 0
                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Debet = vPurchaseReturn + vVAT + vDutystamp
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", vPurchaseReturnOriginalAmount + vVatOriginal + vDutyStampOriginal, Credit, vExchangeRate)
                    '----------------------------------------------------------------------------------------------------------------------------------
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vPurchaseReturn + vVAT + vDutystamp
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", vPurchaseReturnOriginalAmount + vVatOriginal + vDutyStampOriginal, Credit, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vPurchaseReturn
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", vPurchaseReturnOriginalAmount, Credit, vExchangeRate)

                        'masukkan data vat + dutystamp ke ap base
                        If vVatOriginal + vDutyStampOriginal <> 0 Then
                            Debet = vVatOriginal + vDutyStampOriginal
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If

                'simpan data purchase-inventory-kredit
                DTAInv.Fill(TBLInv, vPrimaryPR)
                RoundingIDR = ObjCur.GetDecimal_DefaultCurrency
                IndexRowInv = 0
                JmlReturnInv = 0
                For Each rowInv In TBLInv.Rows
                    If ListAkunInv.Contains(rowInv.Primary_Inventory_Account) = False Then
                        ListAkunInv.Add(rowInv.Primary_Inventory_Account)
                        IndexRowInv += 1
                        If IndexRowInv < TBLInv.Rows.Count Then
                            Credit = Math.Round(rowInv.Inventory_Amount * rowInv.Tax_Exchange_Rate, RoundingIDR, MidpointRounding.AwayFromZero)
                            JmlReturnInv += Credit
                        Else
                            'perhitungan untuk baris terakhir agak beda
                            Credit = vPurchaseReturn - JmlReturnInv
                        End If
                        Debet = 0
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowInv.Primary_Inventory_Account, Debet, Credit, "", Debet, Credit, 1)
                    End If
                Next

                '--------------------------------------------------------------------------------------------
                'simpan data vat in-credit
                '--------------------------------------------------------------------------------------------
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATIn
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vVAT, "", 0, vVAT, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vVatOriginal, "", 0, vVatOriginal, 1)
                        End If
                    End If
                End If

                '--------------------------------------------------------------------------------------------
                'simpan data duty stamp-kredit
                '--------------------------------------------------------------------------------------------
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vDutystamp, "", 0, vDutystamp, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vDutyStampOriginal, "", 0, vDutyStampOriginal, 1)
                        End If
                    End If
                End If


                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    'aop idr
                    DTA_Balance.Insert(vPrimarySupplier, vPrimaryPR, TipeBalance, "", #1/1/1950#, "", vTanggalPR + TimeSpan.FromDays(vPaymentTerm), _
                    vPurchaseReturn * -1, vVAT * -1, vDutystamp * -1, (vPurchaseReturn + vVAT + vDutystamp) * -1, 0, _
                    vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                    (vPurchaseReturn + vVAT + vDutystamp) * -1, 1, PrimaryAkunAP, 0)

                    'buat kartu ap baru
                    ObjKartu.Simpan1(vPrimaryPR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, vPRNumber, _
                    vPrimarySupplier, vPurchaseReturn + vVAT + vDutystamp, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                    vTanggalPR, vPrimaryPR, vTanggalPR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, _
                    Core_Module.PrimaryBaseCurrency, vPurchaseReturn + vVAT + vDutystamp, 0, 1, PrimaryAkunAP_BaseCurrency)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.Insert(vPrimarySupplier, vPrimaryPR, TipeBalance, "", #1/1/1950#, "", vTanggalPR + TimeSpan.FromDays(vPaymentTerm), _
                        vPurchaseReturn * -1, vVAT * -1, vDutystamp * -1, _
                        (vPurchaseReturn + vVAT + vDutystamp) * -1, _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         (vPurchaseReturnOriginalAmount + vVatOriginal + vDutyStampOriginal) * -1, vExchangeRate, _
                        PrimaryAkunAP, 0)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan1(vPrimaryPR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, vPRNumber, _
                        vPrimarySupplier, _
                        vPurchaseReturn + vVAT + vDutystamp, 0, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalPR, vPrimaryPR, vTanggalPR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, _
                        vPrimaryCurrency, _
                        vPurchaseReturnOriginalAmount + vVatOriginal + vDutyStampOriginal, 0, _
                        vExchangeRate, PrimaryAkunAP)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.Insert(vPrimarySupplier, vPrimaryPR, TipeBalance, "", #1/1/1950#, "", vTanggalPR + TimeSpan.FromDays(vPaymentTerm), _
                        vPurchaseReturn * -1, 0, 0, _
                        vPurchaseReturn * -1, _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         vPurchaseReturnOriginalAmount * -1, vExchangeRate, _
                        PrimaryAkunAP, 0)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan1(vPrimaryPR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, vPRNumber, _
                        vPrimarySupplier, _
                        vPurchaseReturn, 0, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalPR, vPrimaryPR, vTanggalPR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, _
                        vPrimaryCurrency, _
                        vPurchaseReturnOriginalAmount, 0, _
                        vExchangeRate, PrimaryAkunAP)

                        If (vVAT + vDutystamp) <> 0 Then
                            'balance purchase tax
                            DTA_Balance.Insert(vPrimarySupplier, vPrimaryPR, TipeBalance, "", _
                            #1/1/1950#, "", vTanggalPR + TimeSpan.FromDays(vPaymentTerm), _
                            0, vVAT * -1, vDutystamp * -1, (vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                             (vVatOriginal + vDutyStampOriginal) * -1, 1, PrimaryAkunAP_BaseIDR, 0)

                            'buat kartu ar purchase tax
                            ObjKartu.Simpan1(vPrimaryPR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, vPRNumber, _
                            vPrimarySupplier, vVatOriginal + vDutyStampOriginal, 0, _
                            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalPR, vPrimaryPR, vTanggalPR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis, _
                            vPrimaryCurrTax, _
                            vVatOriginal + vDutyStampOriginal, 0, 1, PrimaryAkunAP_BaseIDR)
                        End If
                    End If
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTA_JurnalItem.Dispose()
            ObjSupplier = Nothing
            ObjKartu = Nothing
            ObjPR = Nothing
            ObjCur = Nothing
            ListAkunInv = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Private Function GetPrimaryDutyStamp() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)

            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Private Function GetPrimaryAkunVATIn() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Public Function GetPrimaryCurrencyPR(ByVal vPrimaryPR As Double) As Double
        Dim DTA As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryPR).Value
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

    Private Function GetPrimaryAkunAP(ByVal vPrimaryPR As Double) As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Return ObjCur.GetPrimaryAkunAP(Me.GetPrimaryCurrencyPR(vPrimaryPR))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Function

    Private Sub Hapus_Jurnal(ByVal vPrimaryPR As Double)
        Dim DTA_Balance As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim DTA_Jurnal As New DseTransaksiPurchaseReturn_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As EnumTipePRBalance = EnumTipePRBalance.Automatic_SR
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseReturnJournal
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryPR)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryPR, TipeBalance)
            'hapuskan data kartu ap
            Obj.Hapus(vPrimaryPR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Public Function GetAccountingConvertionDate() As Date
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltrans_accounting_closingDataTable
        Dim row As dseTransaksiStockInOut.tbltrans_accounting_closingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                Return row.Convertion_Date
            Else
                Return #1/1/1950#
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimaryTipeAdjustment(ByVal vPrimary As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryTipeAdjustment(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeAdjustment(ByVal vPrimary As Double) As EnumInOutAdjust
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeAdjustment(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTipeAdjustment(vPrimary).Value
            Else
                Return EnumInOutAdjust.NoTipe
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescriptionSettingAdjustment(ByVal vPrimarySetting As Double) As String
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim RetVal As String = ""

        Try
            RetVal = DTA.ScalarQuery_GetDesc(vPrimarySetting)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiClbSetting(ByVal vTBL As dseTransaksiStockInOut.tbl_setting_adjustmentDataTable, ByVal vClb As CheckedListBox)
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim row As dseTransaksiStockInOut.tbl_setting_adjustmentRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(vTBL)

            'set tampilan
            For Each row In vTBL.Rows
                row.Tampilan = row.Code + " - " + row.Desc
            Next

            'insert baris kedua - stock taking
            row = vTBL.Newtbl_setting_adjustmentRow
            row.Code = "Stock Taking"
            row.Desc = ""
            row.Primary = -2
            row.PrimaryGain = -1
            row.PrimaryLoss = -1
            row.Remarks = ""
            row.Tampilan = "Stock Taking Adjustment"
            vTBL.Rows.InsertAt(row, 0)

            'insert baris pertama - no type
            row = vTBL.Newtbl_setting_adjustmentRow
            row.Code = "No Tipe"
            row.Desc = ""
            row.Primary = -1
            row.PrimaryGain = -1
            row.PrimaryLoss = -1
            row.Remarks = ""
            row.Tampilan = "No Adjustment Type"
            vTBL.Rows.InsertAt(row, 0)

            vClb.DataSource = vTBL.DefaultView
            vClb.DisplayMember = "Tampilan"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSetting(ByVal vCbo As ComboBox)
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim tbl As New dseTransaksiStockInOut.tbl_setting_adjustmentDataTable
        Dim row As dseTransaksiStockInOut.tbl_setting_adjustmentRow

        Try
            vCbo.Items.Clear()
            vCbo.Items.Add(" ")
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl)
            For Each row In tbl.Rows
                vCbo.Items.Add(row.Code)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSettingActive(ByVal vCbo As ComboBox, Optional ByVal vOnlyAktive As Boolean = False)
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim tbl As New dseTransaksiStockInOut.tbl_setting_adjustmentDataTable
        Dim row As dseTransaksiStockInOut.tbl_setting_adjustmentRow

        Try 'update ahdar
            vCbo.Items.Clear()
            vCbo.Items.Add(" ")
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl) > 0 Then 'fill all
                For Each row In tbl.Rows
                    vCbo.Items.Add(row.Code)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Sub

    Public Function GetPrimarySettingByCode(ByVal vCode As String) As Double
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim primary As Nullable(Of Double)
        Try
            primary = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If primary.HasValue = True Then
                If primary.Value > 0 Then
                    Return primary.Value
                Else
                    Return -1
                End If
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCodeSettingByPrimary(ByVal vprimary As String) As String
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim code As String = ""
        Try
            code = dta.ScalarQuery_GetCodeByPrimary(vprimary)
            If String.IsNullOrEmpty(code) = False Then
                Return code
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function GetTipeStockCard(ByVal vTipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut) As ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Try
            Select Case vTipeStockInOut
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_BackProcess
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Process
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Purchase
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_SalesReturn
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Process
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub UpdateStockCard_TransaksiStockInOut(ByVal vPrg As ProgressBar, ByVal vLblProgress As Label)
        Dim DTA As New Inventory_Module.dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout5TableAdapter
        Dim TBL As New Inventory_Module.dseTransaksiStockInOut.tbltransaksi_stockinout5DataTable
        Dim row As Inventory_Module.dseTransaksiStockInOut.tbltransaksi_stockinout5Row
        Dim JenisSC As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim InOut As Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim DescSC As String = ""
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New Inventory_Module.dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim ObjAccClosing As New Core_Module.ClsTranskasiAccClosing

        Try
            DTASC.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTASC.Connection.ConnectionString

            DTA.Fill(TBL, #1/1/1950#)
            vPrg.Value = 0
            vPrg.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                vPrg.Value += 1
                vLblProgress.Text = "Updating stock in out stock card data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString + " - Transaction No :" + row.No_Transaksi

                'cek jenis transaksi stock card
                JenisSC = Me.GetTipeStockCard(row.Tipe_Transaksi)
                'cek in/out
                Select Case row.Tipe_Transaksi
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                        DescSC = "Stock in adjust"
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                        DescSC = "Stock in back process"
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                        DescSC = "Stock in process"
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                        DescSC = "Purchase from " + ObjSupplier.GetSupplierName(row.Primary_Supplier)
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                        DescSC = "Sales return from " + ObjCust.GetCustName(row.Primary_Customer)
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                        DescSC = "Stock out adjust"
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                        DescSC = "Stock out process"
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                        DescSC = "Purchase return to " + ObjSupplier.GetSupplierName(row.Primary_Supplier)
                        Exit Select
                    Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                        InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                        DescSC = "Sales to " + ObjCust.GetCustName(row.Primary_Customer)
                        Exit Select
                End Select

                BatchNo = row.Batch_No
                PrimaryLocation = row.Primary_Storage_Area
                Sec_Qty = "-1"
                PrimaryPdO = -1
                ExpiryDate = row.Expiry_Date

                With row
                    If InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                        Qty_In = .Qty
                        Qty_Out = 0
                    Else
                        Qty_In = 0
                        Qty_Out = row.Qty
                    End If

                    ObjSC.Simpan(.Primary, JenisSC, .No_Transaksi, .Primary_Product, _
                    Qty_In, Qty_Out, DescSC, InOut, .Tanggal_Transaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, .Remarks)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSupplier = Nothing
            ObjCust = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjSC = Nothing
            DTASC.Dispose()
            ObjAccClosing = Nothing
        End Try
    End Sub

    Public Function CetakStockOutPurchaseReturn(ByVal vPrimary As Double, ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim ObjSettings As New Core_Module.ClsTransaksiSetting
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockInOut
        Try
            'If ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.StockOutPurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            '    ObjSettings = Nothing
            Return ObjCustom.CetakStockOutPurchaseReturn(vCrv, vPrimary)
            'End If
            ''cetakan standar SMESS::
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function GetJmlByPrimaryPO(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim vTipeTransaksi As EnumTipeStockInOut = EnumTipeStockInOut.StockIn_Purchase
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByPrimaryPO(CInt(vTipeTransaksi), vPrimaryPO)
            If Retval.HasValue = True Then
                Return Retval.Value
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

    Public Function GetPrimaryPickUpOrder(ByVal vPrimaryStockOutAdjust As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryPickUpOrder(vPrimaryStockOutAdjust)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Sub GetQtySelisihInAdjust(ByVal vDateTo As Date, ByVal vDgv As DataGridView)
        Dim Dta As New dseTransaksiStockInOutTableAdapters.februariTableAdapter
        Dim Tbl As New dseTransaksiStockInOut.februariDataTable
        Dim Row As dseTransaksiStockInOut.februariRow
        Dim DtaSC As New dseTransaksiStockInOutTableAdapters.DataTable2TableAdapter
        Dim TblSC As New dseTransaksiStockInOut.DataTable2DataTable
        Dim RowSC As dseTransaksiStockInOut.DataTable2Row
        Dim DtaProd As New dseTransaksiStockInOutTableAdapters.DataTable7TableAdapter
        Dim TblProd As New dseTransaksiStockInOut.DataTable7DataTable
        Dim RowProd As dseTransaksiStockInOut.DataTable7Row
        Dim TblDgv As New dseTransaksiStockInOut.Tbl_DgvDataTable
        Dim RowDgv As dseTransaksiStockInOut.Tbl_DgvRow
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ListProduct As New List(Of String)
        Dim ExpiredDateStr As String
        Dim Qty_Akhir As Decimal = 0

        Try
            If vDgv.RowCount > 0 Then
                vDgv.Rows.Clear()
            End If
            ListProduct.Clear()
            'Dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaSC.Connection.ConnectionString = Core_Module.GetConstring
            DtaProd.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_CodeNoNotNull(Tbl)
            DtaSC.FillBy_TglSampaiDari(TblSC, vDateTo, CDate("01/01/1950"))
            DtaProd.Fill(TblProd)

            For Each Row In Tbl.Rows
                TblSC.DefaultView.RowFilter = Nothing
                TblSC.DefaultView.RowFilter = "[Product Code] = '" & Row.code_no_smess.ToString & "'"
                If TblSC.DefaultView.Count > 0 Then
                    RowSC = TblSC.DefaultView(0).Row
                    If RowSC.Qty_Total < Row.qty Then
                        RowDgv = TblDgv.NewTbl_DgvRow
                        With RowDgv
                            .Primary_Product = RowSC.Primary_Product
                            .Product_Code = RowSC.Product_Code
                            .Description = RowSC.Description
                            .Unit_Code = RowSC.Unit_Code
                            .Lot_No = RowSC.Batch_Number
                            .Expiry_Date = RowSC.Expiry
                            .Primary_Storage_Area = RowSC.primary_location
                            .Storage_Area_Code = RowSC.Storage_Area_Code
                            .Warehouse_Code = RowSC.Warehouse_Code
                            .Primary_Unit = RowSC.Primary_Unit
                            .Qty_Selisih = Row.qty - RowSC.Qty_Total
                            ._Decimal = RowSC._Decimal
                        End With
                        TblDgv.Rows.Add(RowDgv)
                    End If
                Else
                    TblProd.DefaultView.RowFilter = Nothing
                    TblProd.DefaultView.RowFilter = "[Product Code] = '" & Row.code_no_smess.ToString & "'"
                    If TblProd.DefaultView.Count > 0 Then
                        RowProd = TblProd.DefaultView(0).Row
                        RowDgv = TblDgv.NewTbl_DgvRow
                        With RowDgv
                            .Primary_Product = RowProd.Primary
                            .Product_Code = RowProd.Product_Code
                            .Description = RowProd.Description
                            .Unit_Code = RowProd.Unit_Code
                            .Lot_No = RowProd.Lot_No
                            If RowProd.Primary_Product_Item = -1 Then
                                .Expiry_Date = CDate("01/01/1950")
                            Else
                                .Expiry_Date = RowProd.Expiry_Date
                            End If
                            .Primary_Storage_Area = RowProd.Primary_Storage_Area
                            .Storage_Area_Code = RowProd.Storage_Area_Code
                            .Warehouse_Code = RowProd.Warehouse_Code
                            .Primary_Unit = RowProd.Primary_Unit
                            .Qty_Selisih = Row.qty
                            ._Decimal = RowProd._Decimal
                        End With
                        TblDgv.Rows.Add(RowDgv)
                    Else
                        ListProduct.Add(Row.code_no_smess)
                    End If
                End If
            Next

            If ListProduct.Count > 0 Then
                Dim msg As String = ""
                For intA As Integer = 0 To ListProduct.Count - 1
                    If msg = "" Then
                        msg &= ListProduct.Item(intA)
                    Else
                        msg &= ", " & ListProduct.Item(intA)
                    End If
                Next
                If msg <> "" Then
                    MsgBox("There are list of code no smess were not found on tblmaster product : " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                End If
            End If

            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each RowDgv In TblDgv.Rows
                With RowDgv
                    If .Expiry_Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
                    End If
                    Qty_Akhir = ObjCD.DigitGrouping(.Qty_Selisih, ._Decimal)
                    If Qty_Akhir > 0 Then
                        vDgv.Rows.Add("", "", "", .Product_Code, .Description, Qty_Akhir, .Unit_Code, -1, "", .Lot_No, _
                        ExpiredDateStr, .Storage_Area_Code, .Warehouse_Code, "", -1, .Primary_Product, .Primary_Storage_Area, _
                        .Primary_Unit, -1, -1, 0)
                    End If
                End With
            Next
            'set ulang nomor
            ObjCD.SetNumber(vDgv)

            'tampilkan batch
            vDgv.Columns(9).Visible = True

            'tampilkan expire date
            vDgv.Columns(10).Visible = True
            'tampilkan lokasi dan warehouse
            vDgv.Columns(11).Visible = True
            vDgv.Columns(12).Visible = True

            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            DtaSC.Dispose()
            DtaProd.Dispose()
            Tbl.Dispose()
            TblSC.Dispose()
            TblProd.Dispose()
            ObjProduct = Nothing
            ObjSC = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Public Function UpdateDecProductQty() As Boolean
        Dim Dta As New dseTransaksiStockInOutTableAdapters.februariTableAdapter
        Dim Tbl As New dseTransaksiStockInOut.februariDataTable
        Dim Row As dseTransaksiStockInOut.februariRow
        'Dim DtaP As New Core_Module.dseMasterProductTableAdapters.tblmaster_productTableAdapter
        'Dim TblP As New Core_Module.dseMasterProduct.tblmaster_productDataTable
        'Dim RowP As Core_Module.dseMasterProduct.tblmaster_productRow
        Dim DtaP As New dseTransaksiStockInOutTableAdapters.tblmaster_product2TableAdapter
        Dim TblP As New dseTransaksiStockInOut.tblmaster_product2DataTable
        Dim RowP As dseTransaksiStockInOut.tblmaster_product2Row
        Dim ListProduct As New List(Of String)
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DtaP.Connection.ConnectionString = Core_Module.GetConstring
            ListProduct.Clear()
            Dta.FillBy_CodeNoNotNull(Tbl)
            DtaP.Fill(TblP)

            'update decimal qty pada tabel februari dan sesuaikan decimal nya dengan decimal product di smess
            For Each Row In Tbl.Rows
                TblP.DefaultView.RowFilter = Nothing
                TblP.DefaultView.RowFilter = "[Product Code] = '" & Row.code_no_smess.ToString & "'"
                If TblP.DefaultView.Count > 0 Then
                    RowP = TblP.DefaultView(0).Row
                    Row.qty = ObjCD.DigitGrouping(Row.qty, RowP._Decimal)
                    Dta.UpdateQuery_ProductQty(Row.qty, Row.code_no_smess)
                Else
                    ListProduct.Add(Row.code_no_smess)
                End If
            Next

            If ListProduct.Count > 0 Then
                Dim msg As String = ""
                For intA As Integer = 0 To ListProduct.Count - 1
                    If msg = "" Then
                        msg &= ListProduct.Item(intA)
                    Else
                        msg &= ", " & ListProduct.Item(intA)
                    End If
                Next
                If msg <> "" Then
                    MsgBox("There are list of code no smess were not found on tblmaster stock card : " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            DtaP.Dispose()
            TblP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    '# u/ Quick Input 3
    Public Function Isitbl_ByBatch(ByVal tbldummy As dseTransaksiStockInOut.tblmaster_productqtyDataTable, ByVal batchNo As String) As Boolean
        Dim dta As New dseTransaksiStockInOutTableAdapters.tblmaster_productqtyTableAdapter
        Try
            tbldummy.Rows.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbldummy, batchNo.ToString.Trim)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbldummy.Dispose()
        End Try
    End Function

    Public Function getJmlByBatch(ByVal vBatch As String) As Integer
        Dim dta As New dseTransaksiStockInOutTableAdapters.tblmaster_productqtyTableAdapter
        Dim retval As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetJml_ProductByBatch(vBatch)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval.Value
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahStockInPurchaseQuick(ByVal vPrimaryStockInPurchase As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOut_SetAmountTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New dseTransaksiStockInOut_SetAmount.tbltrans_aopDataTable
        Dim row As dseTransaksiStockInOut_SetAmount.tbltrans_aopRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryStockInPurchase)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.Quick_Purchase = 1 Then
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
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimaryPurchaseReturn(ByVal vPrimaryStockOutPurchaseReturn As Double) As Double
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryPOPurchase(vPrimaryStockOutPurchaseReturn)

            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetPrimaryStockInPurchase(ByVal vPrimaryStockInPurchase As Double) As Double
        Return Me.GetPrimaryPurchaseReturn(vPrimaryStockInPurchase)
    End Function


    'update by ahdar: 2024-2-29
    Public Function ApakahAdjustmentTypeAktif(ByVal vPrimarySetting As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim Retval As New Nullable(Of Integer)
        Try
            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            'If vPrimarySetting < 0 Then
            '    Return True
            'End If

            'Retval = DTA.ScalarQuery_GetIsActiveByPrimary(vPrimarySetting)
            'If Retval.HasValue = True Then
            '    If Retval.Value = 0 Then
            '        Return False
            '    Else
            '        Return True
            '    End If
            'End If
            'Return False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function
   
End Class
