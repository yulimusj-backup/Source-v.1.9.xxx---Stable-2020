Option Explicit On
Option Strict Off

Imports System.IO

Imports Microsoft.VisualBasic.FileIO
Public Class clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mProductCode As String = ""
    Dim mPrimaryCategory As Double = -1
    Dim mSpecification As String = ""
    Dim mPrimaryUnit As Double = -1
    Dim mPrimaryMasterUnit As Double = -1
    Dim mQtyPerMasterUnit As Integer = -1
    Dim mPrimarySecondaryUnit As Double = -1
    Dim mBatchNoControl As Boolean = False
    Dim mLocationControl As Boolean = False
    Dim mUseSecondaryUnits As Boolean = False
    Dim mMinimumStock As Double = -1
    Dim mLeadTime As Integer = -1
    Dim mMinimumPurchaseQty As Double = -1
    Dim mDesc As String = ""
    Dim mPdoNoControlled As Boolean = False
    Dim mExpiryDateControlled As Boolean = False
    Dim mCategoryCodeBounded As Boolean = False
    Dim mCategoryDescBounded As Boolean = False
    Dim mDecimal As Integer = 0
    Dim mUseSizing As Integer = 0
    Dim mSize As Integer = 36
    Dim mBarcode As String = ""
    Dim mPrimaryAkunInventory As Double = -1
    Dim mPrimaryAkunWIP As Double = -1
    Dim mTipe As EnumTipeProduct = EnumTipeProduct.Goods
    Dim mPrimarySalesAccount As Double = -1
    Dim mPrimarySalesReturnAccount As Double = -1
    Dim mPrimaryCOGSAccount As Double = -1
    Dim mPrimaryAdjustmentAccount As Double = -1
    Dim mPrimaryStockTakingAccount As Double = -1
    Dim mTipeDependency As EnumTipeDependency = EnumTipeDependency.End_Start1
    Dim mDelayDependency As Double = 0

    'berkaitan dengan bill of material
    Dim mPrimaryProcess As Double = -1
    Dim mQtyRef As Decimal = 1
    Dim mGoodPercentage As Decimal = 100

    'berkaitan dengann verifikasi akun product
    Dim mIsVerified As Integer = 0
    Dim mVerifiedBy As String = ""
    Dim mVerifiedTime As Date = Now

    'berkaitan dengan tipe goods
    Dim mTypeGoods As Integer = EnumTypeGoods.FG
    Dim mTolerance As Decimal = 0

    'add weight & volume
    Dim mNettWeight As Decimal = 0
    Dim mGrossWeight As Decimal = 0
    Dim mPrimaryUnitWeight As Double = -1
    Dim mPrecisionWeight As Integer = 0
    Dim mNettVolume As Decimal = 0
    Dim mGrossVolume As Decimal = 0
    Dim mPrimaryUnitVolume As Double = -1
    Dim mPrecisionVolume As Integer = 0

    Dim mPrimaryTaxProductRef As Double = -1

    Public Property PrimaryTaxProductRef() As Double
        Get
            Return Me.mPrimaryTaxProductRef
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryTaxProductRef = value
        End Set
    End Property

    Public Property NettWeight() As Decimal
        Get
            Return Me.mNettWeight
        End Get
        Set(ByVal value As Decimal)
            Me.mNettWeight = value
        End Set
    End Property

    Public Property GrossWeight() As Decimal
        Get
            Return Me.mGrossWeight
        End Get
        Set(ByVal value As Decimal)
            Me.mGrossWeight = value
        End Set
    End Property

    Public Property PrimaryUnitWeight() As Double
        Get
            Return Me.mPrimaryUnitWeight
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryUnitWeight = value
        End Set
    End Property

    Public Property PrecisionWeight() As Integer
        Get
            Return Me.mPrecisionWeight
        End Get
        Set(ByVal value As Integer)
            Me.mPrecisionWeight = value
        End Set
    End Property

    Public Property NettVolume() As Decimal
        Get
            Return Me.mNettVolume
        End Get
        Set(ByVal value As Decimal)
            Me.mNettVolume = value
        End Set
    End Property

    Public Property GrossVolume() As Decimal
        Get
            Return Me.mGrossVolume
        End Get
        Set(ByVal value As Decimal)
            Me.mGrossVolume = value
        End Set
    End Property

    Public Property PrimaryUnitVolume() As Double
        Get
            Return Me.mPrimaryUnitVolume
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryUnitVolume = value
        End Set
    End Property

    Public Property PrecisionVolume() As Integer
        Get
            Return Me.mPrecisionVolume
        End Get
        Set(ByVal value As Integer)
            Me.mPrecisionVolume = value
        End Set
    End Property
    '------------------------------------------------------------------------------------------------------------
    Public Property Tolerance() As Decimal
        Get
            Return Me.mTolerance
        End Get
        Set(ByVal value As Decimal)
            Me.mTolerance = value
        End Set
    End Property

    Public Enum EnumTipeDependency As Integer
        End_Start1
        Start_Start1
        Start_Start_Strict1
    End Enum

    Public Enum EnumDefaultFindBy As Integer
        ByCode
        ByDesc
    End Enum

    Public Enum EnumDefaultFindBy_SettingUser As Integer
        ByGlobalSetting
        ByCode
        ByDesc
    End Enum

    Public Enum EnumTypeGoods As Integer
        FG
        Sub_Assy
        WIP_
        Raw_Material
    End Enum

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian


    Public Enum EnumProcurementStrategy As Integer
        Build_To_Order
        Build_To_Stock
        Both
    End Enum

    Public Enum EnumSupplierSelectionMode
        The_cheapest_in_the_last_3_order
        Prefered_Supplier
        Most_Recent_PO
    End Enum

    Public Enum EnumOptimumStockStrategy As Integer
        Based_On_Usage
        User_Defined
    End Enum

    Public Enum EnumSafetyStockStrategy As Integer
        Based_On_Usage
        User_Defined
    End Enum

    Public Enum EnumLeadTimeStrategy As Integer
        Based_On_Usage
        User_Defined
    End Enum

    Public Enum EnumProcurementType As Integer
        Buy
        Produce
    End Enum

    Public Enum EnumOrderStrategy As Integer
        Order_To_Min_Qty
        Order_To_Optimum_Qty
        Order_To_Max_Qty
        Order_To_Safety_Stock
        Order_By_Minimum_Order_Qty
        Order_By_Optimum_Order_Qty
        Order_By_Maximum_Order_Qty
    End Enum

    Public Enum EnumTipeProduct As Integer
        Goods
        Services
    End Enum

    Public Property TypeGoods() As EnumTypeGoods
        Get
            Return Me.mTypeGoods
        End Get
        Set(ByVal value As EnumTypeGoods)
            mTypeGoods = value
        End Set
    End Property

    Public Property IsVerified() As Integer
        Get
            Return Me.mIsVerified
        End Get
        Set(ByVal value As Integer)
            mIsVerified = value
        End Set
    End Property

    Public Property VerifiedBy() As String
        Get
            Return Me.mVerifiedBy
        End Get
        Set(ByVal value As String)
            mVerifiedBy = value
        End Set
    End Property

    Public Property VerifiedTime() As Date
        Get
            Return Me.mVerifiedTime
        End Get
        Set(ByVal value As Date)
            mVerifiedTime = value
        End Set
    End Property

    Public Property PrimaryProcess() As Double
        Get
            Return Me.mPrimaryProcess
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryProcess = value
        End Set
    End Property

    Public Property QtyRef() As Decimal
        Get
            Return Me.mQtyRef
        End Get
        Set(ByVal value As Decimal)
            mQtyRef = value
        End Set
    End Property

    Public Property GoodPercentage() As Decimal
        Get
            Return Me.mGoodPercentage
        End Get
        Set(ByVal value As Decimal)
            mGoodPercentage = value
        End Set
    End Property

    Public Property ProductType() As EnumTipeProduct
        Get
            Return Me.mTipe
        End Get
        Set(ByVal value As EnumTipeProduct)
            mTipe = value
        End Set
    End Property
    Public Property PrimaryAkunStockTaking() As Double
        Get
            Return Me.mPrimaryStockTakingAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryStockTakingAccount = value
        End Set
    End Property

    Public Property PrimaryAkunAdjustment() As Double
        Get
            Return Me.mPrimaryAdjustmentAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAdjustmentAccount = value
        End Set
    End Property

    Public Property PrimaryAkunCOGS() As Double
        Get
            Return Me.mPrimaryCOGSAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryCOGSAccount = value
        End Set
    End Property

    Public Property PrimaryAkunSalesReturn() As Double
        Get
            Return Me.mPrimarySalesReturnAccount
        End Get
        Set(ByVal value As Double)
            mPrimarySalesReturnAccount = value
        End Set
    End Property

    Public Property PrimaryAkunSales() As Double
        Get
            Return Me.mPrimarySalesAccount
        End Get
        Set(ByVal value As Double)
            mPrimarySalesAccount = value
        End Set
    End Property

    Public Property PrimaryAkunInventory() As Double
        Get
            Return Me.mPrimaryAkunInventory
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunInventory = value
        End Set
    End Property

    Public Property PrimaryAkunWIP() As Double
        Get
            Return Me.mPrimaryAkunWIP
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunWIP = value
        End Set
    End Property

    Public Property Barcode() As String
        Get
            Return Me.mBarcode
        End Get
        Set(ByVal value As String)
            mBarcode = value
        End Set
    End Property
    Public Property Size() As Integer
        Get
            Return Me.mSize
        End Get
        Set(ByVal value As Integer)
            Me.mSize = value
        End Set
    End Property
    Public Property UseSizing() As Boolean
        Get
            If Me.mUseSizing = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mUseSizing = 1
            Else
                Me.mUseSizing = 0
            End If
        End Set
    End Property
    Public Property AngkaDecimal() As Integer
        Get
            Return Me.mDecimal
        End Get
        Set(ByVal value As Integer)
            mDecimal = value
        End Set
    End Property

    Public Property CategoryDescBounded() As Boolean
        Get
            Return Me.mCategoryDescBounded
        End Get
        Set(ByVal value As Boolean)
            Me.mCategoryDescBounded = value
        End Set
    End Property

    Public Property CategoryCodeBounded() As Boolean
        Get
            Return Me.mCategoryCodeBounded
        End Get
        Set(ByVal value As Boolean)
            mCategoryCodeBounded = value
        End Set
    End Property

    Public Property ExpiryDateControlled() As Boolean
        Get
            Return Me.mExpiryDateControlled
        End Get
        Set(ByVal value As Boolean)
            mExpiryDateControlled = value
        End Set
    End Property

    Public Property PdONumberControlled() As Boolean
        Get
            Return Me.mPdoNoControlled
        End Get
        Set(ByVal value As Boolean)
            mPdoNoControlled = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property UseSecondaryUnits() As Boolean
        Get
            Return Me.mUseSecondaryUnits
        End Get
        Set(ByVal value As Boolean)
            mUseSecondaryUnits = value
        End Set
    End Property

    Public Property MinimumPurchaseQty() As Double
        Get
            Return Me.mMinimumPurchaseQty
        End Get
        Set(ByVal value As Double)
            mMinimumPurchaseQty = value
        End Set
    End Property

    Public Property LeadTime() As Integer
        Get
            Return Me.mLeadTime
        End Get
        Set(ByVal value As Integer)
            mLeadTime = value
        End Set
    End Property

    Public Property MinimumStock() As Double
        Get
            Return Me.mMinimumStock
        End Get
        Set(ByVal value As Double)
            Me.mMinimumStock = value
        End Set
    End Property

    Public Property LocationControl() As Boolean
        Get
            Return Me.mLocationControl
        End Get
        Set(ByVal value As Boolean)
            mLocationControl = value
        End Set
    End Property

    Public Property BatchNoControl() As Boolean
        Get
            Return Me.mBatchNoControl
        End Get
        Set(ByVal value As Boolean)
            mBatchNoControl = value
        End Set
    End Property

    Public Property PrimarySecondaryUnit() As Double
        Get
            Return Me.mPrimarySecondaryUnit
        End Get
        Set(ByVal value As Double)
            mPrimarySecondaryUnit = value
        End Set
    End Property

    Public Property QtyPerMasterUnit() As Double
        Get
            Return Me.mQtyPerMasterUnit
        End Get
        Set(ByVal value As Double)
            mQtyPerMasterUnit = value
        End Set
    End Property

    Public Property PrimaryMasterUnit() As Double
        Get
            Return Me.mPrimaryMasterUnit
        End Get
        Set(ByVal value As Double)
            mPrimaryMasterUnit = value
        End Set
    End Property

    Public Property PrimaryUnit() As Double
        Get
            Return Me.mPrimaryUnit
        End Get
        Set(ByVal value As Double)
            mPrimaryUnit = value
        End Set
    End Property

    Public Property Specification() As String
        Get
            Return Me.mSpecification
        End Get
        Set(ByVal value As String)
            mSpecification = value
            If Me.mPrimaryCategory <> -1 And Me.mCategoryDescBounded = True Then
                Dim obj As New ClsMasterCategory
                Me.mDesc = obj.GetDesc(Me.mPrimaryCategory) + " " + Me.mSpecification
                obj = Nothing
            Else
                Me.mDesc = Me.mSpecification
            End If
        End Set
    End Property

    Public Property PrimaryCategory() As Double
        Get
            Return Me.mPrimaryCategory
        End Get
        Set(ByVal value As Double)
            mPrimaryCategory = value
            If Me.mPrimaryCategory <> -1 And Me.mCategoryDescBounded = True Then
                Dim obj As New ClsMasterCategory
                Me.mDesc = obj.GetDesc(Me.mPrimaryCategory) + " " + Me.mSpecification
                obj = Nothing
            Else
                Me.mDesc = Me.mSpecification
            End If
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

    Public Property ProductCode() As String
        Get
            Return Me.mProductCode
        End Get
        Set(ByVal value As String)
            mProductCode = value
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

    Public Property TipeDependency() As EnumTipeDependency
        Get
            Return Me.mTipeDependency
        End Get
        Set(ByVal value As EnumTipeDependency)
            Me.mTipeDependency = value
        End Set
    End Property

    Public Property DelayDependency() As Double
        Get
            Return Me.mDelayDependency
        End Get
        Set(ByVal value As Double)
            Me.mDelayDependency = value
        End Set
    End Property

    Public Overridable Function SimpanBaru(ByVal vDGVKompetensi As DataGridView, ByVal vUserCode As String, ByVal pbx As PictureBox, _
    ByVal dgvPrice As Windows.Forms.DataGridView, ByVal vDGVRawMaterial As DataGridView, ByVal vFrm As frmMasterProduct, _
    ByVal vDGVMultiUnit As DataGridView, ByVal vDGVAttachment As DataGridView, ByVal vDGVWorkUnit As DataGridView) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProduct, vUserCode)
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
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vDGVKompetensi, vUserCode, pbx, dgvPrice, vDGVRawMaterial, vFrm, vDGVMultiUnit, vDGVAttachment, vDGVWorkUnit)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(750)
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

    Public Function SimpanBaru1(ByVal vDGVKompetensi As DataGridView, ByVal vUserCode As String, ByVal pbx As PictureBox, _
    ByVal dgvPrice As Windows.Forms.DataGridView, ByVal vDGVRawMaterial As DataGridView, _
    ByVal vFrm As frmMasterProduct, ByVal vDGVMultiUnit As DataGridView, ByVal vDGVAttachment As DataGridView, _
    ByVal vDGVWorkUnit As DataGridView) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim row As DataGridViewRow
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter
        Dim AllowUserToAddRow As Boolean = False
        Dim DTARM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bomTableAdapter
        Dim DTAUnit As New dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim DTAAttachment As New dseMasterProductTableAdapters.tblmaster_product_item_attachmentTableAdapter
        Dim DTAWorkUnit As New dseMasterProductTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim DtaInsert As New dseMasterProduct2TableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTARM.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAAttachment.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWorkUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
            DtaInsert.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(vDGVRawMaterial, vFrm) = False Then
                Return False
            End If

            DTARM.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTARM.Connection.ConnectionString
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString

            If pbx.ImageLocation <> "" Then
                '' Create a new stream to load this photo into
                Dim myStream As IO.FileStream = New FileStream(pbx.ImageLocation, FileMode.Open, _
                FileAccess.Read)
                '' Create a buffer to hold the stream of bytes
                Dim myImageBuffer(myStream.Length) As Byte
                '' Read the bytes from this stream and put it into the image buffer
                myStream.Read(myImageBuffer, 0, Convert.ToInt32(myStream.Length))
                '' Close the stream
                myStream.Close()

                If DtaInsert.InsertQuery2(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, Me.mPrimaryMasterUnit, _
                Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, Me.mLocationControl, _
                Me.mUseSecondaryUnits, myImageBuffer, vUserCode, Now, vUserCode, Now, Now, Me.mMinimumStock, _
                Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, _
                Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimaryAdjustmentAccount, Me.mPrimaryStockTakingAccount, Me.mPrimaryCOGSAccount, _
                Me.mPrimarySalesAccount, Me.mPrimarySalesReturnAccount, Me.mPrimaryProcess, Me.mQtyRef, Me.mGoodPercentage, Me.mTypeGoods, _
                Me.mTipeDependency, Me.mDelayDependency, Me.mTolerance, Me.mNettWeight, Me.mGrossWeight, Me.mPrimaryUnitWeight, _
                Me.mPrecisionWeight, Me.mNettVolume, Me.mGrossVolume, Me.mPrimaryUnitVolume, Me.mPrecisionVolume, Me.mPrimaryTaxProductRef) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mProductCode, Me.mPrimaryCategory).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    'input raw material
                    AllowUserToAddRow = vDGVRawMaterial.AllowUserToAddRows
                    vDGVRawMaterial.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVRawMaterial.Rows
                        DTARM.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value, rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value, _
                        rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value, rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value, rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value)
                    Next
                    vDGVRawMaterial.AllowUserToAddRows = AllowUserToAddRow

                    'input multi unit
                    AllowUserToAddRow = vDGVMultiUnit.AllowUserToAddRows
                    vDGVMultiUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVMultiUnit.Rows
                        DTAUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value, rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value, rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value) 'add weight & volume
                    Next
                    vDGVMultiUnit.AllowUserToAddRows = AllowUserToAddRow

                    'input custom value
                    Me.SimpanItemKompetensi(vDGVKompetensi, Me.mPrimary, vFrm)

                    'input attachment
                    AllowUserToAddRow = vDGVAttachment.AllowUserToAddRows
                    vDGVAttachment.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVAttachment.Rows
                        DTAAttachment.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value, rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColAtt_Description.Index).Value)
                    Next
                    vDGVAttachment.AllowUserToAddRows = AllowUserToAddRow

                    'input work unit
                    AllowUserToAddRow = vDGVWorkUnit.AllowUserToAddRows
                    vDGVWorkUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVWorkUnit.Rows
                        DTAWorkUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, _
                            rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value)
                    Next
                    vDGVWorkUnit.AllowUserToAddRows = AllowUserToAddRow

                    'return value
                    Return True
                Else
                    Return False
                End If
            Else
                If DTA.InsertQuery(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, Me.mPrimaryMasterUnit, _
                Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, Me.mLocationControl, _
                Me.mUseSecondaryUnits, Nothing, vUserCode, Now, vUserCode, Now, Now, Me.mMinimumStock, Me.mLeadTime, _
                Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, Me.mCategoryCodeBounded, _
                Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, _
                Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimaryAdjustmentAccount, Me.mPrimaryStockTakingAccount, Me.mPrimaryCOGSAccount, _
                Me.mPrimarySalesAccount, Me.mPrimarySalesReturnAccount, Me.mPrimaryProcess, Me.mQtyRef, Me.mGoodPercentage, Me.mTypeGoods, _
                Me.mTipeDependency, Me.mDelayDependency, Me.mTolerance, Me.mNettWeight, Me.mGrossWeight, Me.mPrimaryUnitWeight, _
                Me.mPrecisionWeight, Me.mNettVolume, Me.mGrossVolume, Me.mPrimaryUnitVolume, Me.mPrecisionVolume) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mProductCode, Me.mPrimaryCategory).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    'input raw material
                    AllowUserToAddRow = vDGVRawMaterial.AllowUserToAddRows
                    vDGVRawMaterial.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVRawMaterial.Rows
                        DTARM.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value, rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value, _
                        rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value, rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value, rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value)
                    Next
                    vDGVRawMaterial.AllowUserToAddRows = AllowUserToAddRow

                    'input multi unit
                    AllowUserToAddRow = vDGVMultiUnit.AllowUserToAddRows
                    vDGVMultiUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVMultiUnit.Rows
                        DTAUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value, rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value, rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value) 'add weight & volume))
                    Next
                    vDGVMultiUnit.AllowUserToAddRows = AllowUserToAddRow

                    'input attachment
                    AllowUserToAddRow = vDGVAttachment.AllowUserToAddRows
                    vDGVAttachment.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVAttachment.Rows
                        DTAAttachment.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value, rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColAtt_Description.Index).Value)
                    Next
                    vDGVAttachment.AllowUserToAddRows = AllowUserToAddRow

                    'input custom value
                    Me.SimpanItemKompetensi(vDGVKompetensi, Me.mPrimary, vFrm)

                    'input work unit
                    AllowUserToAddRow = vDGVWorkUnit.AllowUserToAddRows
                    vDGVWorkUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVWorkUnit.Rows
                        DTAWorkUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, _
                            rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value)
                    Next
                    vDGVWorkUnit.AllowUserToAddRows = AllowUserToAddRow

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTARM.Dispose()
            DTA1.Dispose()
            DTAUnit.Dispose()
            DTAAttachment.Dispose()
            DTAWorkUnit.Dispose()
            DtaInsert.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan nilai property yang dikontrol untuk suatu product
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="Location"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <param name="SecondaryUnit"></param>
    ''' <remarks></remarks>
    Public Sub GetProductProperties(ByVal vPrimaryProduct As Double, ByRef Location As Boolean, ByRef BatchNo As Boolean, _
    ByRef ExpiryDate As Boolean, ByRef SecondaryUnit As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    BatchNo = True
                Else
                    BatchNo = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    Location = True
                Else
                    Location = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    ExpiryDate = True
                Else
                    ExpiryDate = False
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    SecondaryUnit = True
                Else
                    SecondaryUnit = False
                End If
            Else
                Location = False
                BatchNo = False
                ExpiryDate = False
                SecondaryUnit = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub GetProductProperties1(ByVal vPrimaryProduct As Double, ByRef Location As Boolean, ByRef BatchNo As Boolean, _
    ByRef ExpiryDate As Boolean, ByRef SecondaryUnit As Boolean, ByRef decimalPoint As Integer)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    BatchNo = True
                Else
                    BatchNo = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    Location = True
                Else
                    Location = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    ExpiryDate = True
                Else
                    ExpiryDate = False
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    SecondaryUnit = True
                Else
                    SecondaryUnit = False
                End If
                decimalPoint = row._Decimal '2026-03-23: tambah mengembalikan decimal point supaya ga perlu panggil lagi jika butuh. terkait pembulatan stock in adjust WD Indonesia
            Else
                Location = False
                BatchNo = False
                ExpiryDate = False
                SecondaryUnit = False
                decimalPoint = 0 '2026-03-23: tambah mengembalikan decimal point supaya ga perlu panggil lagi jika butuh. terkait pembulatan stock in adjust WD Indonesia
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub IsiCheckBoxControll(ByVal vPrimaryProduct As Double, ByVal chkLocation As CheckBox, ByVal chkBatchNo As CheckBox, ByVal chkExpiryDate As CheckBox, _
    ByVal chkSecondaryUnit As CheckBox, ByVal chkPdoControlled As CheckBox, ByVal TxtPrimaryUnit As TextBox, _
    ByVal TxtPrimarySecondaryUnits As TextBox)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    chkBatchNo.Checked = True
                Else
                    chkBatchNo.Checked = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    chkLocation.Checked = True
                Else
                    chkLocation.Checked = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    chkExpiryDate.Checked = True
                Else
                    chkExpiryDate.Checked = False
                End If
                If CBool(row.PdOControlled) = True Then
                    chkPdoControlled.Checked = True
                Else
                    chkPdoControlled.Checked = False
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    chkSecondaryUnit.Checked = True
                Else
                    chkSecondaryUnit.Checked = False
                End If
                TxtPrimarySecondaryUnits.Text = row.Primary_Secondary_Unit
                TxtPrimaryUnit.Text = row.Primary_Unit
            Else
                chkLocation.Checked = False
                chkBatchNo.Checked = False
                chkExpiryDate.Checked = False
                chkSecondaryUnit.Checked = False
                chkPdoControlled.Checked = False
                TxtPrimarySecondaryUnits.Text = "-1"
                TxtPrimaryUnit.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function IsiListBoxProduct(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterProduct.tblmaster_product2DataTable) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiCheckBoxControll(ByVal vPrimaryProduct As Double, ByVal chkLocation As CheckBox, _
    ByVal chkBatchNo As CheckBox, ByVal chkExpiryDate As CheckBox, ByVal TxtPrimaryUnit As TextBox)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    chkBatchNo.Checked = True
                Else
                    chkBatchNo.Checked = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    chkLocation.Checked = True
                Else
                    chkLocation.Checked = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    chkExpiryDate.Checked = True
                Else
                    chkExpiryDate.Checked = False
                End If
                TxtPrimaryUnit.Text = row.Primary_Unit
            Else
                chkLocation.Checked = False
                chkBatchNo.Checked = False
                chkExpiryDate.Checked = False
                TxtPrimaryUnit.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterProductTableAdapters.DataTable1TableAdapter

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
            ElseIf vSearchBy = EnumSearchCriteria.ByProductSpec Then
                If DTA.ScalarQuery_GetJml3(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3(vKriteria + "%").Value
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

    Public Function GetJml2(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterProductTableAdapters.DataTable1TableAdapter
        Dim DTA2 As New dseMasterProduct2TableAdapters.tblmaster_product_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml_1(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml_1(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2_1(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml2_1(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductSpec Then
                If DTA.ScalarQuery_GetJml3_1(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3_1(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByTaxProductCode Then
                If DTA2.ScalarQuery_GetJmlByTaxProductCodeLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByTaxProductDesc Then
                If DTA2.ScalarQuery_GetJmlByTaxProductDescLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3(vKriteria + "%").Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Function ApakahKodeTerakhirMenggunakanBarcode() As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim vPrimary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryMax.HasValue = False Then
                Return True
            Else
                vPrimary = DTA.ScalarQuery_GetPrimaryMax.Value
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetBarcode(vPrimary)) = True Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVKompetensi As DataGridView, ByVal vUserCode As String, ByVal pbx As PictureBox, ByVal dgvPrice As DataGridView, _
    ByVal vDGVRawMaterial As DataGridView, ByVal vFrm As frmMasterProduct, ByVal vDGVMultiUnit As DataGridView, ByVal vDGVAttachment As DataGridView, _
    ByVal vDGVWorkUnit As DataGridView) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProduct, vUserCode)


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
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vDGVKompetensi, vUserCode, pbx, dgvPrice, vDGVRawMaterial, vFrm, vDGVMultiUnit, vDGVAttachment, vDGVWorkUnit)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
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
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vDGVKompetensi As DataGridView, ByVal vUserCode As String, ByVal pbx As PictureBox, ByVal dgvPrice As DataGridView, _
    ByVal vDGVRawMaterial As DataGridView, ByVal vFrm As frmMasterProduct, ByVal vDGVMultiUnit As DataGridView, _
    ByVal vDGVAttachment As DataGridView, ByVal vDGVWorkUnit As DataGridView) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter
        Dim AllowUserToAddRow As Boolean = False
        Dim DTARM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bomTableAdapter
        Dim row As DataGridViewRow
        Dim DTAUnit As New dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim DTAProduct As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBLProduct As New dseMasterProduct.tblmaster_product3DataTable
        Dim rowProduct As dseMasterProduct.tblmaster_product3Row
        Dim AkunBerbeda As Boolean = False
        Dim DTAAttachment As New dseMasterProductTableAdapters.tblmaster_product_item_attachmentTableAdapter
        Dim DTAWorkUnit As New dseMasterProductTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim DTAUpdate As New dseMasterProduct2TableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTARM.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAProduct.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAAttachment.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWorkUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAUpdate.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(vDGVRawMaterial, vFrm) = False Then
                Return False
            End If

            'cek apakah ada perubahan akun product
            DTAProduct.FillBy(TBLProduct, Me.mPrimary)
            If TBLProduct.Rows.Count > 0 Then
                rowProduct = TBLProduct.Rows(0)
                If rowProduct.Primary_Adjustment_Account <> Me.mPrimaryAdjustmentAccount Or _
                rowProduct.Primary_COGS_Account <> Me.mPrimaryCOGSAccount Or _
                rowProduct.Primary_Inventory_Account <> Me.mPrimaryAkunInventory Or _
                rowProduct.Primary_Sales_Account <> Me.mPrimarySalesAccount Or _
                rowProduct.Primary_Sales_Return_Account <> Me.mPrimarySalesReturnAccount Or _
                rowProduct.Primary_Stock_Taking_Account <> Me.mPrimaryStockTakingAccount Then
                    AkunBerbeda = True
                End If
            End If

            'cek apakah ada perubahan pada batch no controled, expiry date controled, location controlled,
            'PdO controlled, use secondary units
            If pbx.ImageLocation <> "" Then 'ada perubahan data gambar
                '' Create a new stream to load this photo into
                Dim myStream As IO.FileStream = New FileStream(pbx.ImageLocation, FileMode.Open, _
                FileAccess.Read)
                '' Create a buffer to hold the stream of bytes
                Dim myImageBuffer(myStream.Length) As Byte
                '' Read the bytes from this stream and put it into the image buffer
                myStream.Read(myImageBuffer, 0, Convert.ToInt32(myStream.Length))
                '' Close the stream
                myStream.Close()

                If DTAUpdate.UpdateQuery2(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                Me.mLocationControl, Me.mUseSecondaryUnits, myImageBuffer, vUserCode, Now, Now, Me.mMinimumStock, _
                Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, _
                Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimaryAdjustmentAccount, Me.mPrimaryStockTakingAccount, _
                Me.mPrimaryCOGSAccount, Me.mPrimarySalesAccount, Me.mPrimarySalesReturnAccount, Me.mPrimaryProcess, _
                Me.mQtyRef, Me.mGoodPercentage, Me.mTypeGoods, Me.mTipeDependency, Me.mDelayDependency, Me.mTolerance, Me.mNettWeight, _
                Me.mGrossWeight, Me.mPrimaryUnitWeight, Me.mPrecisionWeight, Me.mNettVolume, Me.mGrossVolume, Me.mPrimaryUnitVolume, _
                Me.mPrecisionVolume, Me.mPrimaryTaxProductRef, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'jika akun produknya ada yang berubah, maka update data verifikas menjadi tidak verified
                    If AkunBerbeda = True Then
                        DTA.UpdateQuery_SetAccountVerification(0, "", #1/1/1950#, Me.mPrimary)
                    End If

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data lama
                    DTA1.Delete(Me.mPrimary)
                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    'hapus raw material
                    DTARM.Delete(Me.mPrimary)
                    'input raw material
                    AllowUserToAddRow = vDGVRawMaterial.AllowUserToAddRows
                    vDGVRawMaterial.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVRawMaterial.Rows
                        DTARM.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value, rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value, _
                        rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value, rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value, rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value)
                    Next
                    vDGVRawMaterial.AllowUserToAddRows = AllowUserToAddRow

                    'input multi unit
                    DTAUnit.Delete(Me.mPrimary)
                    AllowUserToAddRow = vDGVMultiUnit.AllowUserToAddRows
                    vDGVMultiUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVMultiUnit.Rows
                        DTAUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value, rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value, rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value, _
                        rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value) 'add weight & volume)
                    Next
                    vDGVMultiUnit.AllowUserToAddRows = AllowUserToAddRow

                    'custom value
                    Me.HapusItemKompetensi(Me.mPrimary)
                    Me.SimpanItemKompetensi(vDGVKompetensi, Me.mPrimary, vFrm)

                    'input attachment
                    DTAAttachment.Delete(Me.mPrimary)
                    AllowUserToAddRow = vDGVAttachment.AllowUserToAddRows
                    vDGVAttachment.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVAttachment.Rows
                        DTAAttachment.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value, rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColAtt_Description.Index).Value)
                    Next
                    vDGVAttachment.AllowUserToAddRows = AllowUserToAddRow

                    'input work unit
                    DTAWorkUnit.Delete(Me.mPrimary)
                    AllowUserToAddRow = vDGVWorkUnit.AllowUserToAddRows
                    vDGVWorkUnit.AllowUserToAddRows = False
                    For Each rowDGV As DataGridViewRow In vDGVWorkUnit.Rows
                        DTAWorkUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, _
                             rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value)
                    Next
                    vDGVWorkUnit.AllowUserToAddRows = AllowUserToAddRow

                    Return True
                Else
                    Return False
                End If
            Else
                If IsNothing(pbx.Image) = True Then
                    'tidak ada data gambar yang perlu disimpan
                    If DTAUpdate.UpdateQuery2(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                    Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                    Me.mLocationControl, Me.mUseSecondaryUnits, Nothing, vUserCode, Now, Now, Me.mMinimumStock, _
                    Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                    Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, _
                    Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimaryAdjustmentAccount, Me.mPrimaryStockTakingAccount, _
                    Me.mPrimaryCOGSAccount, Me.mPrimarySalesAccount, Me.mPrimarySalesReturnAccount, Me.mPrimaryProcess, _
                    Me.mQtyRef, Me.mGoodPercentage, Me.mTypeGoods, Me.mTipeDependency, Me.mDelayDependency, Me.mTolerance, Me.mNettWeight, _
                    Me.mGrossWeight, Me.mPrimaryUnitWeight, Me.mPrecisionWeight, Me.mNettVolume, Me.mGrossVolume, Me.mPrimaryUnitVolume, _
                    Me.mPrecisionVolume, Me.mPrimaryTaxProductRef, Me.mPrimary, Me.mTimestamp) > 0 Then
                        'jika akun produknya ada yang berubah, maka update data verifikas menjadi tidak verified
                        If AkunBerbeda = True Then
                            DTA.UpdateQuery_SetAccountVerification(0, "", #1/1/1950#, Me.mPrimary)
                        End If

                        Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                        Me.mModifiedDate = Me.mTimestamp
                        Me.mModifiedBy = vUserCode

                        'hapus data lama
                        DTA1.Delete(Me.mPrimary)
                        'input data harga
                        For Each row In dgvPrice.Rows
                            DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                        Next

                        'hapus raw material
                        DTARM.Delete(Me.mPrimary)
                        'input raw material
                        AllowUserToAddRow = vDGVRawMaterial.AllowUserToAddRows
                        vDGVRawMaterial.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVRawMaterial.Rows
                            DTARM.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value, rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value, _
                            rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value, rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value, rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value)
                        Next
                        vDGVRawMaterial.AllowUserToAddRows = AllowUserToAddRow

                        'input multi unit
                        DTAUnit.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVMultiUnit.AllowUserToAddRows
                        vDGVMultiUnit.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVMultiUnit.Rows
                            DTAUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value, rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value, rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value) 'add weight & volume))
                        Next
                        vDGVMultiUnit.AllowUserToAddRows = AllowUserToAddRow

                        'custom value
                        Me.HapusItemKompetensi(Me.mPrimary)
                        Me.SimpanItemKompetensi(vDGVKompetensi, Me.mPrimary, vFrm)

                        'input attachment
                        DTAAttachment.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVAttachment.AllowUserToAddRows
                        vDGVAttachment.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVAttachment.Rows
                            DTAAttachment.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value, rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColAtt_Description.Index).Value)
                        Next
                        vDGVAttachment.AllowUserToAddRows = AllowUserToAddRow

                        'input work unit
                        DTAWorkUnit.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVWorkUnit.AllowUserToAddRows
                        vDGVWorkUnit.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVWorkUnit.Rows
                            DTAWorkUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, _
                            rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value)
                        Next
                        vDGVWorkUnit.AllowUserToAddRows = AllowUserToAddRow

                        Return True
                    Else
                        Return False
                    End If
                Else
                    'data gambar tidak mengalami perubahan
                    If DTAUpdate.UpdateQuery3(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                    Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                    Me.mLocationControl, Me.mUseSecondaryUnits, vUserCode, Now, Now, Me.mMinimumStock, _
                    Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                    Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, _
                    Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mTypeGoods, Me.mPrimaryTaxProductRef, Me.mPrimary, Me.mTimestamp) > 0 Then
                        Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                        Me.mModifiedDate = Me.mTimestamp
                        Me.mModifiedBy = vUserCode

                        'hapus data lama
                        DTA1.Delete(Me.mPrimary)
                        'input data harga
                        For i As Integer = 0 To dgvPrice.Rows.Count - 1
                            DTA1.Insert(Me.mPrimary, dgvPrice.Rows(i).Cells(1).Value, dgvPrice.Rows(i).Cells(3).Value, _
                            dgvPrice.Rows(i).Cells(4).Value)
                        Next

                        'hapus raw material
                        DTARM.Delete(Me.mPrimary)
                        'input raw material
                        AllowUserToAddRow = vDGVRawMaterial.AllowUserToAddRows
                        vDGVRawMaterial.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVRawMaterial.Rows
                            DTARM.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value, rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value, _
                            rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value, rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value, rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value)
                        Next
                        vDGVRawMaterial.AllowUserToAddRows = AllowUserToAddRow

                        'input multi unit
                        DTAUnit.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVMultiUnit.AllowUserToAddRows
                        vDGVMultiUnit.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVMultiUnit.Rows
                            DTAUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value, rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value, rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value, rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value, rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value, _
                            rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value) 'add weight & volume))
                        Next
                        vDGVMultiUnit.AllowUserToAddRows = AllowUserToAddRow

                        'custom value
                        Me.HapusItemKompetensi(Me.mPrimary)
                        Me.SimpanItemKompetensi(vDGVKompetensi, Me.mPrimary, vFrm)

                        'input attachment
                        DTAAttachment.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVAttachment.AllowUserToAddRows
                        vDGVAttachment.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVAttachment.Rows
                            DTAAttachment.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value, rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value, _
                        rowDGV.Cells(vFrm.ColAtt_Description.Index).Value)
                        Next
                        vDGVAttachment.AllowUserToAddRows = AllowUserToAddRow

                        'input work unit
                        DTAWorkUnit.Delete(Me.mPrimary)
                        AllowUserToAddRow = vDGVWorkUnit.AllowUserToAddRows
                        vDGVWorkUnit.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In vDGVWorkUnit.Rows
                            DTAWorkUnit.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value, _
                             rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value, rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value)
                        Next
                        vDGVWorkUnit.AllowUserToAddRows = AllowUserToAddRow

                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTARM.Dispose()
            DTAUnit.Dispose()
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            DTAAttachment.Dispose()
            DTAWorkUnit.Dispose()
            DTAUpdate.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String) As Boolean
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
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProduct, vUserCode)

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
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
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

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)    'data pricing otomatis di delete oleh server, data item material d delete otomatis

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

    Public Function ApakahSudahAdaBarcode(ByVal vBarcode As String) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByBarcode(vBarcode).HasValue Then
                If DTA.ScalarQuery_GetJmlByBarcode(vBarcode).Value > 0 Then
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

    Public Function ApakahGoods(ByVal vPrimaryProduct As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlGoods(vPrimaryProduct).HasValue Then
                If DTA.ScalarQuery_GetJmlGoods(vPrimaryProduct).Value > 0 Then
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

    Public Function ApakahSudahAda(ByVal vProductCode As String) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml2(vProductCode).HasValue Then
                If DTA.ScalarQuery_GetJml2(vProductCode).Value > 0 Then
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

    Public Function IsiListBatchNo(ByVal vPrimaryProduct As Double, ByVal LS As List(Of String), _
    ByVal LED As List(Of Date)) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
        Dim row As dseMasterProduct.tblmaster_productqty1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'isi database
            DTA.Fill(TBL, vPrimaryProduct)
            'kosongkan list
            LS.Clear()
            LED.Clear()
            'isi berdasarkan yang ada didalam basis data
            For Each row In TBL.Rows
                LS.Add(row.Lot_No)
                LED.Add(row.Expiry_Date)
            Next
            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahLocationControlled(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_IsLocationControlled(vPrimary).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsLocationControlled(vPrimary).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan nilai true jika sudah pernah digunakan pada PO, SO, Stock In Out, atau bundling
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahSudahPernahDigunakan(ByVal vPrimaryProduct As Double) As Boolean
        Try
            'cek, jika sudah ada penggunaan code no di:
            '1. po
            'Dim DTA As New dseMasterProductTableAdapters.tbltrans_po_itemTableAdapter
            'dta.Connection.ConnectionString = Core_Module.GetConstring
            'If DTA.ScalarQuery_GetJmlPOItemByProduct(vPrimaryProduct).Value > 0 Then
            ' DTA.Dispose()
            'Return True
            'End If
            'DTA.Dispose()
            '2. so
            'Dim DTA1 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            'dta1.Connection.ConnectionString = Core_Module.GetConstring
            'If DTA1.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
            'DTA1.Dispose()
            'Return True
            'End If
            'DTA1.Dispose()
            '3. stock in out
            Dim DTA2 As New dseMasterProductTableAdapters.tbltransaksi_stockinout_itemTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If DTA2.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA2.Dispose()
                Return True
            End If
            DTA2.Dispose()
            '4. stock bundling
            Dim DTA3 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
            DTA3.Connection.ConnectionString = Core_Module.GetConstring
            If DTA3.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA3.Dispose()
                Return True
            End If
            DTA3.Dispose()
            '5. product qty
            Dim DTA4 As New dseMasterProductTableAdapters.tblmaster_productqty5TableAdapter
            DTA4.Connection.ConnectionString = Core_Module.GetConstring
            If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                    DTA4.Dispose()
                    Return True
                End If
            End If
            DTA4.Dispose()

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahPernahClosing(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tbltrans_purchase_closing_item1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vPrimary).Value > 0 Then
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

    Public Function ApakahSudahPernahJurnal(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tbltrans_aop_item1TableAdapter
        Dim DTA1 As New DseMasterProduct1TableAdapters.tblmaster_productqtyTableAdapter
        Dim Jml_Nullable As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    Return True
                End If
            End If

            'cek ke item closing
            Jml_Nullable = DTA1.ScalarQuery_GetJmlClosing(vPrimary)
            If Jml_Nullable.HasValue = True Then
                If Jml_Nullable.Value > 0 Then
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function ApakahBatchNoControlled(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_IsBatchNoControlled(vPrimary).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsBatchNoControlled(vPrimary).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimaryProduct As Double, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim DTARM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bom1TableAdapter
        Dim TBLRM As New dseMasterProduct_BOM.tblmaster_product_item_bom1DataTable
        Dim rowRM As dseMasterProduct_BOM.tblmaster_product_item_bom1Row
        Dim Dummy As String = ""
        Dim Val As New Nullable(Of Long)
        Dim DtaProduct As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Try
            DTARM.Connection.ConnectionString = Core_Module.GetConstring

            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. Transaksi inventory
            'stock in out
            Dim DTA2 As New dseMasterProductTableAdapters.tbltransaksi_stockinout_itemTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If DTA2.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA2.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in stock in/out transaction data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA2.Dispose()
            'stock transfer
            Dim DTA5 As New dseMasterProductTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
            DTA5.Connection.ConnectionString = Core_Module.GetConstring
            If DTA5.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA5.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in stock transfer. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA5.Dispose()
            'stock bundling item in/out
            Dim DTA3 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
            DTA3.Connection.ConnectionString = Core_Module.GetConstring
            If DTA3.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA3.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in stock bundling. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA3.Dispose()
            Dim DTA16 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
            DTA16.Connection.ConnectionString = Core_Module.GetConstring
            If DTA16.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA16.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in stock bundling data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA16.Dispose()
            'stock tag item
            Dim DTA6 As New dseMasterProductTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
            DTA6.Connection.ConnectionString = Core_Module.GetConstring
            If DTA6.ScalarQuery_GEtJml(vPrimaryProduct).Value > 0 Then
                If vShowMsg = True Then
                    MsgBox("This product is already used in stock tag data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                DTA6.Dispose()
                Return False
            End If
            DTA6.Dispose()
            'product qty
            Dim DTA4 As New dseMasterProductTableAdapters.tblmaster_productqty5TableAdapter
            DTA4.Connection.ConnectionString = Core_Module.GetConstring
            If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                    DTA4.Dispose()
                    If vShowMsg = True Then
                        MsgBox("This product is already used in stock data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            DTA4.Dispose()
            '2. Transaksi sales
            'so
            Dim DTA7 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            DTA7.Connection.ConnectionString = Core_Module.GetConstring
            If DTA7.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA7.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in sales order data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA7.Dispose()
            'quotation
            Dim DTA8 As New dseMasterProductTableAdapters.tbltrans_quotation_itemTableAdapter
            DTA8.Connection.ConnectionString = Core_Module.GetConstring
            If DTA8.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA8.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in quotation. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA8.Dispose()
            'customer product code no
            Dim DTA11 As New dseMasterProductTableAdapters.tbltrans_customerproductcodestting_itemTableAdapter
            DTA11.Connection.ConnectionString = Core_Module.GetConstring
            If DTA11.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA11.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in customer code no data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA11.Dispose()
            '3. Transaksi puchase
            'historical purchase
            Dim DTA12 As New dseMasterProductTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
            DTA12.Connection.ConnectionString = Core_Module.GetConstring
            If DTA12.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA12.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in historical purchase data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA12.Dispose()
            'purchase request
            Dim DTA_PRItem As New dseMasterProductTableAdapters.tbltrans_purchase_req_itemTableAdapter
            DTA_PRItem.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_PRItem.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA_PRItem.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in purchase request data. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA_PRItem.Dispose()

            'po
            Dim DTA13 As New dseMasterProductTableAdapters.tbltrans_po_item1TableAdapter
            DTA13.Connection.ConnectionString = Core_Module.GetConstring
            If DTA13.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA13.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in purchase order. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA13.Dispose()

            'monthly purchase
            Dim DTA15 As New dseMasterProductTableAdapters.tbltrans_purchase_closing_itemTableAdapter
            DTA15.Connection.ConnectionString = Core_Module.GetConstring
            If DTA15.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA15.Dispose()
                If vShowMsg = True Then
                    MsgBox("This product is already used in monthly closing. You cannot delete this product", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            DTA15.Dispose()

            'sudah tidak ada raw material yang merujuk produk ini
            DTARM.FillBy_RM(TBLRM, vPrimaryProduct)
            If TBLRM.Rows.Count > 0 Then
                For Each rowRM In TBLRM.Rows
                    If Dummy = "" Then
                        Dummy = rowRM.Product_Code
                    Else
                        Dummy = Dummy + ", " + rowRM.Product_Code
                    End If

                    If vShowMsg = True Then
                        MsgBox("This product is raw material for the following product(s): " + _
                        Dummy + ". You cannot delete this product", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                Next
            End If

            DtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            ''a. item JO
            Val = DtaProduct.ScalarQuery_GetCountInMO_Item(vPrimaryProduct)
            If Val.HasValue Then
                If Val.Value > 0 Then
                    MsgBox("This product is already used in manufacturing order transaction, you cannot delete this product.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            ''b. item RM
            Val = DtaProduct.ScalarQuery_GetCountInMO_ItemRM(vPrimaryProduct)
            If Val.HasValue Then
                If Val.Value > 0 Then
                    MsgBox("This product is already used as raw material in manufacturing order transaction, you cannot delete this product.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'kalo sampe sini maka boleh hapus
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTARM.Dispose()
            TBLRM.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vDGVRM As DataGridView, ByVal vFRM As frmMasterProduct) As Boolean
        Dim Dummy As Boolean = False
        Dim ObjIdentifier As New ClsIdentifier
        Dim ObjABT As New ClsABT
        Dim Identifer As String = ""

        Try
            Dummy = vDGVRM.AllowUserToAddRows
            vDGVRM.AllowUserToAddRows = False
            For Each rowDGV As DataGridViewRow In vDGVRM.Rows
                'RM harus ada dalam dgv
                If Me.ApakahSudahAda(rowDGV.Cells(vFRM.ColBOM_CodeNo.Index).Value) = False Then
                    MsgBox("Invalid code no: " + rowDGV.Cells(vFRM.ColBOM_CodeNo.Index).Value + ".", _
                    MsgBoxStyle.Information, "Information")
                    vDGVRM.AllowUserToAddRows = Dummy
                    Return False
                End If
                'RM harus goods
                If Me.ApakahGoods(rowDGV.Cells(vFRM.ColBOM_PrimaryProduct.Index).Value) = False Then
                    MsgBox("Invalid code no: " + rowDGV.Cells(vFRM.ColBOM_CodeNo.Index).Value + ". Must be goods.", _
                    MsgBoxStyle.Information, "Information")
                    vDGVRM.AllowUserToAddRows = Dummy
                    Return False
                End If
            Next

            Identifer = ObjIdentifier.GetIdentifierDB
            If Identifer = "smess_admiral2" Then
                'kalo M dan Y, harus 2 decimal
                If Me.PrimaryUnit = 19 Or Me.PrimaryUnit = 25 Or Me.PrimaryUnit = 20 Then
                    If Me.AngkaDecimal <> 2 Then
                        MsgBox("Please specify: 2 for default precision data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                End If
                'kalo goods harus location control
                If Me.ProductType = EnumTipeProduct.Goods Then
                    If Me.LocationControl = False Then
                        MsgBox("Please specify location control for this data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                End If

                'kalo benang, harus batch no control
                If ObjABT.ApakahBenang(Me.mPrimaryCategory) = True Then
                    If Me.BatchNoControl = False Then
                        MsgBox("Please specify batch no control for this data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                    If Me.ExpiryDateControlled = True Then
                        MsgBox("Please specify expiry date control = 'false' for this data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                End If

                'kalo kain, harus ga batch no control
                If ObjABT.ApakahKain(Me.mPrimaryCategory) = True Then
                    If Me.BatchNoControl = True Then
                        MsgBox("Please specify batch no control = 'false' for this data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                    If Me.ExpiryDateControlled = True Then
                        MsgBox("Please specify expiry date control = 'false' for this data.", MsgBoxStyle.Information, "Informarion")
                        Return False
                    End If
                End If
            End If

            vDGVRM.AllowUserToAddRows = Dummy
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjIdentifier = Nothing
            ObjABT = Nothing
        End Try
    End Function

    Public Sub IsiCboLocation(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.stock_location_viewDataTable, ByVal vBatchNo As String)
        Dim DTA As New dseMasterProductTableAdapters.stock_location_viewTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()
            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSecondaryUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
        ByVal TBL As dseMasterProduct.tblmaster_productqty3DataTable, ByVal vBatchNo As String)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()

            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSecondaryUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty2DataTable, ByVal vBatchNo As String, _
    ByVal vPrimaryLocation As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()

            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo, vPrimaryLocation) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLUnitPrice(ByVal vPrimaryProduct As Double, ByVal vTBL As dseMasterProduct.DataTable3DataTable)
        Dim DTA As New dseMasterProductTableAdapters.DataTable3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryProduct(vTBL, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCBOMultiUnit(ByVal vCBO As ComboBox, ByVal vPrimaryProduct As Double, _
    ByVal DTAProduct As dseMasterProductTableAdapters.tblmaster_product5TableAdapter, _
    ByVal TBLProduct As dseMasterProduct.tblmaster_product5DataTable, ByVal DTAMultiUnit As dseMasterProductTableAdapters.tblmaster_product_item_multiunit2TableAdapter, _
    ByVal TBLMultiUnit As dseMasterProduct.tblmaster_product_item_multiunit2DataTable, Optional ByVal vDefaultPrimaryUnit As Double = -1)
        Dim rowProduct As dseMasterProduct.tblmaster_product5Row
        Dim rowMultiUnit As dseMasterProduct.tblmaster_product_item_multiunit2Row
        Dim DummyInt As Integer = 0

        Try
            TBLMultiUnit.Rows.Clear()
            TBLProduct.Rows.Clear()

            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAMultiUnit.Connection.ConnectionString = Core_Module.GetConstring

            vCBO.DataSource = Nothing
            vCBO.Items.Clear()

            'load item multi unit
            DTAMultiUnit.Fill(TBLMultiUnit, vPrimaryProduct)

            'tambahkan dengan primari unit default
            DTAProduct.Fill(TBLProduct, vPrimaryProduct)
            If TBLProduct.Rows.Count > 0 Then
                rowProduct = TBLProduct.Rows(0)
                rowMultiUnit = TBLMultiUnit.Newtblmaster_product_item_multiunit2Row
                rowMultiUnit.Primary = -1
                rowMultiUnit.Primary_Product = rowProduct.Primary
                rowMultiUnit.Primary_Unit = rowProduct.Primary_Unit
                rowMultiUnit.Unit_Code = rowProduct.Unit_Code
                rowMultiUnit.Qty_Default = 1
                rowMultiUnit.Qty_Rasio = 1
                rowMultiUnit.Default_Precision = rowProduct._Decimal

                TBLMultiUnit.Rows.Add(rowMultiUnit)
            End If

            'urutkan sesuai abjad
            TBLMultiUnit.DefaultView.Sort = "Unit Code"
            vCBO.DataSource = TBLMultiUnit.DefaultView
            vCBO.DisplayMember = "Unit Code"

            vCBO.SelectedIndex = -1

            If TBLMultiUnit.DefaultView.Count > 0 Then
                'pilih item pertama, jika primary unit default = -1
                If vDefaultPrimaryUnit = -1 Then
                    vCBO.SelectedIndex = 0
                Else
                    'pilih item sesuai primary unit multi yang diberikan
                    vCBO.SelectedIndex = 0
                    For Each DV As DataRowView In TBLMultiUnit.DefaultView
                        rowMultiUnit = DV.Row
                        If rowMultiUnit.Primary_Unit = vDefaultPrimaryUnit Then
                            vCBO.SelectedIndex = DummyInt
                            Exit For
                        Else
                            DummyInt += 1
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Public Sub IsiCboBatchNo(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vCbo.DataSource = Nothing
            vCbo.Refresh()
            vCbo.Items.Clear()
            If DTA.Fill(TBL, vPrimaryProduct) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiDGVBacthNo(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL, vPrimaryProduct) > 0 Then
                vDGV.DataSource = TBL.DefaultView
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByRef TBL As dseMasterProduct.tblmaster_product_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByRef DTA As dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.Fill(TBL)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DESCLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Product code"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    ''' <summary>
    ''' mengisikan dgv, tetapi hanya untuk item - item yang location controlled
    ''' </summary>
    ''' <param name="Criteria"></param>
    ''' <param name="DGV"></param>
    ''' <param name="TBL"></param>
    ''' <param name="SearchCriteria"></param>
    ''' <param name="tsl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVLocationControlledOnly(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterProduct.tblmaster_product_view6DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_view6TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_SpecLIKE_LocationControlled(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Product code"
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

    Public Function IsiDGVLocationControlledOnly(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As dseMasterProduct.tblmaster_product_view0DataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_SpecLIKE_LocationControlled(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Product code"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal vDgvKompetensi As DataGridView, ByVal Primary As Double, ByVal dgvPricing As DataGridView, ByVal pbx As PictureBox, _
    ByVal vDGV_BOM As DataGridView, ByVal vFrm As frmMasterProduct, ByVal vDGVMultiUnit As DataGridView, ByVal vDGVAttachment As DataGridView, _
    ByVal vDGVWorkUnit As DataGridView) As Boolean
        'Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        'Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        'Dim row As dseMasterProduct.tblmaster_product3Row
        Dim DTA As New dseMasterProduct2TableAdapters.tblcaridetailTableAdapter
        Dim TBL As New dseMasterProduct2.tblcaridetailDataTable
        Dim Row As dseMasterProduct2.tblcaridetailRow
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter
        Dim TBL1 As New dseMasterProduct.tblmaster_product_pricingDataTable
        Dim row1 As dseMasterProduct.tblmaster_product_pricingRow
        Dim DTA_BOM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bom1TableAdapter
        Dim TBL_BOM As New dseMasterProduct_BOM.tblmaster_product_item_bom1DataTable
        Dim row_BOM As dseMasterProduct_BOM.tblmaster_product_item_bom1Row
        Dim Dummy As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTAUnit As New dseMasterProductTableAdapters.tblmaster_product_item_multiunit1TableAdapter
        Dim TBLUnit As New dseMasterProduct.tblmaster_product_item_multiunit1DataTable
        Dim rowUnit As dseMasterProduct.tblmaster_product_item_multiunit1Row
        Dim DTAAttachment As New dseMasterProductTableAdapters.tblmaster_product_item_attachmentTableAdapter
        Dim TBLAttachment As New dseMasterProduct.tblmaster_product_item_attachmentDataTable
        Dim rowAttachment As dseMasterProduct.tblmaster_product_item_attachmentRow
        Dim DTAWorkUnit As New dseMasterProductTableAdapters.tblmaster_product_item_workunit1TableAdapter
        Dim TBLWorkUnit As New dseMasterProduct.tblmaster_product_item_workunit1DataTable
        Dim rowWorkUnit As dseMasterProduct.tblmaster_product_item_workunit1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_BOM.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWorkUnit.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mBatchNoControl = .Batch_No_Controlled
                    Me.mLocationControl = .Location_Controlled
                    Me.mExpiryDateControlled = .Expiry_Date_Controlled
                    Me.mCategoryCodeBounded = .Category_Code_Bounded
                    Me.mCategoryDescBounded = .Category_Desc_Bounded
                    Me.mPrimaryCategory = .Primary_Category
                    Me.mPrimaryMasterUnit = .Primary_Master_Unit
                    Me.mPrimarySecondaryUnit = .Primary_Secondary_Unit
                    Me.mPrimaryUnit = .Primary_Unit
                    Me.mProductCode = .Product_Code
                    Me.mQtyPerMasterUnit = .Qty_Per_Master
                    Me.mSpecification = .Specification
                    Me.mUseSecondaryUnits = .Use_Secondary_Unit
                    Me.mLeadTime = .Lead_Time
                    Me.mMinimumPurchaseQty = .Minimum_Purchase_Qty
                    Me.mMinimumStock = .Minimum_Stock
                    Me.mDesc = .Description
                    Me.mPdoNoControlled = .PdOControlled
                    Me.mDecimal = ._Decimal
                    Me.mTypeGoods = .type_product_goods

                    Me.mUseSizing = .Use_Sizing
                    Me.mSize = .Size
                    Me.mBarcode = .Barcode
                    Me.mTipe = .Type
                    Me.mPrimaryAkunInventory = .Primary_Inventory_Account
                    Me.mPrimaryAkunWIP = .Primary_WIP_Account

                    Me.mQtyRef = .Qty_Ref
                    Me.mGoodPercentage = .Good_Precentage
                    Me.mPrimaryProcess = .Primary_Proses

                    Me.mIsVerified = .is_account_verified
                    Me.mVerifiedBy = .verified_by
                    Me.mVerifiedTime = .verified_time

                    Me.mTipeDependency = .tipe_dependency
                    Me.mDelayDependency = .delay_dependency
                    Me.mTolerance = .production_tolerance

                    'add weight & volume
                    Me.mNettWeight = .Nett_Weight
                    Me.mGrossWeight = .Gross_Weight
                    Me.mPrimaryUnitWeight = .Primary_Unit_Weight
                    Me.mPrecisionWeight = .Default_Precision_Weight
                    Me.mNettVolume = .Nett_Volume
                    Me.mGrossVolume = .Gross_Volume
                    Me.mPrimaryUnitVolume = .Primary_Unit_Volume
                    Me.mPrecisionVolume = .Default_Precision_Volume

                    Me.mPrimaryTaxProductRef = Row.Primary_Tax_Product_Ref

                    'salin data gambar
                    If .IsPictureNull = False Then
                        Dim myByteArray() As Byte = .Picture

                        If myByteArray.Length > 0 Then
                            '' Create a new MemoryStream and write all the information from
                            '' the byte array into the stream
                            Dim myStream As MemoryStream = New MemoryStream(myByteArray, True)
                            myStream.Write(myByteArray, 0, myByteArray.Length)

                            '' Use the MemoryStream to create the new BitMap object
                            Dim FinalImage As Bitmap = New Bitmap(myStream)

                            pbx.Image = FinalImage
                            '' Close the stream
                            myStream.Close()
                        End If
                    Else
                        pbx.Image = Nothing
                    End If

                    'salin data pricing
                    dgvPricing.Rows.Clear()
                    DTA1.FillBy(TBL1, Me.mPrimary)
                    Dim Obj1 As New clsMasterCurrency
                    Dim Obj2 As New ClsCekData
                    For Each row1 In TBL1.Rows
                        With row1
                            dgvPricing.Rows.Add("-", .Primary_Currency, Obj1.GetCurrencyCode(.Primary_Currency), _
                            .Unit_Price, .Discount)
                        End With
                    Next
                    Obj2.SetNumber(dgvPricing)
                    Obj1 = Nothing
                    Obj2 = Nothing

                    'salin data bom
                    DTA_BOM.Fill(TBL_BOM, Primary)
                    vDGV_BOM.Rows.Clear()
                    Dummy = vDGV_BOM.AllowUserToAddRows
                    vDGV_BOM.AllowUserToAddRows = False
                    For Each row_BOM In TBL_BOM.Rows
                        vDGV_BOM.Rows.Add()
                        rowDGV = vDGV_BOM.Rows(vDGV_BOM.Rows.Count - 1)

                        With row_BOM
                            rowDGV.Cells(vFrm.ColBOM_CodeNo.Index).Value = .RM_Code
                            rowDGV.Cells(vFrm.ColBOM_Desc.Index).Value = .RM_Desc
                            rowDGV.Cells(vFrm.ColBOM_DrillDown.Index).Value = "Edit"
                            rowDGV.Cells(vFrm.ColBOM_GoodPercentage.Index).Value = ObjCD.DigitGrouping(.good_percentage, 2)
                            rowDGV.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value = .primary_product_raw
                            rowDGV.Cells(vFrm.ColBOM_Qty.Index).Value = ObjCD.DigitGrouping(.qty_raw, ._Decimal)
                            rowDGV.Cells(vFrm.ColBOM_Remarks.Index).Value = .remarks
                            rowDGV.Cells(vFrm.ColBOM_Unit.Index).Value = .Unit_Code
                            rowDGV.Cells(vFrm.ColBOM_Tolerance.Index).Value = ObjCD.DigitGrouping(.production_tolerance, 2)
                        End With
                    Next
                    vDGV_BOM.AllowUserToAddRows = Dummy

                    'salin data multi unit
                    DTAUnit.Fill(TBLUnit, Primary)
                    vDGVMultiUnit.Rows.Clear()
                    Dummy = vDGVMultiUnit.AllowUserToAddRows
                    vDGVMultiUnit.AllowUserToAddRows = False
                    For Each rowUnit In TBLUnit.Rows
                        vDGVMultiUnit.Rows.Add()
                        rowDGV = vDGVMultiUnit.Rows(vDGVMultiUnit.Rows.Count - 1)

                        With rowUnit
                            rowDGV.Cells(vFrm.ColUnit_Desc.Index).Value = .Description
                            rowDGV.Cells(vFrm.ColUnit_No.Index).Value = vDGVMultiUnit.RowCount
                            rowDGV.Cells(vFrm.ColUnit_Precision.Index).Value = .Default_Precision
                            rowDGV.Cells(vFrm.ColUnit_PrimaryUnit.Index).Value = .Primary_Unit
                            rowDGV.Cells(vFrm.ColUnit_QtyDef.Index).Value = ObjCD.DigitGrouping(.Qty_Default, Me.mDecimal)
                            rowDGV.Cells(vFrm.ColUnit_QtyRasio.Index).Value = ObjCD.DigitGrouping(.Qty_Rasio, .Default_Precision)
                            rowDGV.Cells(vFrm.ColUnit_Ratio.Index).Value = ObjCD.BuangNolDibelakangDecimal(.Rasio)
                            rowDGV.Cells(vFrm.ColUnit_Remarks.Index).Value = .Remarks
                            rowDGV.Cells(vFrm.ColUnit_Unit.Index).Value = .Unit_Code
                            rowDGV.Cells(vFrm.ColUnit_NettWeight.Index).Value = ObjCD.DigitGrouping(.Nett_Weight_Multi, .Default_Precision_Weight_Multi)
                            rowDGV.Cells(vFrm.ColUnit_GrossWeight.Index).Value = ObjCD.DigitGrouping(.Gross_Weight_Multi, .Default_Precision_Weight_Multi)
                            rowDGV.Cells(vFrm.ColUnit_UnitWeight.Index).Value = .Unit_Weight
                            rowDGV.Cells(vFrm.ColUnit_PrimaryUnitWeight.Index).Value = .Primary_Unit_Weight_Multi
                            rowDGV.Cells(vFrm.ColUnit_PrecisionWeight.Index).Value = .Default_Precision_Weight_Multi
                            rowDGV.Cells(vFrm.ColUnit_NettVolume.Index).Value = ObjCD.DigitGrouping(.Nett_Volume_Multi, .Default_Precision_Volume_Multi)
                            rowDGV.Cells(vFrm.ColUnit_GrossVolume.Index).Value = ObjCD.DigitGrouping(.Gross_Volume_Multi, .Default_Precision_Volume_Multi)
                            rowDGV.Cells(vFrm.ColUnit_UnitVolume.Index).Value = .Unit_Volume
                            rowDGV.Cells(vFrm.ColUnit_PrimaryUnitVolume.Index).Value = .Primary_Unit_Volume_Multi
                            rowDGV.Cells(vFrm.ColUnit_PrecisionVolume.Index).Value = .Default_Precision_Volume_Multi
                        End With
                    Next
                    vDGVMultiUnit.AllowUserToAddRows = Dummy

                    'salin data attachment
                    DTAAttachment.FillBy_PrimaryProduct(TBLAttachment, Primary)
                    vDGVAttachment.Rows.Clear()
                    Dummy = vDGVAttachment.AllowUserToAddRows
                    vDGVAttachment.AllowUserToAddRows = False
                    For Each rowAttachment In TBLAttachment.Rows
                        vDGVAttachment.Rows.Add()
                        rowDGV = vDGVAttachment.Rows(vDGVAttachment.Rows.Count - 1)

                        With rowAttachment
                            rowDGV.Cells(vFrm.ColAtt_No.Index).Value = rowDGV.Index + 1
                            rowDGV.Cells(vFrm.ColAtt_BtnOpenDelete.Index).Value = "remove"
                            rowDGV.Cells(vFrm.ColAtt_Description.Index).Value = .desc
                            rowDGV.Cells(vFrm.ColAtt_LocationFile.Index).Value = .file_path
                            rowDGV.Cells(vFrm.ColAtt_Preview.Index).Value = "preview"
                            rowDGV.Cells(vFrm.ColAtt_Remarks.Index).Value = .remarks
                        End With
                    Next
                    vDGVAttachment.AllowUserToAddRows = Dummy

                    'salin data work unit
                    DTAWorkUnit.Connection.ConnectionString = DTA.Connection.ConnectionString
                    DTAWorkUnit.FillBy_PrimaryProduct(TBLWorkUnit, Primary)
                    vDGVWorkUnit.Rows.Clear()
                    Dummy = vDGVWorkUnit.AllowUserToAddRows
                    vDGVWorkUnit.AllowUserToAddRows = False
                    For Each rowWorkUnit In TBLWorkUnit.Rows
                        vDGVWorkUnit.Rows.Add()
                        rowDGV = vDGVWorkUnit.Rows(vDGVWorkUnit.Rows.Count - 1)
                        With rowWorkUnit
                            rowDGV.Cells(vFrm.ColWorkUnit_No.Index).Value = rowDGV.Index + 1
                            rowDGV.Cells(vFrm.ColWorkUnit_Code.Index).Value = .work_unit_code
                            rowDGV.Cells(vFrm.ColWorkUnit_Desc.Index).Value = .description
                            rowDGV.Cells(vFrm.ColWorkUnit_Efficiency.Index).Value = ObjCD.DigitGrouping(.efficiency, 2)
                            rowDGV.Cells(vFrm.ColWorkUnit_Rate.Index).Value = ObjCD.DigitGrouping(.rate_per_hour, Me.mDecimal)
                            rowDGV.Cells(vFrm.ColWorkUnit_Remarks.Index).Value = .remarks
                            rowDGV.Cells(vFrm.ColWorkUnit_PrimaryWorkUnit.Index).Value = .primary_work_unit
                        End With
                    Next
                    vDGVWorkUnit.AllowUserToAddRows = Dummy

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mPrimaryStockTakingAccount = .Primary_Stock_Taking_Account
                    Me.mPrimarySalesReturnAccount = .Primary_Sales_Return_Account
                    Me.mPrimarySalesAccount = .Primary_Sales_Account
                    Me.mPrimaryAdjustmentAccount = .Primary_Adjustment_Account
                    Me.mPrimaryCOGSAccount = .Primary_COGS_Account

                    Me.IsiDGVItemKompetensi(vDgvKompetensi, .Primary, vFrm)
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA_BOM.Dispose()
            TBL_BOM.Dispose()
            ObjCD = Nothing
            DTAUnit.Dispose()
            TBLUnit.Dispose()
            DTAAttachment.Dispose()
            TBLAttachment.Dispose()
            DTAWorkUnit.Dispose()
            TBLWorkUnit.Dispose()
        End Try
    End Function

    Public Function GetProductCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

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

    Public Function GetSecondaryUnitCode(ByVal vPrimaryProduct As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSecondaryUnitsCode(vPrimaryProduct)) = False Then
                Return DTA.ScalarQuery_GetSecondaryUnitsCode(vPrimaryProduct)
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

    Public Function GetUnitCode(ByVal vPrimaryProduct As String) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetUnitCode(vPrimaryProduct)) = False Then
                Return DTA.ScalarQuery_GetUnitCode(vPrimaryProduct)
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

    Public Function GetStockAwal(ByVal vPrimary As Double) As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetStockAwal(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetStockAwal(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan jumlah product yang "bounded by category" untuk category tertentu
    ''' </summary>
    ''' <param name="vPrimaryCategory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml_ProductBoundedByCategory(ByVal vPrimaryCategory As Double) As Long
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlBounded(vPrimaryCategory).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlBounded(vPrimaryCategory).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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

    Public Function ApakahUseSecondary(ByVal vPrimaryProduct As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_IsUseSecondary(vPrimaryProduct).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsUseSecondary(vPrimaryProduct).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAngkaDecimal(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDecimal(vPrimaryUnit, vPrimaryProduct).HasValue Then
                Return DTA.ScalarQuery_GetDecimal(vPrimaryUnit, vPrimaryProduct).Value
            Else
                Return Me.GetAngkaDecimal(vPrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAngkaDecimal(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDecimal(vPrimaryProduct)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryCategory As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByCategory(vPrimaryCategory).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCategory(vPrimaryCategory).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vProductCode As String) As Long
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml2(vProductCode).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vProductCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetExpiryDate(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Date
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).HasValue = True Then
                Return DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimary(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByUnit(ByVal vPrimaryUnit As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml_ByPrimaryMasterUnit(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If
            Retval = DTA.ScalarQuery_GetJmlByPrimarySecondaryUnit(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If
            Retval = DTA.ScalarQuery_GetJmlByPrimaryUnit(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If

            'add weight & volume
            Retval = DTA.ScalarQuery_GetJmlByUnitWeight(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If
            Retval = DTA.ScalarQuery_GetJmlByUnitVolume(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If
            Retval = DTA.ScalarQuery_GetJmlByUnitWeightMulti(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If
            Retval = DTA.ScalarQuery_GetJmlByUnitVolumeMulti(vPrimaryUnit)
            If Retval.HasValue Then
                If Retval.Value > 0 Then
                    Return Retval
                End If
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByStorageArea(ByVal vPrimaryStorageArea As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryStorageArea(vPrimaryStorageArea).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryStorageArea(vPrimaryStorageArea).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_PrimaryCurrency(ByVal vPrimaryCurrency As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryCurrency).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSpecification(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getSpecification(vPrimary)) = False Then
                Return DTA.ScalarQuery_getSpecification(vPrimary)
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

    Public Function GetDiscount(ByVal vPrimaryCurrency As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDiscount(vPrimaryProduct, vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetDiscount(vPrimaryProduct, vPrimaryCurrency).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLeadTime(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product1TableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetLeadTime(vPrimaryProduct)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetUnitPriceNDiskon(ByVal vPrimaryCustomer As Double, ByVal vPrimaryCodeNo As Double, ByVal vQty As Decimal, ByVal vPrimaryCurrency As Double, _
    ByVal vPrimaryUnit As Double, ByRef vUnitPrice As Decimal, ByRef vDiskon As Decimal)
        Dim RowHasil As DseMasterProduct_EditPriceList_GetPrice.TblHasilRow
        Dim DTACustomer As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Dim TBLCustomer As New DseMasterProduct_EditPriceList_GetPrice.tblmaster_product_pricelist_customerDataTable
        Dim RowCustomer As DseMasterProduct_EditPriceList_GetPrice.tblmaster_product_pricelist_customerRow
        Dim TipeGroup As frmMasterProduct_PriceList.EnumCustomerGroupTipe
        Dim PrimaryGroup As Double = -1

        Dim DTAGroup As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim TBLGroup As New DseMasterProduct_EditPriceList_GetPrice.tblmaster_product_pricelist_groupDataTable
        Dim RowGroup As DseMasterProduct_EditPriceList_GetPrice.tblmaster_product_pricelist_groupRow

        Dim DTAGeneral As New DseMasterProduct_EditPriceList_GetPriceTableAdapters.DataTable3TableAdapter
        Dim TBLGeneral As New DseMasterProduct_EditPriceList_GetPrice.DataTable3DataTable
        Dim RowGeneral As DseMasterProduct_EditPriceList_GetPrice.DataTable3Row
        Dim TblHasil As New DseMasterProduct_EditPriceList_GetPrice.TblHasilDataTable
        Dim Obj As New ClsMasterProduct_PriceList_GetPrice
        Dim Ketemu As Boolean
        Dim ObjCur As New clsMasterCurrency

        Try
            TblHasil.Rows.Clear()
            DTACustomer.Connection.ConnectionString = Core_Module.GetConstring
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAGeneral.Connection.ConnectionString = DTACustomer.Connection.ConnectionString

            DTACustomer.FillBy_Customer(TBLCustomer, CDec(vPrimaryCodeNo), vPrimaryCurrency, vPrimaryCustomer, vPrimaryUnit)
            If TBLCustomer.Rows.Count > 0 Then
                For Each RowCustomer In TBLCustomer.Rows
                    RowHasil = TblHasil.NewRow
                    RowHasil.Quantitiy = RowCustomer.max_qty
                    RowHasil.Unit_Price = RowCustomer.unit_price
                    RowHasil.Discount = RowCustomer.discount
                    TblHasil.Rows.Add(RowHasil)
                Next
            Else
                TipeGroup = Obj.GetTipeGroup
                If TipeGroup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                    PrimaryGroup = Obj.GetCustCustomerGroup(vPrimaryCustomer)
                ElseIf TipeGroup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                    PrimaryGroup = Obj.GetCustSalesArea(vPrimaryCustomer)
                ElseIf TipeGroup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                    PrimaryGroup = Obj.GetCustSegmentation(vPrimaryCustomer)
                ElseIf TipeGroup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                    PrimaryGroup = Obj.GetCustType(vPrimaryCustomer)
                End If
                DTAGroup.FillBy_Group(TBLGroup, vPrimaryCodeNo, vPrimaryCurrency, PrimaryGroup, vPrimaryUnit)
                If TBLGroup.Rows.Count > 0 Then
                    For Each RowGroup In TBLGroup.Rows
                        RowHasil = TblHasil.NewRow
                        RowHasil.Quantitiy = RowGroup.max_qty
                        RowHasil.Unit_Price = RowGroup.unit_price
                        RowHasil.Discount = RowGroup.discount
                        TblHasil.Rows.Add(RowHasil)
                    Next
                Else
                    DTAGeneral.FillBy_General(TBLGeneral, vPrimaryCurrency, vPrimaryCodeNo, vPrimaryUnit)
                    If TBLGeneral.Rows.Count > 0 Then
                        For Each RowGeneral In TBLGeneral.Rows
                            RowHasil = TblHasil.NewRow
                            RowHasil.Quantitiy = RowGeneral.max_qty
                            RowHasil.Unit_Price = RowGeneral.unit_price
                            RowHasil.Discount = RowGeneral.discount
                            TblHasil.Rows.Add(RowHasil)
                        Next
                    End If
                End If
            End If

            If TblHasil.Rows.Count > 0 Then
                For idx As Integer = 1 To TblHasil.Rows.Count - 1
                    RowHasil = TblHasil.Rows(idx)
                    If vQty <= RowHasil.Quantitiy Then
                        Ketemu = True
                        vUnitPrice = Math.Round(RowHasil.Unit_Price, ObjCur.GetDefaultPrecission(vPrimaryCurrency), MidpointRounding.AwayFromZero)
                        vDiskon = Math.Round(RowHasil.Discount, 2, MidpointRounding.AwayFromZero)
                        Exit Sub
                    End If
                Next
                If Ketemu = False Then
                    RowHasil = TblHasil.Rows(0)
                    vUnitPrice = Math.Round(RowHasil.Unit_Price, ObjCur.GetDefaultPrecission(vPrimaryCurrency), MidpointRounding.AwayFromZero)
                    vDiskon = Math.Round(RowHasil.Discount, 2, MidpointRounding.AwayFromZero)
                    Exit Sub
                End If
            End If

            vUnitPrice = 0
            vDiskon = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            vUnitPrice = 0
            vDiskon = 0
        Finally
            TblHasil.Dispose()
            Obj = Nothing
            DTACustomer.Dispose()
            TBLCustomer.Dispose()
            DTAGroup.Dispose()
            TBLGroup.Dispose()
            DTAGeneral.Dispose()
            TBLGeneral.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Public Function SimpanQtyBaru(ByVal vJmlBaru As Decimal, ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Decimal = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "", _
    Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA1.Insert(vPrimaryProduct, vJmlBaru, vSecondaryQty, vBatchNo, vExpiryDate, vPrimaryLocation) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Decimal
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA1.ScalarQuery_GetJml(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty).HasValue = True Then
                Return DTA1.ScalarQuery_GetJml(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml3(ByVal vPrimaryProduct As Double, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return Me.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
            vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml4(ByVal vPrimaryProduct As Double, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vBatchNo As String = "") As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml2(vPrimaryProduct, vBatchNo, vSecondaryQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml5(ByVal vPrimaryCodeNo As Double, ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml3(vPrimaryCodeNo, vSecQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ProductNotVerified() As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product4TableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJmlNotVerified
            If Jml.HasValue = True Then
                Return Jml.Value
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

    '''mengembalikan yes, jika ada product yg di transaksi pd bulan ybs yang blm diverified
    Public Function ApakahAda_ProductNotVerified(ByVal vRefPeriod As Date) As Boolean
        Dim DTA As New DseMasterProduct1TableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseMasterProduct1.tblmaster_productDataTable
        Dim row As DseMasterProduct1.tblmaster_productRow
        Dim StartDate, EndDate As Date
        Dim ObjCD As New ClsCekData
        Dim DTA1 As New DseMasterProduct1TableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL1 As New DseMasterProduct1.tbltransaksi_stockinout_itemDataTable
        Dim row1 As DseMasterProduct1.tbltransaksi_stockinout_itemRow
        Dim DTA2 As New DseMasterProduct1TableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
        Dim TBL2 As New DseMasterProduct1.tbltrans_stock_taking_tag_itemDataTable
        Dim row2 As DseMasterProduct1.tbltrans_stock_taking_tag_itemRow
        Dim IdxCheck As Integer = 0
        Dim Jml1, Jml2 As Double

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            EndDate = ObjCD.GetLasDateOfTheMonth(vRefPeriod)
            StartDate = ObjCD.GetFirstDateofTheMonth(vRefPeriod)

            DTA.Fill(TBL)
            DTA1.Fill(TBL1, EndDate, StartDate)
            DTA2.Fill(TBL2, EndDate, StartDate)

            While IdxCheck <= TBL.Rows.Count - 1
                row = TBL.Rows(IdxCheck)
                Jml1 = 0
                Jml2 = 0

                If TBL1.Rows.Count > 0 Then
                    row1 = TBL1.Rows(0)
                    While row1.Primary_Product <= row.Primary
                        If row1.Primary_Product = row.Primary Then
                            Jml1 = row1.Jml
                            Exit While
                        ElseIf row1.Primary_Product < row.Primary Then
                            TBL1.Rows.Remove(row1)
                            If TBL1.Rows.Count > 0 Then
                                row1 = TBL1.Rows(0)
                            Else
                                Jml1 = 0
                                Exit While
                            End If
                        Else
                            Exit While
                        End If
                    End While
                Else
                    Jml1 = 0
                End If

                If TBL2.Rows.Count > 0 Then
                    row2 = TBL2.Rows(0)
                    While row2.Primary_Product <= row.Primary
                        If row2.Primary_Product = row.Primary Then
                            Jml2 = row2.JmlItem
                            Exit While
                        ElseIf row2.Primary_Product < row.Primary Then
                            TBL2.Rows.Remove(row2)
                            If TBL2.Rows.Count > 0 Then
                                row2 = TBL2.Rows(0)
                            Else
                                Jml2 = 0
                                Exit While
                            End If
                        Else
                            Exit While
                        End If
                    End While
                Else
                    Jml2 = 0
                End If

                If Jml1 > 0 Or Jml2 > 0 Then
                    Return True
                Else
                    TBL.Rows.Remove(row)
                End If
            End While

            If TBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    ''''mengembalikan jumlah product yang belum diverifikasi jika ada transaksi pada bulan ybs
    'Public Function GetJml_ProductNotVerified(ByVal vRefPeriod As Date) As Decimal
    '    Dim DTA As New dseMasterProductTableAdapters.tblmaster_product6TableAdapter
    '    Dim TBL As New dseMasterProduct.tblmaster_product6DataTable
    '    Dim row As dseMasterProduct.tblmaster_product6Row
    '    Dim EndDate As Date
    '    Dim ObjCD As New ClsCekData
    '    Dim DTA1 As New dseMasterProductTableAdapters.tbltransaksi_stockinout_itemTableAdapter
    '    Dim TBL1 As New dseMasterProduct.tbltransaksi_stockinout_itemDataTable
    '    Dim DTA2 As New dseMasterProductTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
    '    Dim TBL2 As New dseMasterProduct.tbltrans_stock_taking_tag_itemDataTable

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

    '        EndDate = ObjCD.GetLasDateOfTheMonth(vRefPeriod)
    '        DTA.Fill(TBL)
    '        If TBL.Rows.Count > 0 Then
    '            For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
    '                row = TBL.Rows(ij)
    '                DTA1.FillBy_PrimaryProduct(TBL1, EndDate, row.Primary)
    '                DTA2.FillBy_PrimaryProduct(TBL2, EndDate, row.Primary)
    '                If TBL1.Rows.Count = 0 And TBL2.Rows.Count = 0 Then
    '                    'belum pernah transaksi stock in out dan stock taking sama sekali sebelum end date, tidak dihitung datanya
    '                    TBL.Rows.RemoveAt(ij)
    '                End If
    '            Next

    '            Return TBL.Rows.Count
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '        ObjCD = Nothing
    '    End Try
    'End Function

    ''' <summary>
    ''' mengembalikan jumlah product yang location controlled
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml_LocationControlled() As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlLocationControlled.HasValue = True Then
                Return DTA.ScalarQuery_GetJmlLocationControlled.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub SetBatchNumberControlled(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetBatchNumberControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetExpiryDateControlled(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetEDControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetLocationControl(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetLocationControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub DeleteProductQty(ByVal vPrimaryProduct As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub Isikan_TBLRM(ByVal vPrimaryProduct As Double, ByVal vTBLRM As dseMasterProduct_BOM.tblmaster_product_item_bom1DataTable)
        Dim DTA As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bom1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBLRM, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub InsertProductQty(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, ByVal vColBatch As Integer, _
    ByVal vColED As Integer, ByVal vColQty As Integer, ByVal vColPrimaryLocation As Integer, ByVal vQtyNoControl As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim row As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vDGV.RowCount = 0 Then
                If vQtyNoControl > 0 Then
                    DTA.Insert(vPrimaryProduct, vQtyNoControl, -1, _
                    "", #1/1/1950#, -1)
                End If
            Else
                'secondary qty =-1
                For Each row In vDGV.Rows
                    With row
                        DTA.Insert(vPrimaryProduct, .Cells(vColQty).Value, -1, _
                        .Cells(vColBatch).Value, .Cells(vColED).Value, .Cells(vColPrimaryLocation).Value)
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetTanggalMinimum(ByVal vPrimaryProduct As Double, ByVal vTanggalMinimum As Date)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetTanggalMinimum(vTanggalMinimum, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function SetJml(ByVal vJmlBaru As Decimal, ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Boolean
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            If DTA1.UpdateQuery_SetJml(vJmlBaru, vBatchNo, vPrimaryLocation, vPrimaryProduct, vSecondaryQty) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "") As Boolean
        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA1.Delete(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryUnit(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryUnit(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUnit(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCategory(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCategory(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCategory(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vProductCode As String, ByVal vPrimaryCategory As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vProductCode, vPrimaryCategory).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vProductCode, vPrimaryCategory).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryBrand(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryBrand(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryBrand(vPrimaryProduct).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_ByDesc(ByVal vProductDesc As String) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByDesc(vProductDesc).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByDesc(vProductDesc).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vProductCode As String) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary2(vProductCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary2(vProductCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahExpiryDateControlled(ByVal vPrimaryProduct As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_IsiExpiryDateControlled(vPrimaryProduct).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsiExpiryDateControlled(vPrimaryProduct).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub IsiCboBatchNo_UntukUpdateBatch(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
 ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter
        Dim DTAIn As New dseMasterProductTableAdapters.DataTable5TableAdapter
        Dim TBLIn As New dseMasterProduct.DataTable5DataTable
        Dim rowIn As dseMasterProduct.DataTable5Row
        Dim row As dseMasterProduct.tblmaster_productqty1Row
        Dim i As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAIn.Connection.ConnectionString = DTA.Connection.ConnectionString

            vCbo.DataSource = Nothing
            vCbo.Refresh()
            vCbo.Items.Clear()
            DTA.Fill(TBL, vPrimaryProduct)

            'tambahkan dengan semua stock in yang belum ada dalam tabel diatas
            DTAIn.Fill(TBLIn, vPrimaryProduct)
            For Each row In TBL.Rows
                'buang baris2 yang ada di tblIn, dimana batch numbernya sudah ada di dalam TBL
                i = 0
                While i < TBLIn.Rows.Count
                    rowIn = TBLIn.Rows(i)
                    If rowIn.Batch_No.ToLower = row.Lot_No.ToLower Then
                        TBLIn.Rows.Remove(rowIn)
                    Else
                        i += 1
                    End If
                End While
            Next
            For Each rowIn In TBLIn.Rows
                row = TBL.Newtblmaster_productqty1Row
                row.Lot_No = rowIn.Batch_No
                row.Expiry_Date = rowIn.Expiry_Date
                row.Qty = 0
                TBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAIn.Dispose()
            TBLIn.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryByBarcode(ByVal vBarcode As String) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByBarcode(vBarcode).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByBarcode(vBarcode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBarcode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetBarcode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetBarcode(vPrimary)
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

    Public Sub UpdateAccount_AdjustmentAccount(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetAdjustmentAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAccount_StockTakingAccount(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetStockTakingAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAccount_Inventory(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetInventoryAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAccount_Sales(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetSalesAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAccount_SalesReturn(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetSalesReturnAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAccount_COGS(ByVal vPrimaryAkun As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetCOGSAccount(vPrimaryAkun)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub New()
        Me.DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring
    End Sub

    Public Sub SimpanItemKompetensi(ByVal vDGV As DataGridView, ByVal vPrimaryProduct As Double, ByVal frm As frmMasterProduct)
        Dim dta As New dseMasterProductTableAdapters.tblmaster_product_item_kompetensiTableAdapter
        Dim row As DataGridViewRow
        Dim tgl As Date
        Dim i As Integer = 1
        Dim remark As String = ""
        Dim Value_ As String = ""
        Dim DataType As Integer = -1
        Dim ObjKompetensi As New ClsMasterKompetensi_Product
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vDGV.Rows
                If i <> vDGV.Rows.Count Then
                    If row.Cells(frm.ColCompetence_NextCertificationDate.Index).Value = "-" Then
                        tgl = #1/1/1950#
                    Else
                        tgl = row.Cells(frm.ColCompetence_NextCertificationDate.Index).Value
                    End If
                    If String.IsNullOrEmpty(row.Cells(frm.ColCompetence_Remarks.Index).Value) = True Then
                        remark = ""
                    Else
                        remark = row.Cells(frm.ColCompetence_Remarks.Index).Value
                    End If
                    'aul 12/des/2015
                    DataType = ObjKompetensi.GetDataType(row.Cells(frm.ColCompetence_Primary.Index).Value)
                    If DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList Then
                        Value_ = row.Cells(frm.ColCempetence_Identifier.Index).Value
                    Else
                        Value_ = row.Cells(frm.ColCompetence_Value.Index).Value
                    End If
                    dta.InsertQuery(vPrimaryProduct, row.Cells(frm.ColCompetence_Primary.Index).Value, _
                    Value_, remark, _
                    tgl)

                    i = i + 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Function HapusItemKompetensi(ByVal vPrimaryProduct As Double) As Boolean
        Dim dtaKompetensi As New dseMasterProductTableAdapters.tblmaster_product_item_kompetensiTableAdapter
        Try
            dtaKompetensi.Connection.ConnectionString = Core_Module.GetConstring
            dtaKompetensi.DeleteQuery(vPrimaryProduct)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaKompetensi.Dispose()
        End Try
    End Function

    Public Function IsiDGVItemKompetensi(ByVal vDGVItem As DataGridView, ByVal vPrimaryProduct As Double, _
    ByVal Frm As frmMasterProduct) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_item_kompetensiTableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product_item_kompetensiDataTable
        Dim row As dseMasterProduct.tblmaster_product_item_kompetensiRow
        Dim rowDGV As DataGridViewRow
        Dim ObjKompetensi As New ClsMasterKompetensi_Product
        Dim ObjCekData As New ClsCekData
        Dim DataType As Integer = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVItem.AllowUserToAddRows = False

            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    With row
                        DataType = ObjKompetensi.GetDataType(.primary_kompetensi)
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColCompetence_Code.Index).Value = ObjKompetensi.GetDataCode(.primary_kompetensi)
                        rowDGV.Cells(Frm.ColCempetence_Description.Index).Value = ObjKompetensi.GetDesc(.primary_kompetensi)
                        Console.WriteLine(vDGVItem.Item(1, 0).Value)
                        'If ObjJob.GetJmlItem(vPrimaryJob, .primary_kompetensi) > 0 Then
                        '    rowDGV.Cells(Frm.ColCompetence_StandarValue.Index).Value = _
                        '    ObjJob.GetValue_CustomFormatted(vPrimaryJob, .primary_kompetensi, ObjCekData, DataType)
                        'Else
                        rowDGV.Cells(Frm.ColCompetence_StandarValue.Index).Value = _
                        ObjKompetensi.GetValue_CustomFormatted(.primary_kompetensi, ObjCekData, DataType)
                        'End If
                        If DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList Then
                            rowDGV.Cells(Frm.ColCempetence_Identifier.Index).Value = .value
                            rowDGV.Cells(Frm.ColCompetence_Value.Index).Value = ObjKompetensi.GetDataCodeByIdentifier(rowDGV.Cells(Frm.ColCempetence_Identifier.Index).Value)
                        Else
                            rowDGV.Cells(Frm.ColCompetence_Value.Index).Value = .value
                        End If
                        'My.Application.ChangeCulture("id-ID")
                        'My.Application.ChangeUICulture("id-ID")
                        If .next_certification_date = "1/1/1950" Then
                            rowDGV.Cells(Frm.ColCompetence_NextCertificationDate.Index).Value = "-"
                        Else
                            rowDGV.Cells(Frm.ColCompetence_NextCertificationDate.Index).Value = Format(.next_certification_date, "dd MMM yyyy")
                        End If
                        'My.Application.ChangeCulture("en-US")
                        'My.Application.ChangeUICulture("en-US")
                        rowDGV.Cells(Frm.ColCompetence_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColCompetence_Primary.Index).Value = .primary_kompetensi
                        rowDGV.Cells(Frm.ColCempetence_DataType.Index).Value = DataType
                    End With
                Next
            End If

            If Frm.btnOK.Enabled = True Then
                vDGVItem.AllowUserToAddRows = True
            End If
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCekData = Nothing
            ObjKompetensi = Nothing
        End Try
    End Function

    Public Function ApakahSudahVerifikasi(ByVal vListPrimaryProduct As List(Of Double), ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterProductAccountGroupTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseMasterProductAccountGroup.tblmaster_productDataTable
        Dim TBL_Result As New DseMasterProductAccountGroup.tblmaster_productDataTable
        Dim row, row_Result As DseMasterProductAccountGroup.tblmaster_productRow
        Dim PrimaryProduct As Double = -1
        Dim Frm1 As New frmMasterProduct_Verifikasi_Cek
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting

        Try
            If ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_VerifyProductAccount) = "1" Then
                TBL.Rows.Clear()
                TBL_Result.Rows.Clear()

                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.Fill(TBL)

                For Each PrimaryProduct In vListPrimaryProduct
                    TBL.DefaultView.RowFilter = "Primary = " + PrimaryProduct.ToString
                    If TBL.DefaultView.Count > 0 Then
                        row = TBL.DefaultView(0).Row
                        'tidak verified, masukkan ke tabel hasil
                        row_Result = TBL_Result.Newtblmaster_productRow
                        For ij As Integer = 0 To TBL.Columns.Count - 1
                            row_Result.Item(ij) = row.Item(ij)
                        Next
                        TBL_Result.Rows.Add(row_Result)
                    End If
                Next

                If TBL_Result.Rows.Count > 0 Then
                    If vShowMsg = True Then
                        Frm1.TBL_Result = TBL_Result
                        Frm1.ShowDialog()
                    End If
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Frm1 = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    'penambahan untuk perbaikan master product
    Public Function ApakahBolehEditDecimalProduct(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, _
    ByRef vDecimal As Double, Optional ByVal vApakahMulti As Boolean = False) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tbltrans_po_itemTableAdapter
        Dim DTA1 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
        Dim DTA2 As New dseMasterProductTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim DTA3 As New dseMasterProductTableAdapters.tbltrans_quotation_itemTableAdapter
        Dim Retval As Decimal = 0
        Dim DecimalQty As Integer = 0
        Dim RetVal1 As Nullable(Of Decimal)

        Try
            'cek, jika sudah ada penggunaan code no di:
            '1. po
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval = DTA.ScalarQuery_GetTotalQtyProcuctInPO(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval = DTA.ScalarQuery_GetTotalQtyMultiProductInPO(vPrimaryProduct, vPrimaryUnit).Value
            End If

            '2. so
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval += DTA1.ScalarQuery_GetTotalQtyProductInSO(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval += DTA1.ScalarQuery_GetTotalQtyMultiProductInSO(vPrimaryProduct, vPrimaryUnit).Value
            End If

            '3. Purchase request
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            Retval += DTA2.ScalarQuery_GetTotalQtyProductInPReq(vPrimaryProduct, vPrimaryUnit).Value

            '4. Quotation
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval += DTA3.ScalarQuery_GetTotalQtyProductInQuo(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval += DTA3.ScalarQuery_GetTotalQtyMultiProductInQuo(vPrimaryProduct, vPrimaryUnit).Value
            End If

            'stock in out
            If vApakahMulti = False Then
                RetVal1 = DTA3.ScalarQuery_GetJmlByStockInOut(vPrimaryProduct, vPrimaryUnit)
                If RetVal1.HasValue = True Then
                    Retval += RetVal1.Value
                End If
            Else
                RetVal1 = DTA3.ScalarQuery_GetJmlMultiStockInOut(vPrimaryProduct, vPrimaryUnit)
                If RetVal1.HasValue = True Then
                    Retval += RetVal1.Value
                End If
            End If

            'stock bundling
            If vApakahMulti = False Then
                RetVal1 = DTA3.ScalarQuery_GetJmlItemAfterBundling(vPrimaryProduct, vPrimaryUnit)
                If RetVal1.HasValue = True Then
                    Retval += RetVal1.Value
                End If

                RetVal1 = DTA3.ScalarQuery_GetJmlItemBeforeBundling(vPrimaryProduct, vPrimaryUnit)
                If RetVal1.HasValue = True Then
                    Retval += RetVal1.Value
                End If
            End If

            'stock transfer
            If vApakahMulti = False Then
                RetVal1 = DTA3.ScalarQuery_GetJmlQtyTransfer(vPrimaryProduct, vPrimaryUnit)
                If RetVal1.HasValue = True Then
                    Retval += RetVal1.Value
                End If

            End If

            DecimalQty = Me.ObjCD.GetJmlPrecision(Retval)
            If DecimalQty > vDecimal Then
                vDecimal = DecimalQty
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Public Function ApakahSudahPernahDigunakan2(ByVal vPrimaryProduct As Double) As Boolean
        Dim Retval As Decimal = 0

        Try
            'cek, jika sudah ada penggunaan code no di:
            '1. po
            Dim DTA As New dseMasterProductTableAdapters.tbltrans_po_itemTableAdapter
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlPOItemByProduct(vPrimaryProduct).Value > 0 Then
                DTA.Dispose()
                Return True
            End If
            DTA.Dispose()

            '2. so
            Dim DTA1 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA1.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA.Dispose()
                Return True
            End If
            DTA1.Dispose()

            '3. Purchase request
            Dim DTA2 As New dseMasterProductTableAdapters.tbltrans_purchase_req_itemTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If DTA2.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA2.Dispose()
                Return True
            End If
            DTA2.Dispose()

            '4. Quotation
            Dim DTA3 As New dseMasterProductTableAdapters.tbltrans_quotation_itemTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA3.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA3.Dispose()
                Return True
            End If
            DTA3.Dispose()

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'penambahan untuk perbaikan master product
    Public Function ApakahUnitSudahPernahDigunakan(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, _
    Optional ByVal vApakahMulti As Boolean = False) As Boolean
        Dim Retval As Decimal = 0

        Try
            'cek, jika sudah ada penggunaan code no di:
            '1. po
            Dim DTA As New dseMasterProductTableAdapters.tbltrans_po_itemTableAdapter
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval = DTA.ScalarQuery_GetJmlByPrimaryProductNUnit(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval = DTA.ScalarQuery_GetJmlByPrimaryProductNUnitMulti(vPrimaryProduct, vPrimaryUnit).Value
            End If
            If Retval > 0 Then
                DTA.Dispose()
                Return True
            End If
            DTA.Dispose()

            '2. so
            Dim DTA1 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval = DTA1.ScalarQuery_GetJmlByPrimaryProductNUnit(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval = DTA1.ScalarQuery_GetJmlByPrimaryProductNUnitMulti(vPrimaryProduct, vPrimaryUnit).Value
            End If
            If Retval > 0 Then
                DTA.Dispose()
                Return True
            End If
            DTA1.Dispose()

            '3. Purchase request
            Dim DTA2 As New dseMasterProductTableAdapters.tbltrans_purchase_req_itemTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA2.ScalarQuery_GetJmlByPrimaryProductNUnit(vPrimaryProduct, vPrimaryUnit).Value
            If Retval > 0 Then
                DTA2.Dispose()
                Return True
            End If
            DTA2.Dispose()

            '4. Quotation
            Dim DTA3 As New dseMasterProductTableAdapters.tbltrans_quotation_itemTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            If vApakahMulti = False Then
                Retval = DTA3.ScalarQuery_GetJmlByPrimaryProductNUnit(vPrimaryProduct, vPrimaryUnit).Value
            Else
                Retval = DTA3.ScalarQuery_GetJmlByPrimaryProductNUnitMulti(vPrimaryProduct, vPrimaryUnit).Value
            End If
            If Retval > 0 Then
                DTA3.Dispose()
                Return True
            End If
            DTA3.Dispose()

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml_WorkUnitAllocated(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByProduct(vPrimaryProduct)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeDependency(ByVal vValue As String) As EnumTipeDependency
        Try
            If vValue = "End-Start" Then
                Return EnumTipeDependency.End_Start1
            ElseIf vValue = "Start-Start" Then
                Return EnumTipeDependency.Start_Start1
            ElseIf vValue = "Start-Start (Strict)" Then
                Return EnumTipeDependency.Start_Start_Strict1
            Else
                Return EnumTipeDependency.End_Start1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeDependencyStr(ByVal vValue As EnumTipeDependency) As String
        Try
            If vValue = EnumTipeDependency.End_Start1 Then
                Return "End-Start"
            ElseIf vValue = EnumTipeDependency.Start_Start1 Then
                Return "Start-Start"
            ElseIf vValue = EnumTipeDependency.Start_Start_Strict1 Then
                Return "Start-Start (Strict)"
            Else
                Return "End-Start"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "End-Start"
        End Try
    End Function

    Public Sub FillTBLProduct(ByVal vTBLProduct2 As dseMasterProduct.tblmaster_product2DataTable)
        Dim DTAProduct2 As New Core_Module.dseMasterProductTableAdapters.tblmaster_product2TableAdapter

        Try
            DTAProduct2.Connection.ConnectionString = Core_Module.GetConstring
            vTBLProduct2.Rows.Clear()
            DTAProduct2.Fill(vTBLProduct2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehEditDecimalWeight(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitWeight As Double, ByVal vDefPrec As Integer) As Boolean
        Dim ObjCB As New ClsMasterCekBatasan
        Dim MaxDec As Integer = 0

        Try
            MaxDec = ObjCB.GetMaxDecimalWeight(vPrimaryProduct, vPrimaryUnit, vPrimaryUnitWeight)
            If MaxDec > vDefPrec Then
                MsgBox("Minimum value you must can enter is  " & MaxDec, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCB = Nothing
        End Try
    End Function

    Public Function ApakahBolehEditDecimalVolume(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitVolume As Double, ByVal vDefPrec As Integer) As Boolean

        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehEditUnitWeight(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitWeight As Double) As Boolean
        Dim ObjCB As New ClsMasterCekBatasan
        Dim Jml As Integer = 0

        Try
            Jml = ObjCB.GetJmlWeight(vPrimaryProduct, vPrimaryUnit, vPrimaryUnitWeight)
            If Jml > 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCB = Nothing
        End Try
    End Function

    Public Function ApakahBolehEditUnitVolume(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryUnitVolume As Double) As Boolean

        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'mengembalikan nilai = 1, jika product ybs menggunakan unit multi untuk perhitungan harganya
    Public Function GetPriceByMultiUnit(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseMasterProduct1TableAdapters.tblmaster_productqtyTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            RetVal = DTA.ScalarQuery_GetPriceByMultiUnit(vPrimaryProduct)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
