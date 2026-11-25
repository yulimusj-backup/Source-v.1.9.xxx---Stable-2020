Option Explicit On
Option Strict Off

Public Class ClsTrans_StockBundling

    Dim mPrimary As Double = -1
    Dim mNoUrut As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mTanggalTransaksi As Date = Now
    Dim mRemarks As String = ""
    Dim mPrimaryProses As Double = -1
    Dim mcode_process As String = ""
    Dim mdesc_process As String = ""
    Dim mPrimaryWorkUnit As Double = -1
    Dim mcode_wu As String = ""
    Dim mPrimaryWorkShift As Double = -1
    Dim mcode_ws As String = ""

    Dim mJono As String = ""
    Dim mMono As String = ""
    Dim mWono As String = ""

    Dim mIdentifierWO As String = ""
    Dim mIdentifierJO As String = ""
    Dim mGoodPercentage As Decimal = 0
    Dim mTotalReject As Decimal = 0
    Dim mTotalGood As Decimal = 0
    Dim mTotalProduce As Decimal = 0
    Dim mProdRateActual As Double = 0
    Dim mDownTimeActual As Double = 0
    Dim mWorkHourActual As Double = 0
    Dim mSetupTimeActual As Double = 0
    Dim mEndTimeActual As Date = Now
    Dim mStartTimeActual As Date = Now
    Dim mProdRatePlanned As Double = 0
    Dim mWorkHourPlanned As Double = 0
    Dim mSetupTimePlanned As Double = 0
    Dim mEndTimePlanned As Date = Now
    Dim mStartTimePlanned As Date = Now
    Dim mUrutan As Double = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mTipeBundling As EnumTipeBundling = EnumTipeBundling.Inventory

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim ObjCD As New Core_Module.ClsCekData

    Dim objnumbering As New Core_Module.ClsMasterTransactionNumbering

    Dim tblStatusItemWo As New DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch1DataTable
    Dim tblset_status As New DseTrans_StockBundling.tblset_statusDataTable
    Dim rowset_status As DseTrans_StockBundling.tblset_statusRow

    Public Enum EnumTipeBundling As Integer
        Inventory
        Purchase_obs
        Produksi1
    End Enum

    Public Enum enumtipe_reject
        reject_processoutput
        reject_rm
    End Enum

    Dim objMO As New ClsTrans_ManufacturingOrder

    Public Property TipeBundling() As EnumTipeBundling
        Get
            Return Me.mTipeBundling
        End Get
        Set(ByVal value As EnumTipeBundling)
            Me.mTipeBundling = value
        End Set
    End Property

    Public Property WoNo() As String
        Get
            Return Me.mWono
        End Get
        Set(ByVal value As String)
            mWono = value
        End Set
    End Property

    Public Property code_process() As String
        Get
            Return Me.mcode_process
        End Get
        Set(ByVal value As String)
            mcode_process = value
        End Set
    End Property

    Public Property desc_process() As String
        Get
            Return Me.mdesc_process
        End Get
        Set(ByVal value As String)
            mdesc_process = value
        End Set
    End Property

    Public Property code_wu() As String
        Get
            Return Me.mcode_wu
        End Get
        Set(ByVal value As String)
            mcode_wu = value
        End Set
    End Property

    Public Property code_ws() As String
        Get
            Return Me.mcode_ws
        End Get
        Set(ByVal value As String)
            mcode_ws = value
        End Set
    End Property


    Public Property Jono() As String
        Get
            Return Me.mJono
        End Get
        Set(ByVal value As String)
            mJono = value
        End Set
    End Property

    Public Property Mono() As String
        Get
            Return Me.mMono
        End Get
        Set(ByVal value As String)
            mMono = value
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

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryProses() As Double
        Get
            Return mPrimaryProses
        End Get
        Set(ByVal value As Double)
            mPrimaryProses = value
        End Set
    End Property

    Public Property PrimaryWorkUnit() As Double
        Get
            Return mPrimaryWorkUnit
        End Get
        Set(ByVal value As Double)
            mPrimaryWorkUnit = value
        End Set
    End Property

    Public Property PrimaryWorkshift() As Double
        Get
            Return mPrimaryWorkShift
        End Get
        Set(ByVal value As Double)
            mPrimaryWorkShift = value
        End Set
    End Property

    Public Property IdentifierWO() As String
        Get
            Return Me.mIdentifierWO
        End Get
        Set(ByVal value As String)
            mIdentifierWO = value
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

    Public Property TotalReject() As Decimal
        Get
            Return Me.mTotalReject
        End Get
        Set(ByVal value As Decimal)
            mTotalReject = value
        End Set
    End Property

    Public Property TotalGood() As Decimal
        Get
            Return Me.mTotalGood
        End Get
        Set(ByVal value As Decimal)
            mTotalGood = value
        End Set
    End Property

    Public Property TotalProduce() As Decimal
        Get
            Return Me.mTotalProduce
        End Get
        Set(ByVal value As Decimal)
            mTotalProduce = value
        End Set
    End Property

    Public Property ProdRateActual() As Double
        Get
            Return Me.mProdRateActual
        End Get
        Set(ByVal value As Double)
            mProdRateActual = value
        End Set
    End Property

    Public Property DownTimeActual() As Double
        Get
            Return Me.mDownTimeActual
        End Get
        Set(ByVal value As Double)
            mDownTimeActual = value
        End Set
    End Property

    Public Property WorkHourActual() As Double
        Get
            Return Me.mWorkHourActual
        End Get
        Set(ByVal value As Double)
            mWorkHourActual = value
        End Set
    End Property

    Public Property SetupTimeActual() As Double
        Get
            Return Me.mSetupTimeActual
        End Get
        Set(ByVal value As Double)
            mSetupTimeActual = value
        End Set
    End Property

    Public Property EndTimeActual() As Date
        Get
            Return Me.mEndTimeActual
        End Get
        Set(ByVal value As Date)
            mEndTimeActual = value
        End Set
    End Property

    Public Property StartTimeActual() As Date
        Get
            Return Me.mStartTimeActual
        End Get
        Set(ByVal value As Date)
            mStartTimeActual = value
        End Set
    End Property

    Public Property ProdRatePlanned() As Double
        Get
            Return Me.mProdRatePlanned
        End Get
        Set(ByVal value As Double)
            mProdRatePlanned = value
        End Set
    End Property

    Public Property WorkHourPlanned() As Double
        Get
            Return Me.mWorkHourPlanned
        End Get
        Set(ByVal value As Double)
            mWorkHourPlanned = value
        End Set
    End Property

    Public Property SetupTimePlanned() As Double
        Get
            Return Me.mSetupTimePlanned
        End Get
        Set(ByVal value As Double)
            mSetupTimePlanned = value
        End Set
    End Property

    Public Property EndTimePlanned() As Date
        Get
            Return Me.mEndTimePlanned
        End Get
        Set(ByVal value As Date)
            mEndTimePlanned = value
        End Set
    End Property

    Public Property StartTimePlanned() As Date
        Get
            Return Me.mStartTimePlanned
        End Get
        Set(ByVal value As Date)
            mStartTimePlanned = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
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

    Public Property IdentifierJO() As String
        Get
            Return Me.mIdentifierJO
        End Get
        Set(ByVal value As String)
            mIdentifierJO = value
        End Set
    End Property

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksi(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoTransaksi(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItemB4 As DataGridView, ByVal vDGVItemAfter As DataGridView, ByVal vfrm As Inventory_Module.FrmTrans_StockBundling, _
    ByVal tbl_workunit As DseTrans_StockBundling.tblworkunitDataTable) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_caridetailTableAdapter
        Dim TBL As New DseTrans_StockBundling.tbltransaksi_stockbundling_caridetailDataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_caridetailRow
        Dim ObjItem As New clsTrans_StockBundlingItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryProses = .primary_process
                    Me.mPrimaryWorkUnit = .primary_work_unit
                    Me.mPrimaryWorkShift = .primary_workshift
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mTipeBundling = .Tipe_Bundling

                    Me.mcode_process = .process_code
                    Me.mdesc_process = .process_desc
                    Me.mMono = .MO_Number
                    Me.mJono = .JO_Number
                    Me.mcode_wu = .wu_code
                    Me.mWono = .WO_Number

                    vfrm.isBalanceProduce = Me.ObjCD.DigitGrouping(.Balance_Produce, ._Decimal)
                    vfrm.isDecimal = ._Decimal
                    vfrm.isProductCode = .Product_Code

                    Me.mIdentifierWO = .identifier_wo
                    Me.mIdentifierJO = .identifier_jo
                    Me.mGoodPercentage = .good_percentage
                    Me.mTotalReject = .total_reject
                    Me.mTotalGood = .total_good
                    Me.mTotalProduce = .total_produce
                    '-> actual
                    Me.mProdRateActual = .prod__index_actual
                    Me.mDownTimeActual = .down_time_actual
                    Me.mWorkHourActual = .work_hour_actual
                    Me.mSetupTimeActual = .setup_time_actual
                    Me.mEndTimeActual = .end_time_actual
                    Me.mStartTimeActual = .start_time_actual
                    '-> planning
                    Me.mStartTimePlanned = .start_time_planned
                    Me.mEndTimePlanned = .end_time_planned
                    Me.mProdRatePlanned = .prod__index_planned
                    Me.mWorkHourPlanned = .work_hour_planned
                    Me.mSetupTimePlanned = .setup_time_planned

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                Me.IsikanTblWorkUnit(tbl_workunit, mIdentifierWO, mIdentifierJO)
                'salin data item 
                ObjItem.IsiDGV(vDgvItemB4, vDGVItemAfter, Me.mPrimary, vfrm, Me.mPrimaryWorkUnit)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItemB4, vDGVItemAfter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True, _
    Optional ByVal vStrictEdit As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim DTA As New DseTrans_StockBundlingTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTrans_StockBundling.DataTable4DataTable
        Dim row As DseTrans_StockBundling.DataTable4Row
        Dim DTA1 As New DseTrans_StockBundlingTableAdapters.DataTable5TableAdapter
        Dim TBL1 As New DseTrans_StockBundling.DataTable5DataTable
        Dim row1 As DseTrans_StockBundling.DataTable5Row
        Dim ProductCode As String = ""
        Dim DTA2 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
            If Me.objMO.IsJOFinished(Me.mIdentifierJO) Then
                MsgBox("JO for WO " & Me.objMO.GetWONumber(Me.IdentifierWO, Me.mIdentifierJO) & " is already closed, you can't edit log for this WO.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            'msgbox("belum")
            'Suatu transaksi in/out tidak bisa diedit jika:
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
            DTA1.Fill(TBL1, Me.mPrimary)
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    If ProductCode = "" Then
                        ProductCode = row1.Product_Code
                    Else
                        ProductCode = ProductCode + ", " + row1.Product_Code
                    End If
                Next

                MsgBox("Code no " + ProductCode + " have modified it's control state. You cannot edit this data", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            '1. jika strict edit maka tidak bisa edit transaksi
            If vStrictEdit = True Then
                If DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) = "" Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction has been deleted", MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                ElseIf DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) <> vUser Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be editted by: " + DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary), MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal vStrictEdit As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim DTA As New DseTrans_StockBundlingTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTrans_StockBundling.DataTable4DataTable
        Dim row As DseTrans_StockBundling.DataTable4Row
        Dim DTA1 As New DseTrans_StockBundlingTableAdapters.DataTable5TableAdapter
        Dim TBL1 As New DseTrans_StockBundling.DataTable5DataTable
        Dim row1 As DseTrans_StockBundling.DataTable5Row
        Dim ProductCode As String = ""
        Dim DTA2 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
            If Me.objMO.IsJOFinished(Me.mIdentifierJO) Then
                MsgBox("JO for WO " & Me.objMO.GetWONumber(Me.IdentifierWO, Me.mIdentifierJO) & " is already closed, you can't edit log for this WO.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring

            'msgbox("belum")
            'transaksi tidak bisa dihapus bila:
            '0. jika bukan yang bersangkutan yang melakukan proses edit data

            '1. jika sudah ada modifikasi control state untuk product tertentu
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

            DTA1.Fill(TBL1, Me.mPrimary)
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    If ProductCode = "" Then
                        ProductCode = row1.Product_Code
                    Else
                        ProductCode = ProductCode + ", " + row1.Product_Code
                    End If
                Next

                MsgBox("Code no " + ProductCode + " have modified it's control state. You cannot delete this data", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            '1. jika strict edit maka tidak bisa edit transaksi
            If vStrictEdit = True Then
                If DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) = "" Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction has been deleted", MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                ElseIf DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) <> vUser Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be deleted by: " + DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary), MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            '3. jika menghapus transaksi ini akan mengakibatkan data product menjadi minus
            Dim ObjItem As New clsTrans_StockBundlingItem
            If ObjItem.ApakahHapusMengakibatkanMinus(Me.mPrimary) = True Then
                ObjItem = Nothing
                Return False
            End If
            ObjItem = Nothing

            '4. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. 

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            TBL.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vNoTransaksi)
            If Dummy.HasValue = True Then
                Return Dummy.Value
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vFrm As FrmTrans_StockBundling, ByVal vDGVItemAfter As DataGridView, _
    ByVal vDGVItemBefore As DataGridView, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim Obj1 As New Core_Module.clsMasterListedSoftware

        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAfterLama As DseTrans_StockBundling.tbltransaksi_stockbundling_item_afterRow

        Dim DTA1 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBL1 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBeforeLama As DseTrans_StockBundling.tbltransaksi_stockbundling_item_beforeRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
        Dim rowQty As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""

        Try
            If Me.objMO.IsJOFinished(Me.mIdentifierJO) Then
                MsgBox("JO for WO " & Me.objMO.GetWONumber(Me.IdentifierWO, Me.mIdentifierJO) & " is already closed, you can't edit log for this WO.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Produksi1)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Produksi1)
                    End If
                End If
            End If

            '2. Jika semua transaksi inventory sedang dikunci
            If Obj1.IsInventoryLocked = True Then
                If TampilkanPesan = True Then
                    MsgBox("Other user have locked inventory database for maintenance purposes. Please wait or contact your administrator", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif, untuk item after
            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.Fill(TBL, Me.mPrimary)
                For Each rowAfterLama In TBL.Rows
                    With rowAfterLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'cek balance rm berkaitan dengan produksi
            If ApakahQtyRMUsageSudahBenar(vFrm, vFrm.TxtPrimaryStockBundling.Text, vFrm.txtidentifier_wo.Text, vFrm.TxtIdentifierJO.Text) = False Then
                Return False
            End If
            '
            'load data item baru
            For Each rowDGV In vDGVItemAfter.Rows
                If IsNothing(rowDGV.Cells(vFrm.ColBatchNoAfter.Index).Value) = True Then
                    BatchNo = ""
                Else
                    BatchNo = rowDGV.Cells(vFrm.ColBatchNoAfter.Index).Value
                End If
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocationAfter.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiryAfter.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiryAfter.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value
                DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
            Next

            'simulasikan untuk item after
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowAfterLama In TBL.Rows
                    With rowAfterLama
                        Me.Koreksi4HapusIn(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItemAfter.Rows
                BatchNo = rowDGV.Cells(vFrm.ColBatchNoAfter.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocationAfter.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiryAfter.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiryAfter.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value
                Qty = rowDGV.Cells(vFrm.ColQtyAfter.Index).Value

                Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, -1, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While
                ObjProduct = Nothing

                MsgBox("The following product(s) from item after will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            '4. Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif, untuk item before
            'load data item lama
            TBLMasterProductQty.Rows.Clear()
            TBL.Rows.Clear()

            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL1.Rows.Clear()
                DTA1.Fill(TBL1, Me.mPrimary)
                For Each rowBeforeLama In TBL1.Rows
                    With rowBeforeLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
                'persiapan untuk pengurangan item after lama
                DTA.Fill(TBL, Me.mPrimary)
            End If
            'load data item baru
            For Each rowDGV In vDGVItemBefore.Rows
                BatchNo = rowDGV.Cells(vFrm.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
            Next

            'nolkan data item gudang yang lebih kecil dr qty after dengan item after lama -> untuk membatasi jika ada code no before yang sama dengan code no after
            'ini terkait siklik bundling
            'lihat batasan pada form, apakah data benar atau apakah edit sah
            If TBL.Rows.Count > 0 And TBLMasterProductQty.Rows.Count > 0 Then
                For ij As Integer = TBLMasterProductQty.Rows.Count - 1 To 0
                    rowQty = TBLMasterProductQty.Rows(ij)
                    TBL.DefaultView.RowFilter = "[Primary Product] = " + rowQty.Primary_Product.ToString + _
                    " AND [Batch No] = '" + rowQty.Lot_No + "'" + _
                    " AND [Primary Storage Area] = " + rowQty.Primary_Storage_Area.ToString

                    If TBL.DefaultView.Count > 0 Then
                        rowAfterLama = TBL.DefaultView(0).Row
                        '1. nolkan qty gudang
                        '2. nol kan qty after
                        If rowQty.Qty < rowAfterLama.Qty Then
                            TBLMasterProductQty.Rows.Remove(rowQty)
                            TBL.Rows.Remove(rowAfterLama)
                        End If
                    ElseIf TBL.DefaultView.Count > 1 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
            End If

            'simulasikan untuk item before --> stock out
            'simulasi hapus data item lama
            If TBL1.Rows.Count > 0 Then
                For Each rowBeforeLama In TBL1.Rows
                    With rowBeforeLama
                        Me.Koreksi4HapusOut(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItemBefore.Rows
                BatchNo = rowDGV.Cells(vFrm.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(vFrm.ColQty.Index).Value

                Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, ExpiryDate)
            Next

            'kurangkan dengan item after lama -> untuk membatasi jika ada code no before yang sama dengan code no after
            'ini terkait siklik bundling
            If TBL.Rows.Count > 0 Then
                For Each rowAfterLama In TBL.Rows
                    With rowAfterLama
                        If Me.ApakahAdaInputYgSama(rowAfterLama.Primary_Product, rowAfterLama.Batch_No, rowAfterLama.Primary_Storage_Area, _
                        vFrm, vDGVItemBefore) = True Then
                            Me.Koreksi4HapusIn(TBLMasterProductQty, _
                                .Primary_Product, .Qty, .Secondary_Qty, _
                                .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                        End If
                    End With
                Next
            End If

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While
                ObjProduct = Nothing

                MsgBox("The following product(s) from item before will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'If Me.CekPerhitunganTanpaBatas1(vBaru, vDGVItemBefore, vDGVItemAfter, vFrm) = False Then
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj1 = Nothing
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
        End Try
    End Function

    Private Function ApakahAdaInputYgSama(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String, ByVal vPrimaryLocation As Double, _
    ByVal vFrm As FrmTrans_StockBundling, ByVal vDGVItemBefore As DataGridView) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGVItemBefore.Rows
                If rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                rowDGV.Cells(vFrm.ColBatch.Index).Value = vBatchNo And _
                rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function AdaSiklikBundling(ByVal vPrimaryCodeNoAfter As Double, ByVal vPrimaryCodeNoFind As Double, ByRef vDummy As String, _
    ByVal vTBL As DseTrans_StockBundling.DataTable6DataTable) As Boolean
        Dim Filter As String = ""
        Dim ListPrimary As New List(Of Double)
        Dim row As DseTrans_StockBundling.DataTable6Row
        Dim PrimaryProductAfterNew As Double = -1
        Dim dummyLama As String = ""

        Try
            '1. kumpulkan semua item yang primary code no before = primary code after pada parameter
            Filter = "[Primary Product Before] = " + vPrimaryCodeNoAfter.ToString
            vTBL.DefaultView.RowFilter = Filter
            For Each dv As Data.DataRowView In vTBL.DefaultView
                row = dv.Row
                ListPrimary.Add(row.Primary_Product_After)
            Next

            If ListPrimary.Count > 0 Then
                '2. jika ada item after selanjutnya, untuk semua item code no after yang selanjutnya, 
                For Each PrimaryProductAfterNew In ListPrimary
                    dummyLama = vDummy
                    vDummy = vDummy + " -> " + Me.GetCodeNo(PrimaryProductAfterNew, vTBL)
                    'a. cari apakah item code no after selanjutnya = primary code no find 
                    If PrimaryProductAfterNew <> vPrimaryCodeNoFind Then
                        'b. jika tidak sama maka jalankan fungsi ini kembali rekursif
                        If Me.AdaSiklikBundling(PrimaryProductAfterNew, vPrimaryCodeNoFind, vDummy, vTBL) = True Then
                            Return True
                        Else
                            vDummy = dummyLama
                        End If
                    Else
                        'c. jika ada maka kembalikan nilai true, terjadi siklik
                        Return True
                    End If
                Next

                Return False
            Else
                '3. jika tidak ada item after selanjutnya, maka return false
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Function

    Private Function GetCodeNo(ByVal vPrimaryCodeNoAfter As Double, ByVal vTBL As DseTrans_StockBundling.DataTable6DataTable) As String
        Dim Filter As String = ""
        Dim row As DseTrans_StockBundling.DataTable6Row

        Try
            Filter = "[Primary Product After] = " + vPrimaryCodeNoAfter.ToString
            vTBL.DefaultView.RowFilter = Filter

            If vTBL.DefaultView.Count = 0 Then
                MsgBox("Logic error, contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return ""
            Else
                row = vTBL.DefaultView(0).Row
                Return row.Code_No_After
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub HilangkanItem(ByVal TBLDummy As DseTrans_StockBundling.DataTable3DataTable)
        Dim Filter As String = ""
        Filter = "[Primary Transaksi] = " + Me.mPrimary.ToString

        Try
            TBLDummy.DefaultView.RowFilter = Filter
            While TBLDummy.DefaultView.Count > 0
                TBLDummy.Rows.Remove(TBLDummy.DefaultView(0).Row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBLDummy As DseTrans_StockBundling.DataTable3DataTable, ByVal DGVBefore As DataGridView, _
    ByVal DGVAfter As DataGridView, ByVal i As FrmTrans_StockBundling)
        Dim rowDGV_B4 As DataGridViewRow
        Dim rowDGV_After As DataGridViewRow
        Dim rowBundling As DseTrans_StockBundling.DataTable3Row

        Try
            For Each rowDGV_B4 In DGVBefore.Rows
                For Each rowDGV_After In DGVAfter.Rows
                    rowBundling = TBLDummy.NewDataTable3Row
                    rowBundling.Primary_Product_After = rowDGV_After.Cells(i.ColPrimaryCodeNoAfter.Index).Value
                    rowBundling.Primary_Product_Before = rowDGV_B4.Cells(i.ColPrimaryCodeNo.Index).Value
                    rowBundling.Primary_Transaksi = Me.mPrimary
                    rowBundling.Tanggal_Bundling = Me.mTanggalTransaksi
                    TBLDummy.Rows.Add(rowBundling)
                Next
            Next
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
            vSecQty.ToString + " and [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(vBatchNo) + "' and [Primary Storage Area] = " + _
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
            vSecQty.ToString + " and [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(vBatchNo) + "' and [Primary Storage Area] = " _
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

    Public Overridable Function SimpanBaru(ByVal vdgvItemB4 As DataGridView, ByVal vdgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTrans_StockBundling) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockBundling_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vdgvItemB4, vdgvItemAfter, vUserCode, vFrm)
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
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItemB4 As DataGridView, ByVal vdgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTrans_StockBundling) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjItem As New clsTrans_StockBundlingItem
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling
        Dim StartTime As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, vFrm, vdgvItemAfter, vdgvItemB4) = True Then
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling
                Me.mUrutan = Me.objnumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)
                'insert query
                If DTA.Insert(Me.mNoTransaksi, Me.mUrutan, Me.mTanggalTransaksi, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, Me.mTipeBundling, _
                              Me.mPrimaryWorkUnit, Me.mPrimaryProses, Me.mStartTimePlanned, Me.mEndTimePlanned, Me.mSetupTimePlanned, Me.mWorkHourPlanned, _
                              Me.mProdRatePlanned, Me.mStartTimeActual, Me.mEndTimeActual, Me.mSetupTimeActual, Me.mWorkHourActual, Me.mDownTimeActual, _
                              Me.mProdRateActual, Me.mTotalProduce, Me.mTotalGood, Me.mTotalReject, Me.mGoodPercentage, Me.mPrimaryWorkShift, Me.mIdentifierWO, _
                              Me.mIdentifierJO, Me.mIdentifierJO.Substring(0, Me.mIdentifierJO.IndexOf("-"))) > 0 Then

                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'cek berkaitan nomor sama untuk beda bulan
                    Me.CekNomor(DTA)
                    'simpan data item baru
                    StartTime = Now
                    ObjItem.Simpan(vdgvItemB4, vdgvItemAfter, Me.mTanggalTransaksi, Me.mPrimary, _
                    Me.mNoTransaksi, vFrm)
                    'MsgBox("Duration for saving item: " + ObjCD.DigitGrouping((Now - StartTime).TotalSeconds, 1) + "seconds", MsgBoxStyle.Information, "Information")

                    'update work_order
                    StartTime = Now
                    Update_FollowUpFromBundling(Me.mIdentifierWO, Me.PrimaryWorkUnit, Me.mIdentifierJO)
                    'MsgBox("Duration for updating production order balance: " + ObjCD.DigitGrouping((Now - StartTime).TotalSeconds, 1) + "seconds", MsgBoxStyle.Information, "Information")

                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.objnumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Produksi1)
                    Return Me.SimpanBaru1(vdgvItemB4, vdgvItemAfter, vUserCode, vFrm)
                Else
                    MsgBox("Bundling Number duplication, please change Bundling number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mUrutan
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling

        Try
            TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling

            Me.objnumbering.GetTanggalDariSampai(TanggalDari, TanggalSampai, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling, Me.mTanggalTransaksi)
            'TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            'TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If DTA.ScalarQuery_GetJmlByUrutan(Me.mUrutan, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Produksi1)).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mUrutan, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Produksi1)).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mUrutan, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Produksi1)).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mUrutan, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Produksi1))
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Produksi1)
                        Me.mUrutan = Me.objnumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi) 'CInt(Me.mNoTransaksi.Substring(0, 5))
                    End If
                End If
                LoopCount += 1

                If LoopCount > 50 Then
                    MsgBox("Logic error: invalid function: 'Cek Nomor' on clstrans_bundling_produksi, please contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mNoTransaksi, Me.mUrutan, Me.mPrimary)
            End If 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItemB4 As DataGridView, ByVal dgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTrans_StockBundling) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockBundling_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItemB4, dgvItemAfter, vUserCode, vFrm)
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
            'ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItemB4 As DataGridView, ByVal dgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTrans_StockBundling) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjItem As New clsTrans_StockBundlingItem
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling
        Dim IdentifierWOLama, IdentifierJOLama As String
        Dim PrimaryWULama As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(False, vFrm, dgvItemAfter, dgvItemB4) = True Then
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling
                Me.mUrutan = Me.objnumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)
                IdentifierWOLama = DTA.ScalarQuery_GetIdentifierWO(Me.mPrimary)
                IdentifierJOLama = DTA.ScalarQuery_GetIdentifierJO(Me.mPrimary)
                PrimaryWULama = DTA.ScalarQuery_GetPrimaryWorkUnit(Me.mPrimary)

                If DTA.Update(Me.mNoTransaksi, Me.mUrutan, Me.mTanggalTransaksi, Me.mRemarks, vUserCode, Now, Now, _
                              Me.mPrimaryWorkUnit, Me.mPrimaryProses, Me.mStartTimePlanned, Me.mEndTimePlanned, Me.mSetupTimePlanned, Me.mWorkHourPlanned, _
                              Me.mProdRatePlanned, Me.mStartTimeActual, Me.mEndTimeActual, Me.mSetupTimeActual, Me.mWorkHourActual, Me.mDownTimeActual, _
                              Me.mProdRateActual, Me.mTotalProduce, Me.mTotalGood, Me.mTotalReject, Me.mGoodPercentage, Me.mPrimaryWorkShift, Me.mIdentifierWO, _
                              Me.mIdentifierJO, Me.mIdentifierJO.Substring(0, Me.mIdentifierJO.IndexOf("-")), _
                              Me.mPrimary, Me.mTimestamp) > 0 Then

                    'update data yang berubah setelah simpan perubahan
                    Me.mUrutan = Me.objnumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)
                    Me.mTimestamp = get_timestamp1(Me.mPrimary) 'DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, vFrm)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItemB4, dgvItemAfter, Me.mTanggalTransaksi, Me.mPrimary, Me.mNoTransaksi, vFrm)
                    'update work_order
                    If String.IsNullOrEmpty(IdentifierJOLama) = False And String.IsNullOrEmpty(IdentifierWOLama) = False Then
                        If Me.mIdentifierWO <> IdentifierWOLama Or _
                        Me.mIdentifierJO <> IdentifierJOLama Then
                            Update_FollowUpFromBundling(IdentifierWOLama, PrimaryWULama, IdentifierJOLama)
                        End If
                    End If
                    Update_FollowUpFromBundling(Me.mIdentifierWO, Me.PrimaryWorkUnit, Me.mIdentifierJO)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.objnumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Produksi1)
                    Return Me.SimpanPerubahan1(dgvItemB4, dgvItemAfter, vUserCode, vFrm)
                Else
                    MsgBox("Bundling Number duplication, please change Bundling number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function get_timestamp1(ByVal vprimary As Double) As Date
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim retval1 As New Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval1 = dta.ScalarQuery_GetTimestamp(vprimary)
            If retval1.HasValue = True Then
                Return retval1.Value
            Else
                Return Now
            End If
            Return Now
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vTampilkanPesan As Boolean, ByVal vfrm As Inventory_Module.FrmTrans_StockBundling) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockBundling_Hapus, vUserCode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, vStrictEdit, vTampilkanPesan, vfrm)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, _
    ByVal vStrictEdit As Boolean, ByVal vTampilkanPesan As Boolean, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim Obj As New clsTrans_StockBundlingItem
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim vlist As New List(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vUser, vTampilkanPesan, vStrictEdit) = True Then
                'hapus data item 
                Obj.Hapus(vPrimaryTransaksi, vfrm)
                'hapus data header, data item otomatis dihapuskan oleh server
                For Each row As DataGridViewRow In vfrm.dgvItemb4.Rows
                    If vlist.Contains(row.Cells(vfrm.ColPrimaryCodeNo.Index).Value) = False Then
                        vlist.Add(row.Cells(vfrm.ColPrimaryCodeNo.Index).Value)
                    End If
                Next

                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    'update work_order
                    Update_FollowUpFromBundling(Me.mIdentifierWO, Me.PrimaryWorkUnit, Me.mIdentifierJO)
                    'Obj.update_rm_by(Me.mIdentifierWO, vlist, Me.mIdentifierJO)
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

    Public Function GetTanggalTransaksi(ByVal vPrimaryBundling As Double) As Date
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryBundling).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryBundling).Value
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
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal vTipeBundling As EnumTipeBundling _
    = EnumTipeBundling.Inventory) As Long
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", vTipeBundling).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", vTipeBundling)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari, vTipeBundling).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari, vTipeBundling).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", vTipeBundling).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", vTipeBundling).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Dummy = DTA.ScalarQuery_GetJmlByRemarks(vTipeBundling, Criteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProsesCode Then
                Dummy = DTA.ScalarQuery_GetJmlByProses(Criteria + "%", vTipeBundling)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkUnitCode Then
                Dummy = DTA.ScalarQuery_GetJmlByWorkUnit(Criteria + "%", vTipeBundling)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkShift Then
                Dummy = DTA.ScalarQuery_GetJmlByWorkshift(Criteria + "%", vTipeBundling)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", vTipeBundling).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", vTipeBundling).Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Dummy = Nothing
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTrans_StockBundling.tbltransaksi_stockbundling1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal vTipeBundling As EnumTipeBundling = EnumTipeBundling.Inventory) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_Tanggal(TBL, vTipeBundling, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProsesCode Then
                QueryResult = DTA.FillBy_ProcessLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkUnitCode Then
                QueryResult = DTA.FillBy_WorkunitLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkShift Then
                QueryResult = DTA.FillBy_WorkshiftLIKE(TBL, vTipeBundling, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).Visible = False
                DGV.Columns(6).Visible = False
                'DGV.Columns(5).Visible = False
                DGV.Columns(2).HeaderText = "Transaction No"
                DGV.Columns(3).HeaderText = "Transaction Date"
                DGV.Columns(4).HeaderText = "Process Code"
                DGV.Columns(5).HeaderText = "WorkUnit Code"
                DGV.Columns(6).HeaderText = "Workshift Code"
                DGV.Columns(7).HeaderText = "Remarks"

                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(3).DefaultCellStyle = ij
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

    Public Sub ResetObject(ByVal vDgvItemB4 As DataGridView, ByVal vDgvItemAfter As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mUrutan = -1
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            Me.mPrimaryProses = -1
            Me.mPrimaryWorkUnit = -1
            Me.mPrimaryWorkShift = -1
            Me.mIdentifierWO = ""
            Me.mGoodPercentage = 0
            Me.mTotalReject = 0
            Me.mTotalGood = 0
            Me.mTotalProduce = 0
            Me.mProdRateActual = 0
            Me.mDownTimeActual = 0
            Me.mWorkHourActual = 0
            Me.mSetupTimeActual = 0
            Me.mEndTimeActual = Now
            Me.mStartTimeActual = Now
            Me.mProdRatePlanned = 0
            Me.mWorkHourPlanned = 0
            Me.mSetupTimePlanned = 0
            Me.mEndTimePlanned = Now
            Me.mStartTimePlanned = Now
            Me.mUrutan = 0
            vDgvItemB4.Rows.Clear()
            vDgvItemAfter.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateStockCard_TransaksiStockBundling(ByVal vPrg As ProgressBar, ByVal vLblProgress As Label)
        Dim DTAB4 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling2TableAdapter
        Dim TBLB4 As New DseTrans_StockBundling.tbltransaksi_stockbundling2DataTable
        Dim rowB4 As DseTrans_StockBundling.tbltransaksi_stockbundling2Row
        Dim DTAAfter As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling3TableAdapter
        Dim TBLAfter As New DseTrans_StockBundling.tbltransaksi_stockbundling3DataTable
        Dim rowAfter As DseTrans_StockBundling.tbltransaksi_stockbundling3Row

        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
        Dim Qty_In, Qty_Out As Double
        Dim DescSC As String = "Stock bundling"
        Dim Sec_Qty, PrimaryLocation As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            DTAB4.Connection.ConnectionString = Core_Module.GetConstring
            DTAAfter.Connection.ConnectionString = DTAB4.Connection.ConnectionString
            DTAB4.Fill(TBLB4)
            DTAAfter.Fill(TBLAfter)

            vPrg.Value = 0
            vPrg.Maximum = TBLB4.Rows.Count + TBLAfter.Rows.Count

            'simpan data item b4
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
            For Each rowB4 In TBLB4.Rows
                vPrg.Value += 1
                vLblProgress.Text = "Updating stock bundling stock card data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString + " - Transaction No :" + rowB4.No_Transaksi

                BatchNo = rowB4.Batch_No
                PrimaryLocation = rowB4.Primary_Storage_Area
                Sec_Qty = -1
                ExpiryDate = rowB4.Expiry_Date

                'simpan data stock card
                Qty_In = 0
                Qty_Out = rowB4.Qty
                InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                With rowB4
                    ObjSC.Simpan(.Primary_Transaksi, JenisSC, .No_Transaksi, .Primary_Product, _
                    Qty_In, Qty_Out, DescSC, InOut, .Tanggal_Transaksi, BatchNo, ExpiryDate, PrimaryLocation, "")
                End With
            Next

            'simpan data item after
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
            For Each rowAfter In TBLAfter.Rows
                vPrg.Value += 1
                vLblProgress.Text = "Updating stock bundling stock card data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString + " - Transaction No :" + rowAfter.No_Transaksi

                BatchNo = rowAfter.Batch_No
                PrimaryLocation = rowAfter.Primary_Storage_Area
                Sec_Qty = -1
                ExpiryDate = rowAfter.Expiry_Date
                Qty_In = rowAfter.Qty
                Qty_Out = 0
                InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                With rowAfter
                    ObjSC.Simpan(.Primary_Transaksi, JenisSC, .No_Transaksi, .Primary_Product, _
                    Qty_In, Qty_Out, DescSC, InOut, .Tanggal_Transaksi, BatchNo, ExpiryDate, PrimaryLocation, "")
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAB4.Dispose()
            TBLB4.Dispose()
            DTAAfter.Dispose()
            TBLAfter.Dispose()
            DTA1.Dispose()
            ObjSC = Nothing
        End Try
    End Sub

    '===================================================================================================================================================
    'Nomer bundling setting smess

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vTipeBundling As EnumTipeBundling) As String
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim _retval As String = ""
        Dim _tipetransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling
        Dim DummyVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)
            Me.objnumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _tipetransaksi, vTanggalTransaksi)

            'cari no urut
            DummyVal = DTA.ScalarQuery_GetNoUrut(TanggalEnd, TanggalStart, CInt(vTipeBundling))
            If DummyVal.HasValue = True Then
                NoUrut = CStr(DummyVal.Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            _retval = Me.objnumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _tipetransaksi)
            While Me.ApakahSudahAda(_retval) = True
                NoUrut = NoUrut + 1
                _retval = Me.objnumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _tipetransaksi)
            End While
            Return _retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsiDgvByWo(ByVal vPrimaryBundling As Double, ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTrans_StockBundling.tbltrans_mo_item_wo_findDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vStatus1 As ClsTrans_ManufacturingOrder.EnumPilihStatusWO, ByVal vStatus2 As ClsTrans_ManufacturingOrder.EnumPilihStatusWO, _
    ByVal vStatus3 As ClsTrans_ManufacturingOrder.EnumPilihStatusWO, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean

        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_WoNumberLIKE(TBL, vStatus1, vStatus2, vStatus3, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vStatus1, vStatus2, vStatus3, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, vStatus1, vStatus2, vStatus3, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByJONumber Then
                QueryResult = DTA.FillBy_JONumberLIKE(TBL, vStatus1, vStatus2, vStatus3, Criteria + "%")
            End If

            TBL.DefaultView.RowFilter = Nothing
            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).Visible = False
                DGV.Columns(7).Visible = False
                DGV.Columns(8).Visible = False

                DGV.Columns(2).HeaderText = "JO Number"
                DGV.Columns(3).HeaderText = "WO Number"
                DGV.Columns(5).HeaderText = "Date"

                'DGV.Columns(5).HeaderText = "Remarks"

                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(5).DefaultCellStyle = ij
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

    Public Function GetJmlByWo(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal vstatus1 As Integer, ByVal vstatus2 As Integer, ByVal vstatus3 As Integer, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long

        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Dummy = DTA.ScalarQuery_GetJmlByWoNumberLIKE(Criteria + "%", vstatus1, vstatus2, vstatus3)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Dummy = DTA.ScalarQuery_GetJmlByDate(vTanggalDari, vTanggalSampai, vstatus1, vstatus2, vstatus3)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Dummy = DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%", vstatus1, vstatus2, vstatus3)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByJONumber Then
                Dummy = DTA.ScalarQuery_GetJmlByJONumberLIKE(vstatus1, vstatus2, vstatus3, Criteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Dummy = Nothing
        End Try
    End Function

    Public Function ApakahSudahAdaWO(ByVal vIdentifierWo As String, ByVal vschedule As Integer, ByVal vonprogress As Integer, ByVal vIdentifierJO As String) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim retval As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_GetJmlByIdentifierWo(vschedule, vonprogress, vIdentifierWo, vIdentifierJO)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetIdentifierJOByWoNo(ByVal vWoNo As String) As String
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim retval As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_GetIdentifierJO(vWoNo)
            If String.IsNullOrEmpty(retval) = False Then
                Return retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetIdentifierByWo(ByVal vWoNo As String) As String
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim retval As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_GetIdentifierWo(vWoNo)
            If String.IsNullOrEmpty(retval) = False Then
                Return retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function IsikanTblRawMaterial(ByVal vprimary_product As Double, ByVal videntifierWO As String, _
    ByVal tbl_rm As DseTrans_StockBundling.tbltrans_mo_item_bomDataTable, ByVal vIdentifierJO As String) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_bomTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            dta.FillBy(tbl_rm, videntifierWO, vprimary_product, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    'followup simpan/edit/hapus
    '1. actual result
    Private Function get_sumqtyproduce_alltransactionperwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling4TableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_SumQty_Produce(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    '. actual result 2 total keseluran good per wo di bundling
    Private Function get_sumqtygood_alltransactionperwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_SumQtyGoods(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    '. actual start
    Private Function get_actualstart_perwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Date
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling4TableAdapter
        Dim retval As New Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_MinDate(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return #1/1/1950#
            End If
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            dta.Dispose()
        End Try
    End Function

    '. actual end
    Private Function get_actualend_perwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Date
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling4TableAdapter
        Dim retval As New Nullable(Of Date)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_MaxDate(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return #1/1/1950#
            End If
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            dta.Dispose()
        End Try
    End Function

    '. setup time
    Private Function get_sumqtysetuptime_alltransactionperwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim retval As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_SumSetupTime(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    '. workhour
    Private Function get_sumqtyworkhour_alltransactionperwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling4TableAdapter
        Dim retval As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            retval = dta.ScalarQuery_SumWorkHour(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    '. downtime
    Private Function get_sumqtydowntime_alltransactionperwo(ByVal videntifier_wo As String, ByVal vPrimary_wu As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim retval As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_SumDownTime(videntifier_wo, vPrimary_wu)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    '. status
    Private Function get_targetproduce(ByVal videntifier_wo As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyTargetProduce(videntifier_wo)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function GetStatus(ByVal vactual_produce As Decimal, ByVal vtarget_produce As Decimal) As ClsTrans_ManufacturingOrder.EnumPilihStatusWO
        Dim retval As ClsTrans_ManufacturingOrder.EnumPilihStatusWO = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free
        Dim jml As Decimal = 0
        Try
            jml = vactual_produce / vtarget_produce * 100
            If jml >= 100 Then
                retval = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
            ElseIf jml > 0 And jml < 100 Then
                retval = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress
            ElseIf jml = 0 Then
                retval = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            End If
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '. project result
    Private Function get_targetproduceadditional(ByVal videntifier_wo As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyTargetProduceAdd(videntifier_wo)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function get_gptarget(ByVal videntifier_wo As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetGpTarget(videntifier_wo)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    '. qty used
    Public Function get_sumqtyrm_alltransactionperwo(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_sumqtyrm(vprimary_product, videntifier_wo, vIdentifierJO)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
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

    ''#update ke databasenya untuk woitem dan woitem_sch
    'Private Function update_workorder_by1(ByVal videntifier_wo As String, ByVal vprimary_wu As Double) As Boolean
    '    Dim dta As New DseTrans_StockBundling_ProduksiTableAdapters.tbltrans_mo_item_woTableAdapter
    '    Dim status As Integer
    '    Dim balance_produce, balance_result, actual_setuptime, actual_workhour, actual_totaltime, actual_downtime, actual_prodrate, actual_produce, actual_result, target_produce, _
    '    target_produceadditonal, actual_project_result, actual_gptarget, target_result As Decimal
    '    Dim actual_starttime, actual_endtime As Date

    '    Try
    '        dta.Connection.ConnectionString = Core_Module.GetConstring

    '        target_produce = Me.get_targetproduce(videntifier_wo)
    '        target_produceadditonal = Me.get_targetproduceadditional(videntifier_wo)
    '        target_result = Me.get_qtytargetresult(videntifier_wo)
    '        'actual_gptarget = Me.get_gptarget(videntifier_wo)
    '        actual_starttime = Me.get_actualstart_perwo(videntifier_wo, vprimary_wu)
    '        actual_endtime = Me.get_actualend_perwo(videntifier_wo, vprimary_wu)
    '        actual_produce = Me.get_sumqtyproduce_alltransactionperwo(videntifier_wo, vprimary_wu)
    '        actual_result = Me.get_sumqtygood_alltransactionperwo(videntifier_wo, vprimary_wu)
    '        actual_gptarget = (actual_result * 100) / target_result
    '        actual_setuptime = Me.get_sumqtysetuptime_alltransactionperwo(videntifier_wo, vprimary_wu)
    '        actual_workhour = Me.get_sumqtyworkhour_alltransactionperwo(videntifier_wo, vprimary_wu)
    '        actual_downtime = Me.get_sumqtydowntime_alltransactionperwo(videntifier_wo, vprimary_wu)
    '        actual_totaltime = actual_setuptime + actual_workhour + actual_downtime
    '        balance_produce = target_produce - actual_produce
    '        balance_result = target_result - actual_result

    '        If actual_workhour > 0 Then
    '            actual_prodrate = actual_produce / actual_workhour
    '        ElseIf actual_produce = 0 And actual_workhour = 0 Then
    '            actual_prodrate = 0
    '        End If

    '        Update_ForItemWOSCH(actual_produce, target_produce, videntifier_wo, vprimary_wu, actual_starttime, actual_endtime, actual_setuptime, _
    '         actual_workhour, actual_totaltime, actual_prodrate, actual_downtime, actual_produce, actual_result)

    '        status = Me.SetStatusForItemWo(videntifier_wo)
    '        actual_project_result = (target_produce + target_produceadditonal - actual_produce) * (actual_gptarget / 100) + actual_result
    '        If dta.UpdateQuery_ItemWo(status, actual_starttime, actual_endtime, actual_setuptime, actual_workhour, actual_totaltime, _
    '               actual_downtime, actual_prodrate, actual_produce, actual_result, actual_project_result, actual_gptarget, _
    '              balance_produce, balance_result, vprimary_wu, videntifier_wo) > 0 Then
    '            Return True
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return False
    '    Finally
    '        dta.Dispose()
    '    End Try
    'End Function

    Public Function ApakahCodeNoDiRMSudahAda(ByVal videntifier_wo As String, ByVal vprimaryproduct_rm As Double, ByVal vIdentifierJO As String) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetJmlRM(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)

            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qty_wip_allocated_balance(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetWIPBalance(vIdentifierJO, videntifier_wo, vprimary_product)
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

    Public Function get_qtyrmbalance(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyBalanceRM(videntifier_wo, vprimary_product, vIdentifierJO)
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

    Public Function ApakahQtyRMUsageSudahBenar(ByVal vfrm As Inventory_Module.FrmTrans_StockBundling, ByVal vprimary_bundling As Double, _
    ByVal videntifier_wo As String, ByVal vIdentifierJO As String) As Boolean
        Dim jml_dgv, jml_lama, jml_blc, jml_setelahedit, jml_dgvlain As Decimal
        Dim objproduct As New Core_Module.clsMasterProduct
        Dim dummy As String = ""
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim PrimaryMO As Double = -1

        Try
            For Each vr As DataGridViewRow In vfrm.dgvItemb4.Rows
                If String.IsNullOrEmpty(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value) = False Then
                    If ListPrimaryCodeNo.Contains(CDbl(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value)) = False Then
                        PrimaryMO = vIdentifierJO.Substring(0, vIdentifierJO.IndexOf("-"))
                        If Me.objMO.ApakahWIP(videntifier_wo, vIdentifierJO, CDbl(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value), PrimaryMO, False) = False Then
                            jml_blc = Me.get_qtyrmbalance(videntifier_wo, vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value, vIdentifierJO) + Me.get_qty_wip_allocated_balance(videntifier_wo, vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value, vIdentifierJO)
                            jml_lama = Me.get_qtystockbundlinglama(vprimary_bundling, vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value)
                            jml_dgv = vfrm.get_jmlproduct_indgvb4(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value)
                            jml_setelahedit = jml_blc + jml_lama - jml_dgv

                            If jml_setelahedit < 0 Then
                                MsgBox("Maximum value you can enter for raw material " + vr.Cells(vfrm.ColCodeNo.Index).Value + " is: " & Me.ObjCD.DigitGrouping(jml_blc + jml_lama - jml_dgvlain, _
                                objproduct.GetAngkaDecimal(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value)) & " " & objproduct.GetUnitCode(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value), MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            ListPrimaryCodeNo.Add(CDbl(vr.Cells(vfrm.ColPrimaryCodeNo.Index).Value))
                        End If

                        'Else
                        '    MsgBox("")
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryCodeByWo(ByVal videntifier_wo As String, ByVal vIdentifierJO As String) As Double
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As New Nullable(Of Double)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetPrimaryCodeByWo(videntifier_wo, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return -1
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtystockbundlinglama(ByVal vprimary_bundling As Double, ByVal vPrimaryRM As Double) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim retval As Nullable(Of Decimal)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_getqty(vprimary_bundling, vPrimaryRM)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyallocatedrm(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetAllocatedQty(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyadditionalrm(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetAllocatedQty(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtytargetresult(ByVal videntifier_wo As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyTargetResult(videntifier_wo)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function IsikanTblWorkUnit(ByVal TBL As DseTrans_StockBundling.tblworkunitDataTable, ByVal vIdentifierWO As String, ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tblworkunitTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL, vIdentifierWO, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanTblWo4Header(ByVal TBL As DseTrans_StockBundling.tblwoitemDataTable, ByVal vIdentifierWo As String, ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tblwoitemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL, vIdentifierWo, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanTblWo4Item(ByVal TBL As DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch1DataTable, ByVal vIdentifierWo As String, _
    ByVal vPrimaryWU As Double, ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_item_sch1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL, vPrimaryWU, vIdentifierWo, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Update_FollowUpFromBundling(ByVal IdentifierWO As String, ByVal PrimaryWU As Double, ByVal vIdentifierJO As String) As Boolean
        Dim DtaByBundling As New DseTrans_StockBundlingTableAdapters.tblbundling_byidentifierTableAdapter
        Dim TblByBundling As New DseTrans_StockBundling.tblbundling_byidentifierDataTable
        Dim rowByBundling As DseTrans_StockBundling.tblbundling_byidentifierRow
        Dim DtaBySch As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_item_sch2TableAdapter
        Dim TblBySch As New DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch2DataTable
        Dim rowBySch As DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch2Row
        Dim DtaByWOItem As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo1TableAdapter
        Dim TblByWOItem As New DseTrans_StockBundling.tbltrans_mo_item_wo1DataTable
        Dim rowByWOItem As DseTrans_StockBundling.tbltrans_mo_item_wo1Row
        Dim DtaUpdate_ForItemWO As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim Actual_Produce, Actual_SetupTime, Actual_Downtime, Actual_ProdRate, Actual_WorkHour, Actual_GoodPercentage, _
            Actual_Result, Actual_TotalTime, Actual_ProjectResult, Balance_Produce, Balance_Result, Actual_Reject, Balance_MP As Decimal
        Dim Target_Produce, TargetProduceAdditional, Target_Result, RejectAllowance As Decimal
        Dim Status As Integer = 0
        Dim Actual_StartTime As Date = #1/1/1950#
        Dim Actual_EndTime As Date = #1/1/1950#
        Dim Progress As Decimal = 0
        Dim ObjMOItem As New ClsTrans_ManufacturingOrder_Item

        Try
            'MsgBox("consider identifier jo juga")

            DtaByBundling.Connection.ConnectionString = Core_Module.GetConstring
            DtaByBundling.Fill(TblByBundling, IdentifierWO, vIdentifierJO)
            DtaBySch.Connection.ConnectionString = DtaByBundling.Connection.ConnectionString
            DtaBySch.Fill(TblBySch, IdentifierWO, vIdentifierJO)
            DtaByWOItem.Connection.ConnectionString = DtaByBundling.Connection.ConnectionString
            DtaByWOItem.Fill(TblByWOItem, IdentifierWO, vIdentifierJO)

            TblBySch.DefaultView.RowFilter = Nothing
            TblBySch.DefaultView.RowFilter = "[identifier wo]='" & IdentifierWO & "' AND [primary work unit]=" & PrimaryWU
            If TblBySch.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TblBySch.DefaultView.Count - 1
                    rowBySch = TblBySch.DefaultView(ij).Row
                    Target_Produce += rowBySch.qty_produce
                    TargetProduceAdditional += rowBySch.qty_produce_add
                    Target_Result += rowBySch.plan_result
                Next
            End If

            '1. ada dua yang di update : tblitem_wo dan tblitem_wo_item_sch
            '. untuk item_wo_item_sch *
            TblByBundling.DefaultView.RowFilter = Nothing
            TblByBundling.DefaultView.RowFilter = "[identifier wo]='" & IdentifierWO & "' AND [primary work unit]=" & PrimaryWU
            For ij As Integer = 0 To TblByBundling.DefaultView.Count - 1
                rowByBundling = TblByBundling.DefaultView(ij).Row
                With rowByBundling
                    Actual_Produce += .total_produce
                    Actual_Result += .total_good
                    Actual_Reject += .total_reject
                    Actual_SetupTime += .setup_time_actual
                    Actual_Downtime += .down_time_actual
                    Actual_WorkHour += .work_hour_actual
                    If Actual_StartTime = #1/1/1950# Then
                        Actual_StartTime = .start_time_actual
                    Else
                        If Actual_StartTime > .start_time_actual Then
                            Actual_StartTime = .start_time_actual
                        End If
                    End If
                    If Actual_EndTime = #1/1/1950# Then
                        Actual_EndTime = .end_time_actual
                    Else
                        If Actual_EndTime < .end_time_actual Then
                            Actual_EndTime = .end_time_actual
                        End If
                    End If
                End With
            Next

            Actual_TotalTime = Actual_WorkHour + Actual_Downtime + Actual_SetupTime
            Balance_Produce = Target_Produce - Actual_Produce
            Balance_Result = Target_Result - Actual_Result

            If Balance_Produce < 0 Then Balance_Produce = 0
            If Balance_Result < 0 Then Balance_Result = 0
            If Actual_Result > 0 Then Actual_GoodPercentage = Me.ObjCD.DigitGrouping((Actual_Result / Actual_Produce) * 100, 2)
            If Actual_WorkHour > 0 Then Actual_ProdRate = Actual_Produce / (Actual_WorkHour / 60)

            Progress = Math.Round(Actual_Produce * 100 / Target_Produce, 2, MidpointRounding.AwayFromZero)
            If Progress > 100 Then
                Progress = 100
            End If
            Update_ForItem_Wo_Sch(Actual_Produce, Target_Produce, IdentifierWO, PrimaryWU, Actual_StartTime, Actual_EndTime, Actual_SetupTime, _
            Actual_WorkHour, Actual_TotalTime, Actual_ProdRate, Actual_Downtime, Actual_Produce, Actual_Result, Balance_Produce, Balance_Result, _
            Actual_GoodPercentage, Progress, vIdentifierJO)
            '------------------------------------------------------------------------------------------------------------------------------------------------

            '. untuk item_wo *
            Target_Produce = 0
            Target_Result = 0
            TargetProduceAdditional = 0
            Actual_StartTime = #1/1/1950#
            Actual_EndTime = #1/1/1950#
            Actual_SetupTime = 0
            Actual_WorkHour = 0
            Actual_TotalTime = 0
            Actual_ProdRate = 0
            Actual_Downtime = 0
            Actual_Produce = 0
            Actual_Result = 0
            Actual_GoodPercentage = 0
            Actual_ProjectResult = 0
            Balance_Produce = 0
            Balance_Result = 0
            Actual_TotalTime = 0
            Status = 0
            RejectAllowance = 0
            Actual_Reject = 0

            TblByWOItem.DefaultView.RowFilter = Nothing
            TblByWOItem.DefaultView.RowFilter = "[Identifier]='" & IdentifierWO & "'"
            If TblByWOItem.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TblByWOItem.DefaultView.Count - 1
                    rowByWOItem = TblByWOItem.DefaultView(ij).Row
                    With rowByWOItem
                        Target_Produce += .Qty_Target_Prod
                        Target_Result += .Qty_Target_Result
                        TargetProduceAdditional += .Qty_Target_Prod_Add
                    End With
                Next
            End If
            RejectAllowance = Target_Produce - Target_Result

            TblByBundling.DefaultView.RowFilter = Nothing
            For ij As Integer = 0 To TblByBundling.Rows.Count - 1
                rowByBundling = TblByBundling.Rows(ij)
                With rowByBundling
                    Actual_Produce += .total_produce
                    Actual_Result += .total_good
                    Actual_SetupTime += .setup_time_actual
                    Actual_Downtime += .down_time_actual
                    Actual_WorkHour += .work_hour_actual
                    If Actual_StartTime = #1/1/1950# Then
                        Actual_StartTime = .start_time_actual
                    Else
                        If Actual_StartTime > .start_time_actual Then
                            Actual_StartTime = .start_time_actual
                        End If
                    End If
                    If Actual_EndTime = #1/1/1950# Then
                        Actual_EndTime = .end_time_actual
                    Else
                        If Actual_EndTime < .end_time_actual Then
                            Actual_EndTime = .end_time_actual
                        End If
                    End If
                End With
            Next

            'Perhitungan balance MP: Misal
            'Qty Target 100
            'Qty Produce 110
            'Qty Allowance 10

            'contoh over produce
            'Produce I -> Good: 10, Reject 1 -> Balance MP: 90 <- 100-10-Reject, where reject = qty allowance - qty reject = 10 - 1 = 9 = 0 (jika positif, maka masih dalam allowance, set ke nol)
            'Produce II -> Good: 15, Reject 0 -> Balance MP: 75
            'Produce III -> Good: 80, Reject 1 -> Balance MP 0 <- 100-105+Reject, reject sama dengan atas, jika balance -, maka = 0
            Balance_MP = 0
            RejectAllowance = RejectAllowance - TotalReject
            If RejectAllowance > 0 Then
                RejectAllowance = 0
            End If
            Balance_MP = Target_Result - Actual_Result + RejectAllowance
            If Balance_MP < 0 Then
                Balance_MP = 0
            End If

            If Actual_WorkHour > 0 Then
                Actual_ProdRate = Actual_Produce / (Actual_WorkHour / 60)
            End If

            If Actual_Result > 0 Then
                Actual_GoodPercentage = (Actual_Result / Actual_Produce) * 100
            End If

            Actual_TotalTime = Actual_WorkHour + Actual_Downtime + Actual_SetupTime
            Balance_Produce = Target_Produce - Actual_Produce
            Balance_Result = Target_Result - Actual_Result

            If Balance_Produce < 0 Then Balance_Produce = 0
            If Balance_Result < 0 Then Balance_Result = 0
            If Actual_Result > 0 Then Actual_GoodPercentage = (Actual_Result / Actual_Produce) * 100
            If Actual_WorkHour > 0 Then Actual_ProdRate = Actual_Produce / (Actual_WorkHour / 60)

            Status = Me.SetStatusForItemWo(vIdentifierJO, IdentifierWO, Actual_Produce)
            Actual_ProjectResult = (Target_Produce + TargetProduceAdditional - Actual_Produce) * (Actual_GoodPercentage / 100) + Actual_Result
            If DtaUpdate_ForItemWO.UpdateQuery_ItemWo(Status, Actual_StartTime, Actual_EndTime, Actual_SetupTime, Actual_WorkHour, Actual_TotalTime, _
                   Actual_Downtime, Actual_ProdRate, Actual_Produce, Actual_Result, Actual_ProjectResult, Actual_GoodPercentage, _
                  Balance_Produce, Balance_Result, PrimaryWU, Balance_MP, IdentifierWO, vIdentifierJO) > 0 Then
            End If

            'update qty balance jo
            ObjMOItem.SetMOBalance(vIdentifierJO)

            'update qty rm + gp actual + qty wip + wip detail + rm detail
            Me.UpdateQtyBalanceRM(IdentifierWO, PrimaryWU, vIdentifierJO)

            'update wip detail balance + rm detail balance
            Me.Update_Detail_Balance(IdentifierWO, PrimaryWU, vIdentifierJO)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaByBundling.Dispose()
            DtaBySch.Dispose()
            DtaByWOItem.Dispose()
        End Try
    End Function

    Public Function SetStatusForItemWo(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vActualProduce As Decimal) As ClsTrans_ManufacturingOrder.EnumPilihStatusWO
        Dim Dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo2TableAdapter
        Dim Tbl As New DseTrans_StockBundling.tbltrans_mo_item_wo2DataTable
        Dim Row As DseTrans_StockBundling.tbltrans_mo_item_wo2Row
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.Fill(Tbl, vIdentifierJO, vIdentifierWO) > 0 Then
                Row = Tbl.Rows(0)
                If vActualProduce >= Row.Qty_Target_Prod Then
                    Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
                ElseIf vActualProduce > 0 And vActualProduce < Row.Qty_Target_Prod Then
                    Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress
                ElseIf vActualProduce <= 0 Then
                    Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
                End If
            End If
            Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
        End Try
    End Function

    Private Sub Update_Detail_Balance(ByVal vIdentifierWO As String, ByVal vPrimaryWU As Double, ByVal vIdentifierJO As String)
        Dim DTARM As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rm_detail1TableAdapter
        Dim TBLRM As New DseTrans_StockBundling.tbltrans_mo_item_rm_detail1DataTable
        Dim rowRM As DseTrans_StockBundling.tbltrans_mo_item_rm_detail1Row
        Dim DTAWIP As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wip_detail1TableAdapter
        Dim TBLWIP As New DseTrans_StockBundling.tbltrans_mo_item_wip_detail1DataTable
        Dim rowWIP As DseTrans_StockBundling.tbltrans_mo_item_wip_detail1Row
        Dim DTABundling As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before6TableAdapter
        Dim TBLBundling As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before6DataTable
        Dim rowBundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before6Row

        Try
            DTARM.Connection.ConnectionString = Core_Module.GetConstring
            DTAWIP.Connection.ConnectionString = Core_Module.GetConstring
            DTABundling.Connection.ConnectionString = Core_Module.GetConstring

            'load data rm detail
            DTARM.Fill(TBLRM, vIdentifierJO, vPrimaryWU, vIdentifierWO)
            'load data wip detail
            DTAWIP.Fill(TBLWIP, vIdentifierJO, vPrimaryWU, vIdentifierWO)
            'load data log produksi yang sudah terjadi
            DTABundling.FillBy1(TBLBundling, vPrimaryWU, vIdentifierWO, vIdentifierJO)

            'hitung balance untuk setiap item rm dan wip detail
            'ambil dari wip detail dulu -> lebih mudah, klo bundlingnya kemudian diedit
            For Each rowWIP In TBLWIP.Rows
                rowWIP.Used_Qty = 0
                rowWIP.Balance_Qty = rowWIP.Qty

                For ij As Integer = TBLBundling.Rows.Count - 1 To 0 Step -1
                    rowBundling = TBLBundling.Rows(ij)
                    If rowWIP.Primary_Product_Code_WIP = rowBundling.Primary_Product Then
                        If rowWIP.Balance_Qty <= rowBundling.Qty Then
                            rowBundling.Qty -= rowWIP.Balance_Qty
                            rowWIP.Used_Qty += rowWIP.Balance_Qty
                            rowWIP.Balance_Qty = 0
                        Else
                            rowWIP.Balance_Qty -= rowBundling.Qty
                            rowWIP.Used_Qty += rowBundling.Qty
                            rowBundling.Qty = 0
                        End If
                    End If
                    If rowBundling.Qty = 0 Then
                        TBLBundling.Rows.Remove(rowBundling)
                    End If
                Next
            Next
            'ambil dari rm detail
            For Each rowRM In TBLRM.Rows
                rowRM.Used_Qty = 0
                rowRM.Balance_Qty = rowRM.Target_Qty

                For ij As Integer = TBLBundling.Rows.Count - 1 To 0 Step -1
                    rowBundling = TBLBundling.Rows(ij)
                    If rowRM.Primary_Product_Code_RM = rowBundling.Primary_Product Then
                        If rowRM.Balance_Qty <= rowBundling.Qty Then
                            rowBundling.Qty -= rowRM.Balance_Qty
                            rowRM.Used_Qty += rowRM.Balance_Qty
                            rowRM.Balance_Qty = 0
                        Else
                            rowRM.Balance_Qty -= rowBundling.Qty
                            rowRM.Used_Qty += rowBundling.Qty
                            rowBundling.Qty = 0
                        End If
                    End If
                    If rowBundling.Qty = 0 Then
                        TBLBundling.Rows.Remove(rowBundling)
                    End If
                Next
            Next

            'update ke db
            For Each rowWIP In TBLWIP.Rows
                DTAWIP.UpdateQuery(rowWIP.Used_Qty, rowWIP.Balance_Qty, rowWIP.Primary)
            Next
            For Each rowRM In TBLRM.Rows
                DTARM.UpdateQuery(rowRM.Used_Qty, rowRM.Balance_Qty, rowRM.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini memiliki simpanan 
    Private Sub UpdateQtyBalanceRM(ByVal vIdentifierWO As String, ByVal PrimaryWU As Double, ByVal vIdentifierJO As String)
        Dim QtyUsed, QtyBalance As Decimal
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim PrimaryMO As Double = -1
        Dim DTA1 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before5TableAdapter
        Dim TBL1 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before5DataTable
        Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before5Row
        Dim DTA2 As New DseTrans_StockBundlingTableAdapters.tblgets_rmTableAdapter
        'Dim DTA3 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim GP_Actual, QtyRefProcess, QtyRM, QtyActualProduce, QtyUsedWIP, QtyBalanceWIP As Decimal
        Dim DTA4 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim TBL4 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Dim row4 As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim DTA5 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim Dummy As Nullable(Of Decimal)
        Dim a As Decimal = 0
        Dim DTARMDetail As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rm_detailTableAdapter
        Dim TBLRMDetail As New DseTrans_StockBundling.tbltrans_mo_item_rm_detailDataTable
        Dim rowRMDetail As DseTrans_StockBundling.tbltrans_mo_item_rm_detailRow
        Dim DTAWIPDetail As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wip_detailTableAdapter
        Dim TBLWIPDetail As New DseTrans_StockBundling.tbltrans_mo_item_wip_detailDataTable
        Dim rowWIPDetail As DseTrans_StockBundling.tbltrans_mo_item_wip_detailRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            'DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTARMDetail.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWIPDetail.Connection.ConnectionString = DTA.Connection.ConnectionString

            'kumpulkan item rm untuk kombinasi identifier jo + wo
            PrimaryMO = vIdentifierJO.Substring(0, vIdentifierJO.IndexOf("-"))
            DTA.FillBy_IdentifierJO_WO(TBL, PrimaryMO, vIdentifierJO, vIdentifierWO)
            DTA1.Fill(TBL1, vIdentifierWO, vIdentifierJO)
            DTA4.FillBy_Identifier_JO_WO(TBL4, PrimaryMO, vIdentifierJO, vIdentifierWO)
            DTARMDetail.Fill(TBLRMDetail, vIdentifierJO, vIdentifierWO, PrimaryWU)
            DTAWIPDetail.Fill(TBLWIPDetail, PrimaryWU, vIdentifierJO, vIdentifierWO)

            'kumpulkan item before untuk kombinasi identifier jo + wo
            For Each row In TBL.Rows
                'untuk setiap baris item rm
                'cari qty used
                QtyUsed = 0
                For Each row1 In TBL1.Rows
                    If row1.Primary_Product = row.Primary_Product_Code_RM Then
                        QtyUsed += row1.Qty
                    End If
                Next

                If row.is_wip = 1 Then
                    'set nilai balance-> khusus untuk item wip, asumsi menggunakan wip terlebih dahulu, sehingga  qty balance 
                    'rumusnya berbeda jika qty used < qty wip vs qty used => qty wip
                    If row.WIP_Qty - QtyUsed < 0 Then
                        'ada alocated rm yang digunakan
                        QtyBalance = row.Required_Qty + row.Additional_Qty - row.WIP_Qty + (row.WIP_Qty - QtyUsed)
                        'semua item wip sudah digunakan
                        QtyUsedWIP = row.WIP_Qty
                        QtyBalanceWIP = 0
                    Else
                        'allocated rm belum digunakan
                        QtyBalance = row.Required_Qty + row.Additional_Qty - row.WIP_Qty
                        'wip baru digunakan sebagian
                        QtyUsedWIP = QtyUsed
                        QtyBalanceWIP = row.WIP_Qty - QtyUsedWIP
                    End If
                    'update nilai qty used dan qty balance pada item wip
                    If row.WIP_Qty > 0 Then
                        DTAWIPDetail.UpdateQuery_SetQtyWIP(QtyUsedWIP, QtyBalanceWIP, vIdentifierJO, vIdentifierWO, row.Primary_Product_Code_RM)
                    End If
                Else
                    QtyBalance = row.Required_Qty + row.Additional_Qty - row.Replace_Qty - QtyUsed
                End If

                'set nilai gp aktual
                For Each row4 In TBL4.Rows
                    If row4.Primary_Product_RM = row.Primary_Product_Code_RM Then
                        QtyRefProcess = row4.Qty_Ref_Process
                        QtyRM = row4.Qty_RM
                    End If
                Next
                Dummy = DTA5.ScalarQuery_GetJml1(vIdentifierWO, vIdentifierJO)
                If Dummy.HasValue = True Then
                    QtyActualProduce = Dummy.Value
                Else
                    QtyActualProduce = 0
                End If

                If QtyUsed = 0 Then
                    GP_Actual = 0
                Else
                    If QtyRefProcess > 0 Then a = QtyRM / QtyRefProcess Else a = 0
                    GP_Actual = ObjCD.DigitGrouping(((a * QtyActualProduce) / QtyUsed) * 100, 2)
                End If

                'update ke database
                DTA2.UpdateQuery_RM(QtyUsed, QtyBalance, GP_Actual, vIdentifierWO, row.Primary_Product_Code_RM, vIdentifierJO)

                'rmdetail
                If row.is_wip = 0 Then
                    For Each rowRMDetail In TBLRMDetail.Rows
                        If rowRMDetail.Primary_Product_Code_RM = row.Primary_Product_Code_RM Then
                            QtyUsed = 0
                            For Each row1 In TBL1.Rows
                                If row1.Primary_Product = row.Primary_Product_Code_RM Then
                                    QtyUsed += row1.Qty
                                End If
                            Next
                            QtyBalance = rowRMDetail.Target_Qty - QtyUsed
                            DTARMDetail.UpdateQuery_SetQty(QtyUsed, QtyBalance, vIdentifierJO, vIdentifierWO, PrimaryWU, rowRMDetail.Primary_Product_Code_RM)
                        End If
                    Next
                Else
                    'wip detail 
                    For Each rowWIPDetail In TBLWIPDetail.Rows
                        If rowWIPDetail.Primary_Product_Code_WIP = row.Primary_Product_Code_RM Then
                            QtyUsed = 0
                            For Each row1 In TBL1.Rows
                                If row1.Primary_Product = row.Primary_Product_Code_RM Then
                                    QtyUsed += row1.Qty
                                End If
                            Next
                            QtyBalance = rowWIPDetail.Qty - QtyUsed
                            If QtyBalance < 0 Then
                                QtyBalance = 0
                                QtyUsed = rowWIPDetail.Qty
                            End If
                            DTAWIPDetail.UpdateQuery_SetQty(QtyUsed, QtyBalance, PrimaryWU, vIdentifierJO, vIdentifierWO, rowWIPDetail.Primary_Product_Code_WIP)
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
        End Try
    End Sub

    Public Function Update_ForItem_Wo_Sch(ByVal vactual_produce As Decimal, ByVal vtarget_produce As Decimal, ByVal vIdentifierWo As String, ByVal vPrimaryWorkUnit As Double, _
  ByVal vActualStartTime As Date, ByVal vActualEndTime As Date, ByVal vActualSetupTime As Double, ByVal vActualWorkDuration As Double, ByVal vActualTotalDuration As Double, ByVal vActualProdRate As Decimal, _
  ByVal vActualDownTime As Double, ByVal vActualProduce As Decimal, ByVal vActualResult As Decimal, ByVal vBalanceProduce As Decimal, ByVal vBalanceResult As Decimal, _
    ByVal vActualGoodPercentage As Decimal, ByVal vProgress As Decimal, ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim vStatus As New ClsTrans_ManufacturingOrder.EnumPilihStatusWO
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vStatus = Me.GetStatus(vactual_produce, vtarget_produce)
            DTA.UpdateQuery_WoItem_Sch(vStatus, vActualStartTime, vActualEndTime, vActualSetupTime, vActualWorkDuration, vActualTotalDuration, _
            vActualProdRate, vActualDownTime, vactual_produce, vActualResult, vBalanceProduce, vBalanceResult, vActualGoodPercentage, vProgress, _
            vIdentifierWo, vPrimaryWorkUnit, vIdentifierJO)

            'isitbl untuk menentukan status woitem dari data woitem_sch
            Me.IsikanTblSetStatus(vIdentifierWo, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanTblSetStatus(ByVal vIdentifierWo As String, ByVal vIdentifierJO As String) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo_item_sch1TableAdapter
        Dim rowStatusItemWo As DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch1Row
        Try
            tblset_status.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_IdentifierWo(tblStatusItemWo, vIdentifierWo, vIdentifierJO)
            If tblStatusItemWo.Rows.Count > 0 Then
                For Each rowStatusItemWo In tblStatusItemWo.Rows
                    rowset_status = tblset_status.Newtblset_statusRow
                    rowset_status.Work_Unit = rowStatusItemWo.primary_work_unit.ToString
                    If rowStatusItemWo.status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
                        rowset_status.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
                    ElseIf rowStatusItemWo.status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
                        rowset_status.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress
                    ElseIf rowStatusItemWo.status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
                        rowset_status.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
                    End If
                    tblset_status.Addtblset_statusRow(rowset_status)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SetStatusForItemWo() As ClsTrans_ManufacturingOrder.EnumPilihStatusWO
        Dim jmlBaris As Integer = 0
        Try
            jmlBaris = tblset_status.Rows.Count

            tblset_status.DefaultView.RowFilter = Nothing
            tblset_status.DefaultView.RowFilter = "[Status]=" & ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
            If tblset_status.DefaultView.Count = jmlBaris Then
                Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish
            End If

            tblset_status.DefaultView.RowFilter = Nothing
            tblset_status.DefaultView.RowFilter = "[Status]=" & ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            If tblset_status.DefaultView.Count = jmlBaris Then
                Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            End If

            Return ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahWorkTimeBundlingBenar_ByWu(ByVal vPrimaryWu As Double, ByVal vPrimaryTrans As Double, ByVal vStartTime As Date, _
   ByVal vEndTime As Date, ByRef ListNoBundlingStr As String, Optional ByVal vTipeBundling As Integer = 0) As Boolean
        Dim ListNoBundling As New List(Of String)
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling6TableAdapter
        Dim TBL As New DseTrans_StockBundling.tbltransaksi_stockbundling6DataTable
        Dim Row As DseTrans_StockBundling.tbltransaksi_stockbundling6Row

        Try
            ListNoBundling.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL, vPrimaryWu, vTipeBundling) > 0 Then
                TBL.DefaultView.RowFilter = Nothing
                TBL.DefaultView.RowFilter = "[Primary]<> " & vPrimaryTrans & " AND [primary work unit]= " & vPrimaryWu & "AND ([end time actual] >= #" & _
                vStartTime & "# AND [start time actual] <= #" & vEndTime & "#)"
                If TBL.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBL.DefaultView.Count - 1
                        Row = TBL.DefaultView(ij).Row
                        If vStartTime = Row.end_time_actual Or vEndTime = Row.start_time_actual Then
                            'do nothing
                        Else
                            If ListNoBundling.Contains(Row.No_Transaksi) = False Then
                                ListNoBundling.Add(Row.No_Transaksi)
                            End If
                        End If
                    Next
                End If
            End If
            If ListNoBundling.Count > 0 Then
                Dim Nomber As Integer = 1
                For Each ij As String In ListNoBundling
                    If Nomber <= 5 Then
                        If ListNoBundlingStr = "" Then
                            ListNoBundlingStr = Nomber & ". " & ij
                            Nomber += 1
                        Else
                            ListNoBundlingStr &= vbNewLine & Nomber & ". " & ij
                            Nomber += 1
                        End If
                    Else
                        Exit For
                    End If
                Next
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDefaultTime_SetupTime(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_setuptimeTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDefaultSetupTime(vPrimary)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByIdentifierWO_PrimaryWU(ByVal vIdentifierWO As String, ByVal vPrimaryWU As Double) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByWOWU(vIdentifierWO, vPrimaryWU)
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

    Public Function GetJml_ByIdentifierJO(ByVal vIdentifierJO As String) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByIdentifierJO(vIdentifierJO)
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

    Public Sub GetInputAndOutputLocation(ByRef vPrimaryInputLocation As Double, ByRef vInputLocation As String, ByRef vInpuLocationDesc As String, ByRef vPrimaryOutputLocation As Double, ByRef vOutputLocation As String, _
    ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkunit As Double)
        Dim dta As New DseTrans_StockBundlingTableAdapters.DataTable8TableAdapter
        Dim tbl As New DseTrans_StockBundling.DataTable8DataTable
        Dim row As DseTrans_StockBundling.DataTable8Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, vIdentifierJO, vIdentifierWO, vPrimaryWorkunit) > 0 Then
                row = tbl.Rows(0)
                vPrimaryInputLocation = row.input_location_primary
                vPrimaryOutputLocation = row.output_location_primary
                vInputLocation = row.input_location_code
                vInpuLocationDesc = row.input_location_desc
                vOutputLocation = row.output_location_code
            Else
                vPrimaryInputLocation = -1
                vPrimaryOutputLocation = -1
                vInputLocation = String.Empty
                vInpuLocationDesc = String.Empty
                vOutputLocation = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJml_ByJO_WO_PrimaryWorkUnit(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkUnit As Double) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling5TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByJO_WO_PrimaryWorkUnit(vIdentifierJO, vIdentifierWO, vPrimaryWorkUnit)
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
End Class
