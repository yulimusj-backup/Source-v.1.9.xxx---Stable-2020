Option Explicit On
Option Strict Off

Public Class clsTransaksiStockTransfer
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    'transit location
    Private mUseTransit As Integer = 0
    Private mTransitStatus As Integer = 0
    Private mReceivedBy As String = ""
    Private mReceivedDate As Date = Now
    Private mReceivedTimestamp As Date = Now
    Private mIsReceivement As Integer = 0
    Private mPrimaryRefDoc As Double = -1

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

    Public Enum EnumTransitStatus As Integer
        No_Transit
        On_Transit
        Finished_Transit
    End Enum

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

    Public Property UseTransit() As Integer
        Get
            Return mUseTransit
        End Get
        Set(ByVal value As Integer)
            mUseTransit = value
        End Set
    End Property

    Public Property TransitStatus() As Integer
        Get
            Return mTransitStatus
        End Get
        Set(ByVal value As Integer)
            mTransitStatus = value
        End Set
    End Property

    Public Property ReceivedBy() As String
        Get
            Return mReceivedBy
        End Get
        Set(ByVal value As String)
            mReceivedBy = value
        End Set
    End Property

    Public Property ReceivedDate() As Date
        Get
            Return mReceivedDate
        End Get
        Set(ByVal value As Date)
            mReceivedDate = value
        End Set
    End Property

    Public Property ReceivedTimestamp() As Date
        Get
            Return mReceivedTimestamp
        End Get
        Set(ByVal value As Date)
            mReceivedTimestamp = value
        End Set
    End Property

    Public Property IsReceivement() As Integer
        Get
            Return mIsReceivement
        End Get
        Set(ByVal value As Integer)
            mIsReceivement = value
        End Set
    End Property

    Public Property PrimaryRefDoc() As Double
        Get
            Return mPrimaryRefDoc
        End Get
        Set(ByVal value As Double)
            mPrimaryRefDoc = value
        End Set
    End Property

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrut(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrut(TanggalEnd, TanggalStart).Value + 1)
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
            Return NoUrut + "/" + "TRF" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter

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

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim Retval As String = ""

        Try
            Retval = DTA.ScalarQuery_GetNoTransaksiByPrimary(vPrimary)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim Retval As String = ""

        Try
            Retval = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GeTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim Retval As New Nullable(Of Date)

        Try
            Retval = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Retval.HasValue = True Then
                Return Retval
            End If

            Return "1/1/1950"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1/1/1950"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vApakahTransaksiReceivement As Boolean, _
    ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable, _
    Optional ByVal vFromTransaksiReceive As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim TBL As New dseTransaksiStockTransfer.tbltransaksi_stocktransferDataTable
        Dim row As dseTransaksiStockTransfer.tbltransaksi_stocktransferRow
        Dim ObjItem As New clsTransaksiStockTransferItem
        Dim TransitStatus_Enum As EnumTransitStatus
        Dim ObjCustTrf As New CustomSMESS.clsTransaksiStockTransfer

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi

                    Me.mUseTransit = .Use_Transit
                    If .Transit_Status = 0 Then
                        Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
                        TransitStatus_Enum = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
                    ElseIf .Transit_Status = 1 Then
                        Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.On_Transit
                        TransitStatus_Enum = clsTransaksiStockTransfer.EnumTransitStatus.On_Transit
                    ElseIf .Transit_Status = 2 Then
                        Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit
                        TransitStatus_Enum = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit
                    End If
                    Me.mIsReceivement = .Is_Receivement
                    Me.mPrimaryRefDoc = .Primary_Ref_Doc
                    Me.mReceivedBy = .Received_By
                    Me.mReceivedDate = .Received_Date
                    Me.mReceivedTimestamp = .Received_Timestamp

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                'If vFromTransaksiReceive = False Then
                '    If Me.mIsReceivement = 1 Then
                '        vApakahDariTransaksiReceivement = False
                '    End If
                'End If

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, vFromTransaksiReceive, TransitStatus_Enum)
                ObjCustTrf.GetStockTransferItem_Custom(vTblItemCustom, vPrimaryTransaksi)
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

    Public Overridable Function ApakahBolehEdit(ByVal vUser As String, ByVal vStrictEdit As Boolean, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim DTA As New dseTransaksiStockTransferTableAdapters.DataTable3TableAdapter
        Dim TBL As New dseTransaksiStockTransfer.DataTable3DataTable
        Dim row As dseTransaksiStockTransfer.DataTable3Row
        Dim ProductCode As String = ""
        Dim ObjCustom As New CustomSMESS.clsTransaksiStockTransfer

        Try
            'Terkait Transit
            If Me.mIsReceivement = 1 Then
                If TampilkanPesan = True Then
                    MsgBox("Please use form 'Location Transfer - Set Status' to edit data in this transaction.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                If TampilkanPesan = True Then
                    MsgBox("This transaction has been received, you can't edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'msgbox("belum")
            'Suatu transaksi in/out tidak bisa diedit jika:
            '1. jika ada product dalam stock transfer ini yang mengalami perubahan kontrol
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

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. Jika strict edit nyala maka hanya user yang bersangkutan yang boleh melakukan proses edit data
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be edit by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            If ObjCustom.ApakahBolehEdit_Custom(Me.mPrimary, TampilkanPesan) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter

        Try
            If DTA.ScalarQuery_GetTanggal(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggal(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTransitStatus(ByVal vPrimary As Double) As EnumTransitStatus
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Retval = dta.ScalarQuery_GetTransitStatus(vPrimary)
            If Retval.HasValue = True Then
                If Retval.Value = 0 Then
                    Return EnumTransitStatus.No_Transit
                ElseIf Retval.Value = 1 Then
                    Return EnumTransitStatus.On_Transit
                Else
                    Return EnumTransitStatus.Finished_Transit
                End If
            Else
                Return EnumTransitStatus.No_Transit
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahDibuatReceive(ByVal vPrimaryRefDoc As Double, ByVal vPrimaryStockTransfer As Double) As Boolean
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Retval = dta.ScalarQuery_GetJmlReceived(vPrimaryRefDoc, vPrimaryStockTransfer)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, ByVal vStrictEdit As Boolean, _
    ByRef vTblProdQty As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable, _
    ByRef vListProduct As List(Of Double), Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim DTA As New dseTransaksiStockTransferTableAdapters.DataTable3TableAdapter
        Dim TBL As New dseTransaksiStockTransfer.DataTable3DataTable
        Dim row As dseTransaksiStockTransfer.DataTable3Row
        Dim ProductCode As String = ""
        Dim ObjCustom As New CustomSMESS.clsTransaksiStockTransfer
        Try
            'Terkait Transit
            If Me.mIsReceivement = 1 Then
                If TampilkanPesan = True Then
                    MsgBox("Please use form 'Location Transfer - Set Status' to delete data in this transaction.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Me.mTransitStatus = Me.GetTransitStatus(Me.mPrimary)
            If Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                If TampilkanPesan = True Then
                    MsgBox("This transaction has been received, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'msgbox("belum")
            'transaksi tidak bisa dihapus bila:
            '1. jika ada product dalam stock transfer ini yang mengalami perubahan kontrol
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

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. jika menghapus transaksi ini akan mengakibatkan data product menjadi minus
            Dim ObjItem As New clsTransaksiStockTransferItem
            If ObjItem.ApakahHapusMengakibatkanMinus(Me.mPrimary) = True Then
                ObjItem = Nothing
                If TampilkanPesan = True Then
                    MsgBox("Deleting this data will cause minus in stock quantity, process cancelled.", MsgBoxStyle.Information, _
                    "Information")
                End If

                Return False
            End If
            ObjItem = Nothing

            '4. Jika strict edit nyala maka hanya user yang bersangkutan yang boleh melakukan proses edit data
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be delete by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            If ObjCustom.ApakahBolehHapus_Custom(Me.mPrimary, False, vTblProdQty, vListProduct, TampilkanPesan) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCL = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimary(vNoTransaksi)
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDgvItem As DataGridView, _
   ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable, _
   ByVal vtblProdQty As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable, ByVal vListProductCustom As List(Of Double), _
   Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim Obj1 As New Core_Module.clsMasterListedSoftware
        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
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

            '3. Untuk simpan perubahan, simulasikan terlebih dahulu data setelah simpan, 
            '   pastikan tidak ada data yang negatif
            If Me.CekBatasanMinusQty(Me.mPrimary, vDgvItem, True) = False Then
                Return False
            End If

            'Terkait Transit
            If Me.mIsReceivement = 1 Then
                'jika receivement pastikan lagi jika document reference belum di hapus
                Dim NoTransaksiReference As String = Me.GetNoTransaksi(Me.mPrimaryRefDoc)
                If String.IsNullOrEmpty(NoTransaksiReference) = True Then
                    'jika document reference tidak ada maka tidak bisa simpan transaksi
                    If TampilkanPesan = True Then
                        MsgBox("Other user has been deleted transaction reference", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            'Terkait Transit
            If Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                If TampilkanPesan = True Then
                    MsgBox("This transaction can't be save because has been finished transit", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '5. tidak bisa receive jika sudah dibuat transaksi receive (khusus transaksi receivement)
            If Me.mIsReceivement = 1 Then
                If Me.ApakahSudahDibuatReceive(Me.mPrimaryRefDoc, Me.mPrimary) = True Then
                    Dim NoRefDoc As String = Me.GetNoTransaksi(Me.mPrimaryRefDoc)
                    MsgBox("Transaction stock transfer: " & NoRefDoc.ToString & " has been received", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim ObjCustTrf As New CustomSMESS.clsTransaksiStockTransfer
            If ObjCustTrf.ApakahDataBenar_Custom(Me.mPrimary, vBaru, vtblProdQty, vTblItemCustom, vListProductCustom) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj1 = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_StockTransfer_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vTblItemCustom)
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

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vKoreksiStock As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim ObjItem As New clsTransaksiStockTransferItem

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustTrf As New CustomSMESS.clsTransaksiStockTransfer
        Dim ListProductCustom As New List(Of Double)

        Try
            If Me.ApakahDataBenar(True, vdgvItem, vTblItemCustom, tblProdQty, ListProductCustom, True, vCekDataBenar) = True Then
                'insert query
                If DTA.Insert(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, _
                 Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, _
                 Me.mUseTransit, Me.mTransitStatus, Me.mPrimaryRefDoc, Me.mIsReceivement) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    If vFromCutOff = False Then
                        'cek berkaitan nomor sama untuk beda bulan
                        Me.CekNomor(DTA)
                    End If

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                    Me.mNoTransaksi, "", "", vKoreksiStock)

                    'jika simpan transaksi stock transfer receivement maka update status received pada transaksi reference documentnya
                    If Me.mIsReceivement = 1 And Me.mPrimaryRefDoc <> -1 Then
                        Me.UpdateStatusReceived(vUserCode, Me.mCreatedDate, Me.mTimestamp, Me.mPrimaryRefDoc, EnumTransitStatus.Finished_Transit)
                    End If

                    ''################################CUSTOM INOUT#############################################
                    ObjCustTrf.Simpan_Custom(Me.mPrimary, vTblItemCustom, tblProdQty, ListProductCustom)
                    ''################################CUSTOM INOUT#############################################

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vdgvItem, vUserCode, vTblItemCustom, vCekDataBenar, vKoreksiStock, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function UpdateStatusReceived(ByVal vUserCode As String, ByVal vReceivedDate As Date, _
    ByVal vReceivedTimestamp As Date, ByVal vPrimaryRefDoc As Double, ByVal vStatusTransit As EnumTransitStatus) As Boolean
        Dim dta As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_TransitStatus(vStatusTransit, vUserCode, vReceivedDate, vReceivedTimestamp, vPrimaryRefDoc) > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mNoUrut
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari)
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
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

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_StockTransfer_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, vTblItemCustom)
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

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vTblItemCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim ObjItem As New clsTransaksiStockTransferItem

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustTrf As New CustomSMESS.clsTransaksiStockTransfer
        Dim ListProductCustom As New List(Of Double) 

        Try
            If Me.ApakahDataBenar(False, dgvItem, vTblItemCustom, tblProdQty, ListProductCustom) = True Then
                If DTA.Update(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, _
                    Me.mRemarks, vUserCode, Now, Now, Me.mUseTransit, Me.mTransitStatus, Me.mIsReceivement, _
                    Me.mPrimaryRefDoc, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, Me.mNoTransaksi)
                    'nilai balikan

                    'jika simpan transaksi stock transfer receivement maka update status received pada transaksi reference documentnya
                    If Me.mIsReceivement = 1 And Me.mPrimaryRefDoc <> -1 Then
                        Me.UpdateStatusReceived(vUserCode, Me.mCreatedDate, Me.mTimestamp, Me.mPrimaryRefDoc, EnumTransitStatus.Finished_Transit)
                    End If

                    ''################################CUSTOM INOUT#############################################
                    ObjCustTrf.Hapus_Custom(Me.mPrimary, tblProdQty, ListProductCustom, False)
                    ObjCustTrf.Simpan_Custom(Me.mPrimary, vTblItemCustom, tblProdQty, ListProductCustom)
                    ObjCustTrf.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)
                    ''################################CUSTOM INOUT#############################################

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan1(dgvItem, vUserCode, vTblItemCustom)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            tblProdQty.Dispose()
            ObjCustTrf = Nothing
            ListProductCustom = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_Hapus, vUserCode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, vStrictEdit)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean) As Boolean
        Dim Obj As New clsTransaksiStockTransferItem
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustTrf As New CustomSMESS.clsTransaksiStockTransfer
        Dim ListProductCustom As New List(Of Double)

        Try
            If Me.ApakahBolehHapus(vUserCode, vStrictEdit, tblProdQty, ListProductCustom) = True Then
                'hapus data item 
                Obj.Hapus(vPrimaryTransaksi)
                'hapus data header, data item otomatis dihapuskan oleh server
                ObjCustTrf.Hapus_Custom(Me.mPrimary, tblProdQty, ListProductCustom, True)
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

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransferTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggal(vTanggalSampai, vTanggalDari).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%").Value
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

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseTransaksiStockTransfer.tbltransaksi_stocktransfer1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_Tanggal(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Transaction No"
                DGV.Columns(2).HeaderText = "Transaction Date"
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            Me.ResetObject()
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ResetObject()
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            Me.mUseTransit = 0
            Me.mTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
            Me.mReceivedBy = ""
            Me.mReceivedDate = Now
            Me.mReceivedTimestamp = Now
            Me.mIsReceivement = 0
            Me.mPrimaryRefDoc = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Decimal
        Dim dta As New dseTransaksiStockTransferTableAdapters.DataTable4TableAdapter
        Dim val As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            val = dta.ScalarQuery_GetQty(vPrimaryProduct, vPrimaryLocation, vBatchNo)
            If val.HasValue Then Return val
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetQtyDB(ByVal vFrom As Boolean, ByVal vPrimaryTrans As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Decimal
        Dim dta As New dseTransaksiStockTransferTableAdapters.DataTable4TableAdapter
        Dim val As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vFrom Then
                val = dta.ScalarQuery_GetQtyFrom(vPrimaryTrans, vPrimaryProduct, vPrimaryLocation, vBatchNo)
            Else
                val = dta.ScalarQuery_GetQtyTo(vPrimaryTrans, vPrimaryProduct, vPrimaryLocation, vBatchNo)
            End If
            If val.HasValue Then Return val
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'add aul batasan minus stock transfer (vCalculateDgv = untuk menentukan ApakahDataBenar/ApakahBolehHapus karena ApakahBolehHapus tidak memperhitungkan qty dgv)
    Public Function CekBatasanMinusQty(ByVal vPrimaryTransaksi As Double, ByVal vDgvItem As DataGridView, ByVal vCalculateDgv As Boolean) As Boolean
        'product qty
        Dim dtaStock As New dseTransaksiStockTransferTableAdapters.tblmaster_productqtyTableAdapter
        Dim tblStock As New dseTransaksiStockTransfer.tblmaster_productqtyDataTable
        Dim rowStock As dseTransaksiStockTransfer.tblmaster_productqtyRow
        'qty db (list item product stock transfer)
        Dim dtaDB As New dseTransaksiStockTransferTableAdapters.tbl_qty_dbTableAdapter
        Dim tblDB As New dseTransaksiStockTransfer.tbl_qty_dbDataTable
        Dim rowDB As dseTransaksiStockTransfer.tbl_qty_dbRow
        'qty dgv (list item product stock transfer)
        Dim tblDgv, tblDgvSort As New dseTransaksiStockTransfer.tbl_qty_dgvDataTable
        Dim rowDgv As dseTransaksiStockTransfer.tbl_qty_dgvRow
        'qty minus dummy
        Dim TblDum, TblMinus As New dseTransaksiStockTransfer.tbl_batasan_minusDataTable
        Dim rowDum, rowMinus As dseTransaksiStockTransfer.tbl_batasan_minusRow
        Dim DefPrec As Integer = 0
        Dim Frm As New FrmTransaksiStockTransfer
        Dim ListPrimaryProduct As New List(Of Double)
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim UnitCode As String = "'"

        Try
            dtaDB.Connection.ConnectionString = Core_Module.GetConstring
            dtaStock.Connection.ConnectionString = Core_Module.GetConstring

            dtaDB.Fill(tblDB, vPrimaryTransaksi)
            'get primary product (qty db)
            For Each rowDB In tblDB
                If ListPrimaryProduct.Contains(rowDB.Primary_Product) = False Then
                    ListPrimaryProduct.Add(rowDB.Primary_Product)
                End If
            Next

            If vCalculateDgv = True Then
                For Each rd As DataGridViewRow In vDgvItem.Rows
                    If String.IsNullOrEmpty(rd.Cells(Frm.ColPrimaryCodeNo.Index).Value) = True Then
                        rd.Cells(Frm.ColPrimaryCodeNo.Index).Value = -1
                    End If
                    If String.IsNullOrEmpty(rd.Cells(Frm.ColBatch.Index).Value) = True Then
                        rd.Cells(Frm.ColBatch.Index).Value = ""
                    End If
                    If rd.Cells(Frm.ColPrimaryCodeNo.Index).Value <> -1 Then
                        If ListPrimaryProduct.Contains(CDbl(rd.Cells(Frm.ColPrimaryCodeNo.Index).Value)) = False Then
                            ListPrimaryProduct.Add(CDbl(rd.Cells(Frm.ColPrimaryCodeNo.Index).Value))
                        End If

                        'add ke tbl_qty_dgv
                        rowDgv = tblDgv.Newtbl_qty_dgvRow
                        rowDgv.primary_product = CDbl(rd.Cells(Frm.ColPrimaryCodeNo.Index).Value)
                        rowDgv.primary_location_asal = CDbl(rd.Cells(Frm.ColPrimaryLocation.Index).Value)
                        rowDgv.primary_location_tujuan = CDbl(rd.Cells(Frm.ColPrimaryDestination.Index).Value)
                        rowDgv.batch = rd.Cells(Frm.ColBatch.Index).Value
                        rowDgv.qty = CDec(rd.Cells(Frm.ColQty.Index).Value)
                        rowDgv.Product_Code = rd.Cells(Frm.ColCodeNo.Index).Value
                        rowDgv.Location_Asal = rd.Cells(Frm.ColLocation.Index).Value
                        rowDgv.Location_Tujuan = rd.Cells(Frm.ColDestinationLoc.Index).Value
                        rowDgv._Decimal = CInt(Me.ObjCD.GetJmlPrecision(rd.Cells(Frm.ColQty.Index).Value))
                        tblDgv.Rows.Add(rowDgv)
                    End If
                Next

                If tblDgv.Count > 0 Then
                    tblDgv.DefaultView.RowFilter = ""
                    tblDgv.DefaultView.Sort = "[primary product], [primary location asal], [primary location tujuan], [batch]"
                    tblDgvSort.Merge(tblDgv.DefaultView.ToTable)
                End If
            Else
                tblDgvSort.Rows.Clear()
            End If

            If ListPrimaryProduct.Count > 0 Then
                ListPrimaryProduct.Sort()
                If ListPrimaryProduct.Count > 1 Then
                    If ListPrimaryProduct.Item(0) > ListPrimaryProduct.Item(1) Then
                        ListPrimaryProduct.Sort()
                    End If
                End If

                For Each PrimaryProduct As Double In ListPrimaryProduct
                    'clear tbl dummy per looping product
                    TblDum.Rows.Clear()
                    DefPrec = ObjProd.GetAngkaDecimal(PrimaryProduct)
                    UnitCode = ObjProd.GetUnitCode(PrimaryProduct)

                    'qty dgv (jika tidak calculate dgv harusnya tidak masuk looping karena tblDgvSort nya kosong)
                    For Each rowDgv In tblDgvSort
                        With rowDgv
                            If .primary_product = PrimaryProduct Then
                                'qty dgv jika lokasi asal = Kurang (-)
                                Me.CekBatasanMinusQty_AddTblDum(TblDum, .primary_product, .primary_location_asal, .batch, .qty, .Product_Code, .Location_Asal, False, DefPrec, UnitCode)
                                'qty dgv jika lokasi tujuan = Tambah (+)
                                Me.CekBatasanMinusQty_AddTblDum(TblDum, .primary_product, .primary_location_tujuan, .batch, .qty, .Product_Code, .Location_Tujuan, True, DefPrec, UnitCode)
                            End If
                        End With
                    Next

                    'qty db
                    For Each rowDB In tblDB
                        With rowDB
                            If .Primary_Product = PrimaryProduct Then
                                'qty db jika lokasi asal = Tambah (+)
                                Me.CekBatasanMinusQty_AddTblDum(TblDum, .Primary_Product, .Primary_Storage_Area, .Batch_No, .Qty, .Product_Code, .Location_Asal, True, DefPrec, UnitCode)
                                'qty db jika lokasi tujuan = Kurang (-)
                                Me.CekBatasanMinusQty_AddTblDum(TblDum, .Primary_Product, .primary_destination_area, .Batch_No, .Qty, .Product_Code, .Location_Tujuan, False, DefPrec, UnitCode)
                            End If
                        End With
                    Next

                    'qty stock
                    TblDum.DefaultView.RowFilter = ""
                    TblDum.DefaultView.Sort = "[primary product], [primary location], [batch]"

                    tblStock.Rows.Clear()
                    dtaStock.Fill(tblStock, PrimaryProduct)
                    tblStock.DefaultView.RowFilter = ""
                    tblStock.DefaultView.Sort = "[Primary Product] DESC, [Primary Storage Area] DESC, [Lot No] DESC"
                    For Each dv As Data.DataRowView In TblDum.DefaultView
                        rowDum = dv.Row
                        For i As Integer = tblStock.DefaultView.Count - 1 To 0 Step -1
                            rowStock = tblStock.DefaultView(i).Row
                            'cocokan berdasarkan 1.primary product 2.primary location 3.batch
                            If rowDum.primary_product = rowStock.Primary_Product And rowDum.primary_location = rowStock.Primary_Storage_Area And rowDum.batch = rowStock.Lot_No Then
                                rowDum.qty += rowStock.Qty 'tambahkan dengan qty stock

                                Exit For
                            Else
                                'tblStock.Rows.Remove(rowStock)
                            End If
                        Next
                    Next

                    For Each rowDum In TblDum
                        'jika minus maka masukan ke tbl minus
                        If rowDum.qty < 0 Then
                            rowMinus = TblMinus.Newtbl_batasan_minusRow
                            For a1 As Integer = 0 To TblDum.Columns.Count - 1
                                rowMinus.Item(a1) = rowDum.Item(a1)
                            Next
                            TblMinus.Rows.Add(rowMinus)
                        End If
                    Next
                Next
            End If


            If TblMinus.Count > 0 Then
                Dim FrmMinus As New FrmTransaksiStockTransfer_BatasanMinus
                FrmMinus.TblMinus = TblMinus
                FrmMinus.ShowDialog()
                FrmMinus.Dispose()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaDB.Dispose()
            tblDB.Dispose()
            dtaStock.Dispose()
            tblStock.Dispose()
            TblDum.Dispose()
            TblMinus.Dispose()
            Frm.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    Private Sub CekBatasanMinusQty_AddTblDum(ByRef TblDum As dseTransaksiStockTransfer.tbl_batasan_minusDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatch As String, ByVal vQty As Decimal, _
    ByVal vProductCode As String, ByVal vLocation As String, ByVal vAdd As Boolean, ByVal vDefPrec As Integer, _
    ByVal vUnitCode As String)
        Dim rowDum As dseTransaksiStockTransfer.tbl_batasan_minusRow
        Dim Ketemu As Boolean = False

        Try
            For Each rowDum In TblDum
                If vPrimaryProduct = rowDum.primary_product And vPrimaryLocation = rowDum.primary_location And vBatch = rowDum.batch Then
                    Ketemu = True 'jika ketemu maka calculate qty
                    If vAdd = True Then
                        'add
                        rowDum.qty += vQty
                    Else
                        'reduce
                        rowDum.qty -= vQty
                    End If
                End If
            Next

            If Ketemu = False Then 'jika tidak ketemu maka tambah baris baru
                rowDum = TblDum.Newtbl_batasan_minusRow
                rowDum.primary_product = vPrimaryProduct
                rowDum.primary_location = vPrimaryLocation
                rowDum.batch = vBatch
                If vAdd = True Then
                    'add
                    rowDum.qty += vQty
                Else
                    'reduce
                    rowDum.qty -= vQty
                End If
                rowDum.Product_Code = vProductCode
                rowDum.Location = vLocation
                rowDum._Decimal = vDefPrec
                rowDum.Unit_Code = vUnitCode
                TblDum.Rows.Add(rowDum)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
