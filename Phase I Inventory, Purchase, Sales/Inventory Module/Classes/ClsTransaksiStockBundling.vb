Option Explicit On
Option Strict Off

Public Class ClsTransaksiStockBundling
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    Private mPrimaryProses As Double = -1
    Private mPrimaryWorkUnit As Double = -1
    Private mPrimaryWorkShift As Double = -1

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

    Public Enum EnumTipeBundling As Integer
        Inventory
        Purchase1
        Produksi1
    End Enum

    Public Property TipeBundling() As EnumTipeBundling
        Get
            Return Me.mTipeBundling
        End Get
        Set(ByVal value As EnumTipeBundling)
            Me.mTipeBundling = value
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

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vTipeBundling As EnumTipeBundling) As String
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DummyVal As Nullable(Of Integer)

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

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

            'pastikan no urut terdiri atas 5 huruf
            While NoUrut.Length < 5
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/" + "BDL" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItemB4 As DataGridView, ByVal vDGVItemAfter As DataGridView, ByVal vFrm As FrmTransaksiStockBundling) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBL As New dseTransaksiStockBundling.tbltransaksi_stockbundlingDataTable
        Dim row As dseTransaksiStockBundling.tbltransaksi_stockbundlingRow
        Dim ObjItem As New clsTransaksiStockBundlingItem
        Dim ObjCustBundling As New CustomSMESS.ClsTransaksiStockBundling

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryProses = .primary_process
                    Me.mPrimaryWorkUnit = .primary_work_unit
                    Me.mPrimaryWorkShift = .primary_workshift
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItemB4, vDGVItemAfter, Me.mPrimary)
                ObjCustBundling.GetStockBundlingItem_Custom(vFrm.TblItemBefCustom, vFrm.TblItemAftCustom, vPrimaryTransaksi)

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
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.DataTable4TableAdapter
        Dim TBL As New dseTransaksiStockBundling.DataTable4DataTable
        Dim row As dseTransaksiStockBundling.DataTable4Row
        Dim DTA1 As New dseTransaksiStockBundlingTableAdapters.DataTable5TableAdapter
        Dim TBL1 As New dseTransaksiStockBundling.DataTable5DataTable
        Dim row1 As dseTransaksiStockBundling.DataTable5Row
        Dim ProductCode As String = ""
        Dim DTA2 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockBundling
        Try
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
                ElseIf DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) <> vUser And vUser <> "USER" Then
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
            DTA1.Dispose()
            TBL.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, ByVal vBaru As Boolean, _
   ByVal vTblProdQty As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable, _
   ByVal vListProductCustom As List(Of Double), Optional ByVal TampilkanPesan As Boolean = True, _
   Optional ByVal vStrictEdit As Boolean = True) As Boolean
        Dim TanggalStockInOut As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.DataTable4TableAdapter
        Dim TBL As New dseTransaksiStockBundling.DataTable4DataTable
        Dim row As dseTransaksiStockBundling.DataTable4Row
        Dim DTA1 As New dseTransaksiStockBundlingTableAdapters.DataTable5TableAdapter
        Dim TBL1 As New dseTransaksiStockBundling.DataTable5DataTable
        Dim row1 As dseTransaksiStockBundling.DataTable5Row
        Dim ProductCode As String = ""
        Dim DTA2 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockBundling
        Try
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
                ElseIf DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary) <> vUser And vUser <> "USER" Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be deleted by: " + DTA2.ScalarQuery_GetCreatedBy(Me.mPrimary), MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            '3. jika menghapus transaksi ini akan mengakibatkan data product menjadi minus
            Dim ObjItem As New clsTransaksiStockBundlingItem
            If ObjItem.ApakahHapusMengakibatkanMinus(Me.mPrimary) = True Then
                ObjItem = Nothing
                Return False
            End If
            ObjItem = Nothing

            '4. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalStockInOut = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalStockInOut Then
                MsgBox("The period " + Format(TanggalStockInOut, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. 
            If ObjCustom.ApakahBolehHapus_Custom(Me.mPrimary, vBaru, vTblProdQty, vListProductCustom, TampilkanPesan) = False Then
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
            TBL1.Dispose()
            TBL.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vFrm As FrmTransaksiStockBundling, ByVal vDGVItemAfter As DataGridView, _
    ByVal vDGVItemBefore As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByRef vTblProdQty As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable, _
       ByRef vListProduct As List(Of Double), Optional ByVal TampilkanPesan As Boolean = True, _
    Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim ListKombinasi As New List(Of String)
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAfterLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow

        Dim DTA1 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBL1 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBeforeLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
        Dim TBLMasterProductQty_Dummy As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
        Dim rowQty, rowQty_Dummy, rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim Expiry1 As Date = #1/1/1950#
        Dim Dummy As String = ""

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            'DTAMasterProductQty.Connection.ConnectionString = DTA.Connection.ConnectionString

            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                    End If
                End If
            End If

            '2. karena after dan before, bisa memiliki barang yang persis sama, mungkin hanya beda qty saja
            'primary product, batch number, expiry, location bisa sama
            'maka simulasi after dan beforenya digabungkan.
            'load item lama after dan before
            TBLMasterProductQty.Rows.Clear()
            If vPrimaryTransaksi <> -1 Then
                DTA.Fill(TBL, vPrimaryTransaksi)
                DTA1.Fill(TBL1, vPrimaryTransaksi)
            End If

            'load tabel master product qty
            '1. row after lama
            '2. row before lama
            '3. dgv after
            '4. dgv before
            For Each rowAfterLama In TBL.Rows
                If ListKombinasi.Contains(rowAfterLama.Primary_Product.ToString + "-" + _
                rowAfterLama.Primary_Storage_Area.ToString + "-" + _
                rowAfterLama.Batch_No) = False Then
                    ListKombinasi.Add(rowAfterLama.Primary_Product.ToString + "-" + _
                    rowAfterLama.Primary_Storage_Area.ToString + "-" + _
                    rowAfterLama.Batch_No)

                    TBLMasterProductQty_Dummy.Rows.Clear()
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty_Dummy, -1, _
                    rowAfterLama.Primary_Product, rowAfterLama.Batch_No, rowAfterLama.Expiry_Date)

                    For Each rowQty_Dummy In TBLMasterProductQty_Dummy.Rows
                        rowQty = TBLMasterProductQty.Newtblmaster_productqtyRow
                        For ij As Integer = 0 To TBLMasterProductQty.Columns.Count - 1
                            rowQty.Item(ij) = rowQty_Dummy.Item(ij)
                        Next
                        rowQty.Primary = Me.GetPrimaryMax(TBLMasterProductQty) + 1
                        TBLMasterProductQty.Rows.Add(rowQty)
                    Next
                End If
            Next
            For Each rowBeforeLama In TBL1.Rows
                If ListKombinasi.Contains(rowBeforeLama.Primary_Product.ToString + "-" + _
                rowBeforeLama.Primary_Storage_Area.ToString + "-" + _
                rowBeforeLama.Batch_No) = False Then
                    ListKombinasi.Add(rowBeforeLama.Primary_Product.ToString + "-" + _
                    rowBeforeLama.Primary_Storage_Area.ToString + "-" + _
                    rowBeforeLama.Batch_No)

                    TBLMasterProductQty_Dummy.Rows.Clear()
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty_Dummy, -1, _
                    rowBeforeLama.Primary_Product, rowBeforeLama.Batch_No, rowBeforeLama.Expiry_Date)

                    For Each rowQty_Dummy In TBLMasterProductQty_Dummy.Rows
                        rowQty = TBLMasterProductQty.Newtblmaster_productqtyRow
                        For ij As Integer = 0 To TBLMasterProductQty.Columns.Count - 1
                            rowQty.Item(ij) = rowQty_Dummy.Item(ij)
                        Next
                        rowQty.Primary = Me.GetPrimaryMax(TBLMasterProductQty) + 1
                        TBLMasterProductQty.Rows.Add(rowQty)
                    Next
                End If
            Next
            For Each rowDGV In vDGVItemAfter.Rows
                With rowDGV
                    If ListKombinasi.Contains(.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value.ToString + "-" + _
                    .Cells(vFrm.ColPrimaryLocationAfter.Index).Value.ToString + "-" + _
                    .Cells(vFrm.ColBatchNoAfter.Index).Value) = False Then
                        ListKombinasi.Add(.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value.ToString + "-" + _
                        .Cells(vFrm.ColPrimaryLocationAfter.Index).Value.ToString + "-" + _
                        .Cells(vFrm.ColBatchNoAfter.Index).Value)

                        If IsDate(.Cells(vFrm.ColExpiryAfter.Index).Value) = False Then
                            Expiry1 = #1/1/1950#
                        Else
                            Expiry1 = CDate(.Cells(vFrm.ColExpiryAfter.Index).Value)
                        End If

                        TBLMasterProductQty_Dummy.Rows.Clear()

                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty_Dummy, -1, _
                        .Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value.ToString, _
                        .Cells(vFrm.ColBatchNoAfter.Index).Value, Expiry1)

                        For Each rowQty_Dummy In TBLMasterProductQty_Dummy.Rows
                            rowQty = TBLMasterProductQty.Newtblmaster_productqtyRow
                            For ij As Integer = 0 To TBLMasterProductQty.Columns.Count - 1
                                rowQty.Item(ij) = rowQty_Dummy.Item(ij)
                            Next
                            rowQty.Primary = Me.GetPrimaryMax(TBLMasterProductQty) + 1
                            TBLMasterProductQty.Rows.Add(rowQty)
                        Next
                    End If
                End With
            Next
            For Each rowDGV In vDGVItemBefore.Rows
                With rowDGV
                    If ListKombinasi.Contains(.Cells(vFrm.ColPrimaryCodeNo.Index).Value.ToString + "-" + _
                    .Cells(vFrm.ColPrimaryLocation.Index).Value.ToString + "-" + _
                    .Cells(vFrm.ColBatch.Index).Value) = False Then
                        ListKombinasi.Add(.Cells(vFrm.ColPrimaryCodeNo.Index).Value.ToString + "-" + _
                        .Cells(vFrm.ColPrimaryLocation.Index).Value.ToString + "-" + _
                        .Cells(vFrm.ColBatch.Index).Value)

                        If IsDate(.Cells(vFrm.ColExpiredDate.Index).Value) = False Then
                            Expiry1 = #1/1/1950#
                        Else
                            Expiry1 = CDate(.Cells(vFrm.ColExpiredDate.Index).Value)
                        End If

                        TBLMasterProductQty_Dummy.Rows.Clear()

                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty_Dummy, -1, _
                        .Cells(vFrm.ColPrimaryCodeNo.Index).Value.ToString, _
                        .Cells(vFrm.ColBatch.Index).Value, Expiry1)

                        For Each rowQty_Dummy In TBLMasterProductQty_Dummy.Rows
                            rowQty = TBLMasterProductQty.Newtblmaster_productqtyRow
                            For ij As Integer = 0 To TBLMasterProductQty.Columns.Count - 1
                                rowQty.Item(ij) = rowQty_Dummy.Item(ij)
                            Next
                            rowQty.Primary = Me.GetPrimaryMax(TBLMasterProductQty) + 1
                            TBLMasterProductQty.Rows.Add(rowQty)
                        Next
                    End If
                End With
            Next

            'hapus item lama before dan after
            For Each rowBeforeLama In TBL1.Rows
                With rowBeforeLama
                    Me.Koreksi4HapusOut(TBLMasterProductQty, .Primary_Product, .Qty, -1, .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                End With
            Next
            For Each rowAfterLama In TBL.Rows
                With rowAfterLama
                    Me.Koreksi4HapusIn(TBLMasterProductQty, .Primary_Product, .Qty, -1, .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                End With
            Next

            'simpan item baru, before dan after (stlh item before, diselak dulu, liat point 3)
            For Each rowDGV In vDGVItemBefore.Rows
                With rowDGV
                    If IsDate(.Cells(vFrm.ColExpiredDate.Index).Value) = False Then
                        Expiry1 = #1/1/1950#
                    Else
                        Expiry1 = CDate(.Cells(vFrm.ColExpiredDate.Index).Value)
                    End If

                    Me.Koreksi4SimpanOut(TBLMasterProductQty, .Cells(vFrm.ColPrimaryCodeNo.Index).Value, .Cells(vFrm.ColQty.Index).Value, _
                    -1, .Cells(vFrm.ColPrimaryLocation.Index).Value, .Cells(vFrm.ColBatch.Index).Value, Expiry1)
                End With
            Next

            For Each rowDGV In vDGVItemAfter.Rows
                With rowDGV
                    If IsDate(.Cells(vFrm.ColExpiryAfter.Index).Value) = False Then
                        Expiry1 = #1/1/1950#
                    Else
                        Expiry1 = CDate(.Cells(vFrm.ColExpiryAfter.Index).Value)
                    End If

                    Me.Koreksi4SimpanIn(TBLMasterProductQty, .Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value, .Cells(vFrm.ColQtyAfter.Index).Value, _
                    -1, .Cells(vFrm.ColPrimaryLocationAfter.Index).Value, .Cells(vFrm.ColBatchNoAfter.Index).Value, -1, Expiry1)
                End With
            Next

            'jika ada yang negatif, maka salah
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

            '3. untuk mencegah mandeknya perhitungan cost karenasiklik bundling, maka ditambahkan 1 batasan lagi
            'misal: kondisi awal qty product A = 5
            'lalu bikin bundling A + B -> A = A before 5, A after 10
            'lalu bundlingnya di edit, A + B - > A, dimana A before diedit: 5 -> 10
            'hal ini secara qty tidak akan bikin minus, tetapi secara perhitungan cost akan mengakibatkan siklik bundling
            'ini harus dihindari. secara fisik hal ini juga tdk mungking karena sebab, ketika edit A bisa di ubah dari 5 ke 10, 5-nya sebenarnya adalah hasil dari bundling
            'secara fisik, ini tidak mungkin terjadi
            'solusi: pastikan qty before tidak diambil dari qty after bundling yang sama, tapi dari transaksi lain.
            'teknis:
            '1. load data product qty
            '2. kurangkan dengan item before dan after lama
            '3. kurangkan dengan item dgv before
            '4. cek qty, jika ada yang minus, maka salah
            Dummy = Me.ApakahSiklikBundlingTepat(vFrm, vDGVItemAfter, vDGVItemBefore, vPrimaryTransaksi, TBL, TBL1)
            If Dummy <> "" Then
                MsgBox("The following product(s) from item after will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            Dim ObjCustBundling As New CustomSMESS.ClsTransaksiStockBundling
            If ObjCustBundling.ApakahDataBenar_Custom02(vPrimaryTransaksi, vBaru, vTblProdQty, vFrm.TblItemBefCustom, vFrm.TblItemAftCustom, vListProduct) = False Then
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
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vTBLMasterProductQty As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable) As Double
        Dim RowMasterProductQty As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each RowMasterProductQty In vTBLMasterProductQty.Rows
                If RetVal < RowMasterProductQty.Primary Then
                    RetVal = RowMasterProductQty.Primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahSiklikBundlingTepat(ByVal vFrm As FrmTransaksiStockBundling, ByVal vDGVItemAfter As DataGridView, _
    ByVal vDGVItemBefore As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLAfter As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable, _
    ByVal vTBLBefore As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable) As String
        Dim rowBeforeLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow
        Dim rowAfterLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow
        Dim RetVal As String = ""
        Dim ListPrimaryProduct As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim Lanjut As Boolean = False
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ListPrimaryProduct2 As New List(Of Double)
        Dim JmlGudang, JmlObjLama, JmlObjAfter, JmlDGVTotal, JmlSetelahEdit As Decimal
        Dim PrimaryProduct, SecondaryQty, PrimaryLocation As Double
        Dim BatchNo As String = ""
        Dim objStockBundlingItem As New clsTransaksiStockBundlingItem
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            '0. cek apakah ada siklik bundling, ada kalo ada di after dan ada di before
            ListPrimaryProduct.Clear()
            For Each rowDGV In vDGVItemAfter.Rows
                If ListPrimaryProduct.Contains(rowDGV.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value) = False Then
                    ListPrimaryProduct.Add(rowDGV.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value)
                End If
            Next
            For Each rowDGV In vDGVItemBefore.Rows
                '0. hanya berlaku untuk product yang ada di after dan before
                If ListPrimaryProduct.Contains(rowDGV.Cells(vFrm.ColPrimaryCodeNoAfter.Index).Value) = True Then
                    'jika ada maka lakukan beberapa hal dibawah ini, hanya untuk item yang ada siklik bundlingnya saja
                    '1. load data product qty
                    '2. kurangkan dengan item before dan after lama -> mengembalikan ke kondisi awal
                    '3. kurangkan dengan selisih = item dgv before - item after 
                    '4. cek qty, jika ada yang minus, maka salah
                    With rowDGV
                        PrimaryProduct = .Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        SecondaryQty = -1
                        PrimaryLocation = .Cells(vFrm.ColPrimaryLocation.Index).Value
                        BatchNo = .Cells(vFrm.ColBatch.Index).Value
                        If IsDate(.Cells(vFrm.ColExpiredDate.Index).Value) = False Then
                            ExpiryDate = #1/1/1950#
                        Else
                            ExpiryDate = .Cells(vFrm.ColExpiredDate.Index).Value
                        End If

                        JmlGudang = ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                        'perhitungan menggunakan nilai jumlah dgv yang baru
                        JmlObjLama = objStockBundlingItem.GetJmlB4(vPrimaryTransaksi, PrimaryProduct, _
                        SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
                        JmlObjAfter = objStockBundlingItem.GetJmlAfter(vPrimaryTransaksi, PrimaryProduct, _
                        SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)

                        'Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
                        JmlDGVTotal = .Cells(vFrm.ColQty.Index).Value

                        If JmlGudang >= JmlObjAfter Then
                            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal - JmlObjAfter
                        Else
                            JmlSetelahEdit = JmlObjLama - JmlDGVTotal
                        End If
                        If JmlSetelahEdit < 0 Then
                            If ListPrimaryProduct2.Contains(.Cells(vFrm.ColPrimaryCodeNo.Index).Value) = False Then
                                ListPrimaryProduct2.Add(.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                            End If
                        End If
                    End With
                End If
            Next

            If ListPrimaryProduct2.Count > 0 Then
                For Each PrimaryProduct In ListPrimaryProduct2
                    If RetVal = "" Then
                        RetVal = ObjProduct.GetProductCode(PrimaryProduct)
                    Else
                        RetVal = RetVal + ", " + ObjProduct.GetProductCode(PrimaryProduct)
                    End If
                Next
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListPrimaryProduct = Nothing
            ListPrimaryProduct2 = Nothing
            objStockBundlingItem = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar_Obs(ByVal vBaru As Boolean, ByVal vFrm As FrmTransaksiStockBundling, ByVal vDGVItemAfter As DataGridView, _
    ByVal vDGVItemBefore As DataGridView, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim Obj1 As New Core_Module.clsMasterListedSoftware

        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAfterLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow

        Dim DTA1 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBL1 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBeforeLama As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow

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
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
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
    ByVal vFrm As FrmTransaksiStockBundling, ByVal vDGVItemBefore As DataGridView) As Boolean
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

    'Private Function CekPerhitunganTanpaBatas1(ByVal baru As Boolean, ByVal vDGVItemBefore As DataGridView, _
    'ByVal vDGVItemAfter As DataGridView, ByVal i As FrmTransaksiStockBundling) As Boolean
    '    Dim TBLDummy1 As New dseTransaksiStockBundling.DataTable3DataTable
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim DTA As New dseTransaksiStockBundlingTableAdapters.DataTable3TableAdapter
    '    Dim TglDari As Date = ObjCD.GetFirstDateofTheMonth(Me.mTanggalTransaksi)
    '    Dim TglSampai As Date = ObjCD.GetLasDateOfTheMonth(Me.mTanggalTransaksi)
    '    Dim rowBundling2 As dseTransaksiStockBundling.DataTable3Row
    '    Dim Filter As String = ""
    '    Dim rowDGV, rowDGV1 As DataGridViewRow
    '    Dim DTA1 As New dseTransaksiStockBundlingTableAdapters.DataTable6TableAdapter
    '    Dim TBL1 As New dseTransaksiStockBundling.DataTable6DataTable
    '    Dim row1 As dseTransaksiStockBundling.DataTable6Row
    '    Dim Dummy As String = ""

    '    Try
    '        DTA.Fill(TBLDummy1, TglSampai, TglDari)

    '        'hilangkan item2 dari transaksi ini
    '        If baru = False Then
    '            Me.HilangkanItem(TBLDummy1)
    '        End If

    '        'tambahkan item2 dari transaksi ini
    '        Me.TambahkanItem(TBLDummy1, vDGVItemBefore, vDGVItemAfter, i)

    '        'cek kemungkinkan perhitungan tanpa batas, pada dgv
    '        '1. jika ada item after yang menggunakan item before
    '        For Each rowDGV In vDGVItemBefore.Rows
    '            For Each rowDGV1 In vDGVItemAfter.Rows
    '                Filter = "[Primary Transaksi] <> " + Me.mPrimary.ToString + _
    '                " AND [Primary Product Before] = " + rowDGV1.Cells(i.ColPrimaryCodeNoAfter.Index).Value.ToString + _
    '                " AND [Primary Product After] = " + rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value.ToString

    '                TBLDummy1.DefaultView.RowFilter = Filter
    '                If TBLDummy1.DefaultView.Count <> 0 Then
    '                    rowBundling2 = TBLDummy1.DefaultView(0).Row
    '                    MsgBox("This transaction will conflict with " + _
    '                    Me.GetNoTransaksi(CDbl(Me.mPrimary)) + _
    '                    " and will cause error in closing data. Please contact your vendor for further infomation.", _
    '                    MsgBoxStyle.Information, "Information")
    '                    Return False
    '                End If
    '            Next
    '        Next

    '        '2. jika siklik tidak langsung, misal A+B -> C, C+D -> E, C+A->B
    '        'a. load data item
    '        DTA1.Fill(TBL1, TglSampai, TglDari, Me.mPrimary)
    '        'b. tambahkan kombinasi item before after untuk transaksi ini
    '        For Each rowDGV In vDGVItemBefore.Rows
    '            For Each rowDGV1 In vDGVItemAfter.Rows
    '                filter = "[Primary Product After] = " + rowDGV1.Cells(i.ColPrimaryCodeNoAfter.Index).Value.ToString + _
    '                " AND [Primary Product Before] = " + rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value.ToString
    '                TBL1.DefaultView.RowFilter = filter
    '                If TBL1.DefaultView.Count = 0 Then
    '                    row1 = TBL1.NewDataTable6Row

    '                    row1.Primary_Product_After = rowDGV1.Cells(i.ColPrimaryCodeNoAfter.Index).Value
    '                    row1.Primary_Product_Before = rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value
    '                    row1.Code_No_After = rowDGV1.Cells(i.ColCodeNoAfter.Index).Value
    '                    row1.Code_No_Before = rowDGV.Cells(i.ColCodeNo.Index).Value

    '                    TBL1.Rows.Add(row1)
    '                End If
    '            Next
    '        Next
    '        'b. untuk setiap kombinasi item before after, jalankan fungsi cari siklik
    '        For Each rowDGV In vDGVItemBefore.Rows
    '            For Each rowDGV1 In vDGVItemAfter.Rows
    '                Dummy = rowDGV.Cells(i.ColCodeNo.Index).Value.ToString _
    '                + " -> " + rowDGV1.Cells(i.ColCodeNoAfter.Index).Value.ToString

    '                If Me.AdaSiklikBundling(rowDGV1.Cells(i.ColPrimaryCodeNoAfter.Index).Value, _
    '                rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value, Dummy, TBL1) = True Then
    '                    MsgBox("Cyclic bundling occured: " + Dummy + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
    '                    Return False
    '                End If
    '            Next
    '        Next

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBLDummy1.Dispose()
    '        DTA.Dispose()
    '        ObjCD = Nothing
    '        DTA1.Dispose()
    '        TBL1.Dispose()
    '    End Try
    'End Function

    Private Function AdaSiklikBundling(ByVal vPrimaryCodeNoAfter As Double, ByVal vPrimaryCodeNoFind As Double, ByRef vDummy As String, _
    ByVal vTBL As dseTransaksiStockBundling.DataTable6DataTable) As Boolean
        Dim Filter As String = ""
        Dim ListPrimary As New List(Of Double)
        Dim row As dseTransaksiStockBundling.DataTable6Row
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

    Private Function GetCodeNo(ByVal vPrimaryCodeNoAfter As Double, ByVal vTBL As dseTransaksiStockBundling.DataTable6DataTable) As String
        Dim Filter As String = ""
        Dim row As dseTransaksiStockBundling.DataTable6Row

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

    Private Sub HilangkanItem(ByVal TBLDummy As dseTransaksiStockBundling.DataTable3DataTable)
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

    Private Sub TambahkanItem(ByVal TBLDummy As dseTransaksiStockBundling.DataTable3DataTable, ByVal DGVBefore As DataGridView, _
    ByVal DGVAfter As DataGridView, ByVal i As FrmTransaksiStockBundling)
        Dim rowDGV_B4 As DataGridViewRow
        Dim rowDGV_After As DataGridViewRow
        Dim rowBundling As dseTransaksiStockBundling.DataTable3Row

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

    Public Overridable Function SimpanBaru(ByVal vdgvItemB4 As DataGridView, ByVal vdgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiStockBundling) As Boolean
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
    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiStockBundling, Optional ByVal vTampilkanPesan As Boolean = True, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vKoreksiStock As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjItem As New clsTransaksiStockBundlingItem

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustBundling As New CustomSMESS.ClsTransaksiStockBundling
        Dim ListProductCustom As New List(Of Double)
        Dim DescSC As String = "Stock bundling"

        Try
            If Me.ApakahDataBenar(True, vFrm, vdgvItemAfter, vdgvItemB4, -1, tblProdQty, ListProductCustom, vCekDataBenar) = True Then
                'insert query
                If DTA.Insert(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, _
                                Me.mPrimaryProses, Me.mPrimaryWorkUnit, Me.mRemarks, vUserCode, Now, vUserCode, _
                                Now, Now, Me.mTipeBundling, Me.mPrimaryWorkShift) > 0 Then

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
                    ObjItem.Simpan(vdgvItemB4, vdgvItemAfter, Me.mTanggalTransaksi, Me.mPrimary, _
                    Me.mNoTransaksi, vKoreksiStock)

                    '################################CUSTOM INOUT#############################################
                    ObjCustBundling.Simpan_Custom(Me.mPrimary, Me.mNoTransaksi, Me.mTanggalTransaksi, DescSC, _
                    vFrm.TblItemBefCustom, vFrm.TblItemAftCustom, tblProdQty, ListProductCustom)
                    '################################CUSTOM INOUT#############################################

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                Return Me.SimpanBaru1(vdgvItemB4, vdgvItemAfter, vUserCode, vFrm, vTampilkanPesan, vCekDataBenar, vKoreksiStock, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter)
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
                If DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Purchase1)).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Purchase1)).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Purchase1)).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, CInt(EnumTipeBundling.Purchase1))
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                        Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    End If
                End If
                LoopCount += 1

                If LoopCount > 50 Then
                    MsgBox("Logic error: invalid function: 'Cek Nomor' on clstrans_bundling, please contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
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

    Public Overridable Function SimpanPerubahan(ByVal dgvItemB4 As DataGridView, ByVal dgvItemAfter As DataGridView, _
    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiStockBundling) As Boolean
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
    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiStockBundling) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim ObjItem As New clsTransaksiStockBundlingItem

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustBundling As New CustomSMESS.ClsTransaksiStockBundling
        Dim ListProductCustom As New List(Of Double)
        Dim DescSC As String = "Stock bundling"

        Try
            If Me.ApakahDataBenar(False, vFrm, dgvItemAfter, dgvItemB4, Me.mPrimary, tblProdQty, ListProductCustom) = True Then
                If DTA.Update(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, _
                                  Me.mPrimaryProses, Me.mPrimaryWorkUnit, Me.mRemarks, vUserCode, Now, Now, _
                                  Me.mPrimaryWorkShift, Me.mPrimary, Me.mTimestamp) > 0 Then

                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItemB4, dgvItemAfter, Me.mTanggalTransaksi, Me.mPrimary, Me.mNoTransaksi)

                    '################################CUSTOM INOUT#############################################
                    ObjCustBundling.Hapus_Custom(Me.mPrimary, tblProdQty, ListProductCustom, False)
                    ObjCustBundling.Simpan_Custom(Me.mPrimary, Me.mNoTransaksi, Me.mTanggalTransaksi, DescSC, _
                    vFrm.TblItemBefCustom, vFrm.TblItemAftCustom, tblProdQty, ListProductCustom)
                    ObjCustBundling.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)
                    '################################CUSTOM INOUT#############################################

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
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, EnumTipeBundling.Purchase1)
                Return Me.SimpanPerubahan1(dgvItemB4, dgvItemAfter, vUserCode, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vTampilkanPesan As Boolean) As Boolean
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
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, vStrictEdit, vTampilkanPesan)
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
    ByVal vStrictEdit As Boolean, ByVal vTampilkanPesan As Boolean) As Boolean
        Dim Obj As New clsTransaksiStockBundlingItem
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockBundling
        Dim ListProductCustom As New List(Of Double)

        Try
            If Me.ApakahBolehHapus(vUser, False, tblProdQty, ListProductCustom, vTampilkanPesan, vStrictEdit) = True Then
                'hapus data item 
                Obj.Hapus(vPrimaryTransaksi)

                '################################CUSTOM INOUT#############################################
                ObjCustInOut.Hapus_Custom(Me.mPrimary, tblProdQty, ListProductCustom, True)
                '################################CUSTOM INOUT#############################################

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

    Public Function GetTanggalTransaksi(ByVal vPrimaryBundling As Double) As Date
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter

        Try
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
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
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
    ByVal DGV As DataGridView, ByVal TBL As dseTransaksiStockBundling.tbltransaksi_stockbundling1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal vTipeBundling As EnumTipeBundling = EnumTipeBundling.Inventory) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%", vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_Tanggal(TBL, vTanggalSampai, vTanggalDari, vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%", vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%", vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, vTipeBundling, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProsesCode Then
                QueryResult = DTA.FillBy_ProcessLIKE(TBL, Criteria + "%", vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkUnitCode Then
                QueryResult = DTA.FillBy_WorkUnitLIKE(TBL, Criteria + "%", vTipeBundling)
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkShift Then
                QueryResult = DTA.FillBy_WorkshiftLIKE(TBL, Criteria + "%", vTipeBundling)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(5).Visible = False
                DGV.Columns(1).HeaderText = "Transaction No"
                DGV.Columns(2).HeaderText = "Transaction Date"
                DGV.Columns(3).HeaderText = "Process Code"
                DGV.Columns(4).HeaderText = "WorkUnit Code"

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

    Public Sub ResetObject(ByVal vDgvItemB4 As DataGridView, ByVal vDgvItemAfter As DataGridView)
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
            vDgvItemB4.Rows.Clear()
            vDgvItemAfter.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateStockCard_TransaksiStockBundling(ByVal vPrg As ProgressBar, ByVal vLblProgress As Label)
        Dim DTAB4 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling2TableAdapter
        Dim TBLB4 As New dseTransaksiStockBundling.tbltransaksi_stockbundling2DataTable
        Dim rowB4 As dseTransaksiStockBundling.tbltransaksi_stockbundling2Row
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling3TableAdapter
        Dim TBLAfter As New dseTransaksiStockBundling.tbltransaksi_stockbundling3DataTable
        Dim rowAfter As dseTransaksiStockBundling.tbltransaksi_stockbundling3Row

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
                    Qty_In, Qty_Out, DescSC, InOut, .Tanggal_Transaksi, BatchNo, ExpiryDate, PrimaryLocation, .Remarks)
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
                    Qty_In, Qty_Out, DescSC, InOut, .Tanggal_Transaksi, BatchNo, ExpiryDate, PrimaryLocation, .Remarks)
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

    Public Function isidgv_wo(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As dseTransaksiStockBundling.tbltrans_mo_item_wo_findDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vstatus_firmplan As ClsTrans_ManufacturingOrder.EnumPilihStatusWO, ByVal vstatus_onprogress As ClsTrans_ManufacturingOrder.EnumPilihStatusWO, _
   Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean

        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_wonumber_like(TBL, Criteria + "%", vstatus_firmplan, vstatus_onprogress)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_tgldari_tglsampai(TBL, vTanggalSampai, vTanggalDari, vstatus_firmplan, vstatus_onprogress)
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkUnitCode Then
                QueryResult = DTA.FillBy_workunitcode_like(TBL, Criteria + "%", vstatus_firmplan, vstatus_onprogress)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_productcode_like(TBL, Criteria + "%", vstatus_firmplan, vstatus_onprogress)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_product_desc_like(TBL, Criteria + "%", vstatus_firmplan, vstatus_onprogress)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_remarks_like(TBL, Criteria + "%", vstatus_firmplan, vstatus_onprogress)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).Visible = False
                DGV.Columns(2).HeaderText = "Wo Number"
                DGV.Columns(3).HeaderText = "Date"
                DGV.Columns(4).HeaderText = "Work Unit"
                DGV.Columns(5).HeaderText = "Code No"
                DGV.Columns(6).HeaderText = "Remarks"

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

    Public Function getjml_wo(ByVal Criteria As String, _
  ByVal SearchCriteria As EnumSearchCriteria, ByVal vstatus_firmplan As Integer, ByVal vstatus_onprogress As Integer, _
  Optional ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long

        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltrans_mo_item_wo_findTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Dummy = DTA.ScalarQuery_getjmlby_wonumber_like(Criteria + "%", vstatus_firmplan, vstatus_onprogress)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Dummy = DTA.ScalarQuery_getjmlby_tgldarisampai(vTanggalSampai, vTanggalDari, vstatus_firmplan, vstatus_onprogress)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByWorkUnitCode Then
                Dummy = DTA.ScalarQuery_getjmlby_workunitcode_like(Criteria + "%", vstatus_firmplan, vstatus_onprogress)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Dummy = DTA.ScalarQuery_getjmlby_productcode_like(Criteria + "%", vstatus_firmplan, vstatus_onprogress)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Dummy = DTA.ScalarQuery_getjmlby_desclike(Criteria + "%", vstatus_firmplan, vstatus_onprogress)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Dummy = DTA.ScalarQuery_getjmlby_remarks_like(Criteria + "%", vstatus_firmplan, vstatus_onprogress)
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
End Class
