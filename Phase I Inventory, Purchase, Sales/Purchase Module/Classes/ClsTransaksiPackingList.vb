Option Explicit On
Option Strict Off

Public Class ClsTransaksiPackingList
    Dim ObjCD As New Core_Module.ClsCekData

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Private mPrimary As Double = 0
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Today
    Private mUrutan As Integer = 0
    Private mSupplier_sGITNo As String = ""
    Private mPrimarySupplier As Double = -1
    Private mReceievedDate As Date = Today
    Private mExtimatedArrivalDate As Date = Today
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

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

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property EstimatedArrivalDate() As Date
        Get
            Return Me.mExtimatedArrivalDate
        End Get
        Set(ByVal value As Date)
            mExtimatedArrivalDate = value
        End Set
    End Property

    Public Property ReceivedDate() As Date
        Get
            Return Me.mReceievedDate
        End Get
        Set(ByVal value As Date)
            mReceievedDate = value
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

    Public Property SupplierGITNumber() As String
        Get
            Return Me.mSupplier_sGITNo
        End Get
        Set(ByVal value As String)
            mSupplier_sGITNo = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
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

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Function Cetak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPackingList As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objPL_Custom As New CustomSMESS.ClsTransaksiPackingList

        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PackingList, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            ObjSetting = Nothing
            Return objPL_Custom.Cetak(CRV, vPrimaryPackingList)
        End If

        Dim Report As New CrpCetakTransaksiPackingList
        Dim DS As New DseTransaksiPackingList
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1

        Try
            Dim DTA1 As New DseTransaksiPackingListTableAdapters.TblCetakHeaderTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.FillBy_PrimaryPackingLIst(DS.TblCetakHeader, vPrimaryPackingList)
            DTA1.Dispose()
            Dim DTA2 As New DseTransaksiPackingListTableAdapters.tblCetakItemTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.FillBy_PrimaryPackingList(DS.tblCetakItem, vPrimaryPackingList)
            DTA2.Dispose()

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)

            CRV.ReportSource = Report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Public Function GetJmlItem_ByPrimaryPO(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim DTA1 As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTA1.Connection.ConnectionString

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDesc(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                Return DTA1.ScalarQuery_GetJmlByKodeSupplierLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                Return DTA1.ScalarQuery_GetJmlBySupplierNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.SupplierPackingListNo Then
                Return DTA.ScalarQuery_GetJmlBySupplierPackingListNumber(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                Return DTA.ScalarQuery_GetJmlByPONumberLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/GIT/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglistDataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglistRow
        Dim ObjItem As New ClsTransaksiPackingListItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimarySupplier = .Primary_Supplier
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mReceievedDate = .Received_Date
                    Me.mSupplier_sGITNo = .Supplier_Packing_List_No
                    Me.mExtimatedArrivalDate = .Estimated_Arrival_Date

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

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.SupplierPackingListNo Then
                QueryResult = DTA.FillBy_SupplierPackingListNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SupplierCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_SupplierNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                QueryResult = DTA.FillBy_PONumberLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(7).DefaultCellStyle = ij
                DGV.Columns(6).DefaultCellStyle = ij
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryPackingList As Double, _
    ByVal StrictEditting As Boolean, ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'packing list tidak boleh diedit jika:
            '0. Untuk strict editting harus lihat user edit apakah sama dengan yang menciptakan
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryPackingList)
                If CreatedBy = vUser Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            '1. batas tutup transaksi purchase sudah lewat
            'msgbox("belum")
            '2. sudah ada stock in yang merujuk ke packing list yang bersangkutan
            Dim NoStockIn As String = ObjStockInOut.GetStockInNumberByPrimaryPackingList(vPrimaryPackingList)
            If NoStockIn <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This packing list is refered by stock in purchase:" + NoStockIn + ". You cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjStockInOut = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryPackingList As Double, _
    ByVal StrictEditting As Boolean, ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjBookingGIT As New Sales_Module.ClsTransaksiBookingGIT

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'packing list tidak boleh diedit jika:
            '1. batas tutup transaksi purchase sudah lewat
            'msgbox("belum")
            '2. sudah ada stock in yang merujuk ke packing list yang bersangkutan
            Dim NoTransaksi As String = ObjStockInOut.GetStockInNumberByPrimaryPackingList(vPrimaryPackingList)
            If NoTransaksi <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This packing list is refered by stock in purchase: " + NoTransaksi + ". You cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. Ada booking GIT yang merujuk pada transaksi ini
            NoTransaksi = ObjBookingGIT.GetNoBookingGITAktifByPrimaryGIT(vPrimaryPackingList)
            If NoTransaksi <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This packing list is refered by GIT booking: " + NoTransaksi + ". You cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '0. Untuk strict editting harus lihat user edit apakah sama dengan yang menciptakan
            If StrictEditting = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryPackingList)
                If CreatedBy = vUser Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'kalo mpe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjStockInOut = Nothing
            ObjBookingGIT = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Try
            'msgbox("belum")
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
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetPrimary(vNoTransaksi)
            If retval.HasValue Then
                If retval.Value <> -1 Then
                    Return retval
                End If
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Function


    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPackingList_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode)
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
    Optional ByVal vSETPOB As Boolean = True, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim ObjItem As New ClsTransaksiPackingListItem
        Dim ObjLS As New Core_Module.clsMasterListedSoftware

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, vCekDataBenar) = True Then
                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA.Insert(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mTanggalTransaksi, _
                Me.mPrimarySupplier, Me.mSupplier_sGITNo, Me.mReceievedDate, Me.mExtimatedArrivalDate, Me.mRemarks, _
                CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mUrutan = CInt(Me.mNoTransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.Simpan_Multi(vdgvItem, Me.mPrimary, vSETPOB)
                    'simpan data booking GIT
                    'Me.SimpanBookingGIT(Me.mPrimary, Me.mTanggalTransaksi, vUserCode)
                    If vCekDataBenar = True Then
                        Me.CekNomor(DTA)
                    End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mUrutan
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If DTA.ScalarQuery_GetJmlByUrutan(Me.mUrutan, TanggalSampai, TanggalDari).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mUrutan, TanggalSampai, TanggalDari).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mUrutan, TanggalSampai, TanggalDari).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mUrutan, TanggalSampai, TanggalDari)
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                        Me.mUrutan = CInt(Me.mNoTransaksi.Substring(0, 6))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_Urutan(Me.mNoTransaksi, Me.mUrutan, Me.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPackingList_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode)
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

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim ObjItem As New ClsTransaksiPackingListItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(False) = True Then
                If DTA.Update(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mTanggalTransaksi, _
                Me.mPrimarySupplier, Me.mSupplier_sGITNo, Me.mReceievedDate, Me.mExtimatedArrivalDate, _
                Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mUrutan = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me, False)
                    'simpan data item baru
                    ObjItem.Simpan_Multi(dgvItem, Me.mPrimary)
                    'hapus data booking git-software generated yang lama, kemudian dibuat ulang booking invnya
                    'Me.Hapus_Booking_GIT(Me.mPrimary)
                    'Me.SimpanBookingGIT(Me.mPrimary, Me.mTanggalTransaksi, vUserCode)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(dgvItem, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    'Public Sub SimpanBookingGIT(ByVal vPrimaryPackingList As Double, _
    'ByVal vTanggalPL As Date, ByVal vUserCode As String)
    '    Dim TBLPackingListItem As New DseTransaksiPackingList.tbltrans_packinglist_itemDataTable
    '    Dim DTAPackingListItem As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
    '    Dim Row_PL As DseTransaksiPackingList.tbltrans_packinglist_itemRow
    '    Dim Primary_PO_Min As Double = -1
    '    Dim TBL_Item_PO As New DseTransaksiBookingGIT.tbltrans_po_itemDataTable
    '    Dim DTA_Item_PO As New DseTransaksiBookingGITTableAdapters.tbltrans_po_itemTableAdapter
    '    Dim row_Item_PO As DseTransaksiBookingGIT.tbltrans_po_itemRow
    '    Dim Filter As String = ""
    '    Dim Sort As String = "[Primary IPO] ASC"
    '    Dim rowView As DataRowView
    '    Dim Qty_Booking As Double
    '    Dim TBLHeader As New DseTransaksiBookingGIT.tbltrans_booking_gitDataTable
    '    Dim TBLIsi As New DseTransaksiBookingGIT.tbltrans_booking_gititemDataTable
    '    Dim DTABookingGIT As New Inventory_Module.DseTransaksiBookingInvTableAdapters.tbltrans_booking_gititemTableAdapter
    '    Dim TBLBookingGIT As New Inventory_Module.DseTransaksiBookingInv.tbltrans_booking_gititemDataTable
    '    Dim DTABookingInv As New Inventory_Module.DseTransaksiBookingInvTableAdapters.DataTable7TableAdapter
    '    Dim TBLBookingInv As New Inventory_Module.DseTransaksiBookingInv.DataTable7DataTable
    '    Dim DTASOB As New Inventory_Module.DseTransaksiBookingInvTableAdapters.tbltrans_so_item1TableAdapter
    '    Dim TBLSOB As New Inventory_Module.DseTransaksiBookingInv.tbltrans_so_item1DataTable
    '    Dim DTAPendingDO As New Inventory_Module.DseTransaksiBookingInvTableAdapters.tbltrans_do_itemTableAdapter
    '    Dim TBLPendingDO As New Inventory_Module.DseTransaksiBookingInv.tbltrans_do_itemDataTable
    '    Dim JmlSOB, JmlGITAktif, JmlBookingInv, JmlPendingDO As Double
    '    Dim Filter1 As String = ""

    '    Try
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        'hapus booking git yang merujuk ke packing ini
    '        Me.Hapus_Booking_GIT(vPrimaryPackingList)
    '        'cari primary PO paling kecil pada packing 
    '        DTAPackingListItem.FillBy_PrimaryTransaksi(TBLPackingListItem, vPrimaryPackingList)
    '        Primary_PO_Min = Me.GetPrimaryPOMin(TBLPackingListItem)
    '        'load data item PO yang berasal dari IPO dan primary transaksinya lebih besar sama dengan primary
    '        'po minimal hasil perhitungan
    '        If DTA_Item_PO.Fill(TBL_Item_PO, Primary_PO_Min) = 0 Then
    '            'tidak ada item ipo
    '            Exit Sub
    '        End If
    '        'kurangkan data po dengan stock in purchase yang sudah terjadi
    '        Me.KurangkanDenganStockInPurchase(TBL_Item_PO, Primary_PO_Min)
    '        'kurangkan data po dengan booking GIT yang sudah terjadi
    '        Me.Kurangkan_Dengan_Booking_GIT(TBL_Item_PO)

    '        DTABookingGIT.Fill(TBLBookingGIT)
    '        DTABookingInv.Fill(TBLBookingInv)
    '        DTASOB.Fill(TBLSOB)
    '        DTAPendingDO.Fill(TBLPendingDO)

    '        'untuk setiap item packing list
    '        For Each Row_PL In TBLPackingListItem.Rows
    '            '1. kumpulkan semua item po berdasarkan primary po dan primary product
    '            ', urutkan berdasarkan primary IPO(asc)
    '            Filter = "[Primary Transaksi] = " + Row_PL.Primary_PO.ToString + _
    '            " AND [Primary Product] = " + Row_PL.Primary_Product.ToString
    '            TBL_Item_PO.DefaultView.Sort = Sort
    '            TBL_Item_PO.DefaultView.RowFilter = Filter
    '            If TBL_Item_PO.DefaultView.Count > 0 Then
    '                For Each rowView In TBL_Item_PO.DefaultView
    '                    '2. selama qty item packing list tsb > 0
    '                    If Row_PL.Total_Qty > 0 Then
    '                        row_Item_PO = rowView.Row
    '                        Filter1 = "[Primary SO] =" + row_Item_PO.Primary_SO.ToString + _
    '                        " AND [Primary Product] =" + row_Item_PO.Primary_Product.ToString + ""
    '                        '3. cari qty booking: yaitu qty minimum antara: qty item po, _
    '                        'qty SOB-qty other booking inv-booking inv-pending do, qty packing list
    '                        Qty_Booking = row_Item_PO.Qty  '--> bandingkan dengan qty item po
    '                        If Qty_Booking > Row_PL.Total_Qty Then
    '                            Qty_Booking = Row_PL.Total_Qty   '--> bandingkan dengan qty packing list
    '                        End If
    '                        'belum memperhitungkan SOB
    '                        JmlSOB = Me.GetJmlSOB(Filter1, TBLSOB)
    '                        JmlGITAktif = Me.GetJmlBookingGITAktif(Filter1, TBLBookingGIT)
    '                        JmlBookingInv = Me.GetJmlBookingInv(Filter1, TBLBookingInv)
    '                        JmlPendingDO = Me.GetJmlPendingDO(Filter1, TBLPendingDO)
    '                        JmlSOB = JmlSOB - JmlGITAktif - JmlBookingInv - JmlPendingDO
    '                        If Qty_Booking > JmlSOB Then
    '                            Qty_Booking = JmlSOB
    '                        End If

    '                        '4. jika qty minimum <> 0 maka masukan item booking inv ke dalam tabel sementara
    '                        '   kurangkan qty packing list sejumlah yang dibooking
    '                        '  item booking dikelompokan berdasarkan kode customer
    '                        If Qty_Booking <> 0 Then
    '                            'kurangkan yang dapat dibooking pada packing list
    '                            Row_PL.Total_Qty -= Qty_Booking
    '                            'kurangkan kebutuhan booking
    '                            row_Item_PO.Qty -= Qty_Booking
    '                            'tambahkan booking GIT aktif
    '                            Me.TambahBookingGITAktif(row_Item_PO.Primary_Product, _
    '                            row_Item_PO.Primary_SO, Qty_Booking, TBLBookingGIT, Filter1)
    '                            'generate item booking
    '                            Me.GenerateBookingGIT(TBLHeader, TBLIsi, row_Item_PO, Qty_Booking, _
    '                            vUserCode, vPrimaryPackingList)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        Next
    '        'Simpan data booking GIT, booking GIT sudah berdasarkan customer code
    '        Me.SimpanBookingGIT(TBLHeader, TBLIsi, vTanggalPL)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA_Item_PO.Dispose()
    '        TBL_Item_PO.Dispose()
    '        DTAPackingListItem.Dispose()
    '        TBLPackingListItem.Dispose()
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '        TBLHeader.Dispose()
    '        TBLIsi.Dispose()
    '        DTABookingGIT.Dispose()
    '        TBLBookingGIT.Dispose()
    '        DTABookingInv.Dispose()
    '        TBLBookingInv.Dispose()
    '        DTASOB.Dispose()
    '        TBLSOB.Dispose()
    '        DTAPendingDO.Dispose()
    '        TBLPendingDO.Dispose()
    '    End Try
    'End Sub

    Private Sub TambahBookingGITAktif(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal vQtyBooking As Double, ByVal vTBLBoooking As _
    Inventory_Module.DseTransaksiBookingInv.tbltrans_booking_gititemDataTable, _
    ByVal vFilter As String)
        Dim rowBooking As Inventory_Module.DseTransaksiBookingInv.tbltrans_booking_gititemRow

        Try
            vTBLBoooking.DefaultView.RowFilter = vFilter
            If vTBLBoooking.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowBooking = vTBLBoooking.Newtbltrans_booking_gititemRow
                rowBooking.Primary_Product = vPrimaryProduct
                rowBooking.Primary_SO = vPrimarySO
                rowBooking.QtyBookingGITAktif = vQtyBooking
                vTBLBoooking.Rows.Add(rowBooking)
            Else
                'tambahkan qty booking
                rowBooking = vTBLBoooking.DefaultView(0).Row
                rowBooking.QtyBookingGITAktif += vQtyBooking
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlBookingInv(ByVal vFilter As String, _
    ByVal vTBL As Inventory_Module.DseTransaksiBookingInv.DataTable7DataTable) As Decimal
        Dim row As Inventory_Module.DseTransaksiBookingInv.DataTable7Row

        Try
            vTBL.DefaultView.RowFilter = vFilter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.QtyBookingInv
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlBookingGITAktif(ByVal vFilter As String, _
    ByVal vTBL As Inventory_Module.DseTransaksiBookingInv.tbltrans_booking_gititemDataTable) As Double
        Dim row As Inventory_Module.DseTransaksiBookingInv.tbltrans_booking_gititemRow

        Try
            vTBL.DefaultView.RowFilter = vFilter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.QtyBookingGITAktif
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlPendingDO(ByVal vFilter As String, _
    ByVal vTBL As Inventory_Module.DseTransaksiBookingInv.tbltrans_do_itemDataTable) As Decimal
        Dim row As Inventory_Module.DseTransaksiBookingInv.tbltrans_do_itemRow

        Try
            vTBL.DefaultView.RowFilter = vFilter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.Total_Pending
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSOB(ByVal vFilter As String, _
    ByVal vTBL As Inventory_Module.DseTransaksiBookingInv.tbltrans_so_item1DataTable) As Decimal
        Dim row As Inventory_Module.DseTransaksiBookingInv.tbltrans_so_item1Row

        Try
            vTBL.DefaultView.RowFilter = vFilter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.SOB
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub SimpanBookingGIT(ByVal TBLHeader As DseTransaksiBookingGIT.tbltrans_booking_gitDataTable, _
    'ByVal TBLIsi As DseTransaksiBookingGIT.tbltrans_booking_gititemDataTable, ByVal vTanggalPackingList As Date)
    '    Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gitTableAdapter
    '    Dim DTAItem As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititemTableAdapter
    '    Dim PrimaryBaru As Double = -1
    '    Dim rowHeader As DseTransaksiBookingGIT.tbltrans_booking_gitRow
    '    Dim RowItem As DseTransaksiBookingGIT.tbltrans_booking_gititemRow
    '    Dim ObjBookingGIT As New ClsTransaksiBookingGIT
    '    Dim Status As Integer = 0

    '    Try
    '        'simpan data header
    '        For Each rowHeader In TBLHeader.Rows
    '            With rowHeader
    '                .No_Transaksi = ObjBookingGIT.GetNoBaru(rowHeader._Date)
    '                If .Status_Aktif = False Then
    '                    Status = 0
    '                Else
    '                    Status = 1
    '                End If
    '                DTA.Insert(.No_Transaksi, .No_Transaksi.Substring(0, 6), vTanggalPackingList, _
    '                .Primary_Customer, .Primary_GIT, .Remarks, .SoftwareGenerated, 1, .Created_By, _
    '                .Created_Date, .Modified_By, .Modified_Date, .Timestamp)
    '                PrimaryBaru = ObjBookingGIT.GetPrimary(.No_Transaksi)
    '            End With

    '            'simpan data item
    '            For Each RowItem In TBLIsi.Rows
    '                If RowItem.Primary_Booking = rowHeader.Primary Then
    '                    With RowItem
    '                        DTAItem.Insert(PrimaryBaru, .PrimarySO, .Primary_IPO, .Primary_Product, _
    '                        .Qty, .Primary_Unit, .Remarks, .TransferedToInvBooking)
    '                    End With
    '                End If
    '            Next
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        DTAItem.Dispose()
    '    End Try
    'End Sub

    ''' <summary>
    ''' membuat data header dan data item untuk booking GIT
    ''' </summary>
    ''' <param name="TBLHeader"></param>
    ''' <param name="TBLIsi"></param>
    ''' <param name="rowAcuan"></param>
    ''' <param name="QtyBooking"></param>
    ''' <remarks></remarks>
    'Private Sub GenerateBookingGIT(ByVal TBLHeader As DseTransaksiBookingGIT.tbltrans_booking_gitDataTable, _
    'ByVal TBLIsi As DseTransaksiBookingGIT.tbltrans_booking_gititemDataTable, _
    'ByVal rowAcuan As DseTransaksiBookingGIT.tbltrans_po_itemRow, ByVal QtyBooking As Double, _
    'ByVal vUser As String, ByVal vPrimaryGIT As Double)
    '    Dim Index As Integer = -1
    '    Dim rowHeader As DseTransaksiBookingGIT.tbltrans_booking_gitRow
    '    Dim rowItem As DseTransaksiBookingGIT.tbltrans_booking_gititemRow
    '    Dim Timestamp As DateTime = Now

    '    Try
    '        'cari ditabel header apakah sudah ada customer yang bersangkutan
    '        Index = Me.GetPrimaryBooking(TBLHeader, rowAcuan.Primary_Customer)

    '        'bila tidak ada maka buat data header baru
    '        If Index = -1 Then
    '            Index = TBLHeader.Rows.Count + 1
    '            rowHeader = TBLHeader.Newtbltrans_booking_gitRow
    '            With rowHeader
    '                .Primary = Index
    '                .Primary_Customer = rowAcuan.Primary_Customer
    '                .Created_By = vUser
    '                .Created_Date = Timestamp
    '                .Modified_By = vUser
    '                .Modified_Date = Timestamp
    '                .No_Transaksi = ""
    '                .Primary_GIT = vPrimaryGIT
    '                .Remarks = ""
    '                .SoftwareGenerated = 1
    '                .Status_Aktif = 1
    '                .Timestamp = Timestamp
    '                .Urutan = 0
    '                ._Date = Today
    '            End With
    '            TBLHeader.Rows.Add(rowHeader)
    '        End If

    '        'masukan data item booking berdasarkan primary transaksi acuan
    '        rowItem = TBLIsi.Newtbltrans_booking_gititemRow
    '        With rowItem
    '            .Primary = TBLIsi.Rows.Count + 1
    '            .Primary_Booking = Index
    '            .Primary_IPO = rowAcuan.Primary_IPO
    '            .Primary_Product = rowAcuan.Primary_Product
    '            .Primary_Unit = rowAcuan.Primary_Units
    '            .PrimarySO = rowAcuan.Primary_SO
    '            .Qty = QtyBooking
    '            .Remarks = ""
    '            .TransferedToInvBooking = 0
    '        End With
    '        TBLIsi.Rows.Add(rowItem)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Function GetPrimaryBooking(ByVal vTBLHeader As DseTransaksiBookingGIT.tbltrans_booking_gitDataTable, _
    'ByVal vPrimaryCustomer As Double) As Integer
    '    Dim row As DseTransaksiBookingGIT.tbltrans_booking_gitRow
    '    Try
    '        For Each row In vTBLHeader.Rows
    '            If row.Primary_Customer = vPrimaryCustomer Then
    '                Return row.Primary
    '            End If
    '        Next

    '        Return -1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Sub Kurangkan_Dengan_Booking_GIT(ByVal TBL_Item_PO As DseTransaksiBookingGIT.tbltrans_po_itemDataTable)
    '    Dim Primary_IPO_Min As Double = Double.MaxValue
    '    Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititemTableAdapter
    '    Dim TBL_Item_Booking As New DseTransaksiBookingGIT.tbltrans_booking_gititemDataTable
    '    Dim row_Item_PO As DseTransaksiBookingGIT.tbltrans_po_itemRow
    '    Dim row_Item_Booking As DseTransaksiBookingGIT.tbltrans_booking_gititemRow
    '    Dim counter, counter1 As Double
    '    Dim Filter As String = ""
    '    Dim rowView As DataRowView

    '    Try
    '        'cari ipo minimum
    '        If TBL_Item_PO.Rows.Count = 0 Then
    '            Exit Sub
    '        Else
    '            For Each row_Item_PO In TBL_Item_PO.Rows
    '                If row_Item_PO.Primary_IPO < Primary_IPO_Min Then
    '                    Primary_IPO_Min = row_Item_PO.Primary_IPO
    '                End If
    '            Next
    '        End If

    '        'cari semua item booking GIT yang primary IPO-nya lebih besar atau sama dengan 
    '        'nilai IPO minimum, booking GIT ini harus berstatus masih aktif --> belum ditransfer menjadi
    '        'booking inventory
    '        DTA.FillBy_PrimaryIPOMin(TBL_Item_Booking, Primary_IPO_Min)

    '        'untuk setiap item po kurangkan dengan booking GIT yang sudah terjadi
    '        While counter < TBL_Item_PO.Rows.Count
    '            row_Item_PO = TBL_Item_PO.Rows(counter)
    '            Filter = "[Primary IPO] =" + row_Item_PO.Primary_IPO.ToString + _
    '            " and [Primary Product] =" + row_Item_PO.Primary_Product.ToString
    '            TBL_Item_Booking.DefaultView.RowFilter = Filter

    '            For Each rowView In TBL_Item_Booking.DefaultView
    '                row_Item_Booking = rowView.Row
    '                If row_Item_Booking.Qty > row_Item_PO.Qty Then
    '                    row_Item_PO.Qty = 0
    '                    row_Item_Booking.Qty -= row_Item_PO.Qty
    '                    Exit For
    '                Else
    '                    row_Item_PO.Qty -= row_Item_Booking.Qty
    '                    row_Item_Booking.Qty = 0
    '                End If
    '            Next

    '            While counter1 < TBL_Item_Booking.Rows.Count
    '                row_Item_Booking = TBL_Item_Booking.Rows(counter1)
    '                If row_Item_Booking.Qty = 0 Then
    '                    TBL_Item_Booking.Rows.RemoveAt(counter1)
    '                End If
    '            End While

    '            If row_Item_PO.Qty = 0 Then
    '                TBL_Item_PO.Rows.RemoveAt(counter)
    '            Else
    '                counter += 1
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub KurangkanDenganStockInPurchase(ByVal TBL_Item_PO As DseTransaksiBookingGIT.tbltrans_po_itemDataTable, _
    'ByVal vPrimaryPOMin As Double)
    '    Dim TBL_In_Purchase As New DseTransaksiBookingGIT.tbltransaksi_stockinout_itemDataTable
    '    Dim DTA_In_Purchase As New DseTransaksiBookingGITTableAdapters.tbltransaksi_stockinout_itemTableAdapter
    '    Dim row_In As DseTransaksiBookingGIT.tbltransaksi_stockinout_itemRow
    '    Dim row_PO As DseTransaksiBookingGIT.tbltrans_po_itemRow
    '    Dim rowView As DataRowView
    '    Dim Filter As String = ""
    '    Dim counter As Double = 0
    '    Dim counter1 As Double = 0

    '    Try
    '        DTA_In_Purchase.FillBy_PrimaryPOMin(TBL_In_Purchase, vPrimaryPOMin)
    '        While counter < TBL_Item_PO.Rows.Count
    '            row_PO = TBL_Item_PO.Rows(counter)
    '            Filter = "[Primary PO Purchase] =" + row_PO.Primary_Transaksi.ToString + " AND [Primary Product] =" _
    '            + row_PO.Primary_Product.ToString
    '            TBL_In_Purchase.DefaultView.RowFilter = Filter
    '            For Each rowView In TBL_In_Purchase.DefaultView
    '                row_In = rowView.Row
    '                If row_In.Qty > 0 Then
    '                    If row_In.Qty > row_PO.Qty Then
    '                        'kurangkan sejumlah po
    '                        row_In.Qty -= row_PO.Qty
    '                        row_PO.Qty = 0
    '                        Exit For
    '                    Else
    '                        'kurangkan sejumlah in qty
    '                        row_In.Qty = 0
    '                        row_PO.Qty -= row_In.Qty
    '                    End If
    '                End If
    '            Next
    '            'hapus baris in purchase yang sudah 0
    '            While counter1 < TBL_In_Purchase.Rows.Count
    '                row_In = TBL_In_Purchase.Rows(counter1)
    '                If row_In.Qty = 0 Then
    '                    TBL_In_Purchase.Rows.RemoveAt(counter1)
    '                Else
    '                    counter1 += 1
    '                End If
    '            End While
    '            'cek item po
    '            If row_PO.Qty = 0 Then
    '                'hapuskan baris dengan qty nol
    '                TBL_Item_PO.Rows.RemoveAt(counter)
    '            Else
    '                'pindah untuk kurangkan item po pada baris selanjutnya
    '                counter += 1
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBL_In_Purchase.Dispose()
    '        DTA_In_Purchase.Dispose()
    '    End Try
    'End Sub

    Private Function GetPrimaryPOMin(ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_itemDataTable) As Double
        Dim rowItem As DseTransaksiPackingList.tbltrans_packinglist_itemRow
        Dim dummyPrimary As Double = Double.MaxValue
        Try
            If TBL.Rows.Count > 0 Then
                For Each rowItem In TBL.Rows
                    If rowItem.Primary_PO < dummyPrimary Then
                        dummyPrimary = rowItem.Primary_PO
                    End If
                Next
            Else
                dummyPrimary = -1
            End If

            Return dummyPrimary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetPrimaryBooking_GIT(ByVal vPrimaryCustomer As Double, _
    'ByVal TBL As DseTransaksiBookingGIT.tbltrans_booking_gitDataTable) As Double
    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''' <summary>
    ''' menghapus semua booking GIT yang merujuk pada transaksi ini
    ''' </summary>
    ''' <param name="vPrimaryPackingList"></param>
    ''' <remarks></remarks>
    'Public Sub Hapus_Booking_GIT_ALL(ByVal vPrimaryPackingList As Double)
    '    Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gitTableAdapter

    '    Try
    '        DTA.DeleteQuery_ByPrimaryPackingListAll(vPrimaryPackingList)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Sub

    ''' <summary>
    ''' hapus data transaksi booking inventory yang software generated dan merujuk pada transaksi packing list ybs
    ''' </summary>
    ''' <param name="vPrimaryPackingList"></param>
    ''' <remarks></remarks>
    'Public Sub Hapus_Booking_GIT(ByVal vPrimaryPackingList As Double)
    '    Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gitTableAdapter

    '    Try
    '        DTA.DeleteQuery_DeleteByPrimaryPackingList(vPrimaryPackingList)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Sub

    Public Sub GetPrimaryPO(ByVal vPrimaryPackingList As Double, ByVal vListPrimaryPO As List(Of Double))
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item2TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item2DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vListPrimaryPO.Clear()
            DTA.FillBy_PrimaryPackingList(TBL, vPrimaryPackingList)
            For Each row In TBL.Rows
                vListPrimaryPO.Add(row.Primary_PO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, _
            Optional ByVal TampilkanPesan As Boolean = True) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPackingList_Hapus, vUser)

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
                            retval = Me.Hapus1(vPrimaryTransaksi)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim vListPrimary As New List(Of Double)
        Dim ObjPOItem As New ClsTransaksiPOItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vPrimaryTransaksi, False, "", TampilkanPesan) = True Then
                'hapus semua booking GIT yang merujuk pada transaksi ini
                'Me.Hapus_Booking_GIT_ALL(vPrimaryTransaksi)

                Me.GetPrimaryPO(vPrimaryTransaksi, vListPrimary)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    ObjPOItem.SetOrderBalance(vListPrimary)
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            vListPrimary = Nothing
            ObjPOItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = 0
            mNoTransaksi = ""
            mTanggalTransaksi = Today
            mUrutan = 0
            mSupplier_sGITNo = ""
            mPrimarySupplier = -1
            mReceievedDate = Today
            mExtimatedArrivalDate = Today
            mRemarks = ""

            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mTimestamp = Now

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimaryPackingList As Double) As Date
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTimestamp(vPrimaryPackingList)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub
End Class
