Option Explicit On
Option Strict Off

Public Class ClsTransaksiBookingInv
    Private mprimary As Double = -1
    Private mBookingNo As String = ""
    Private mNoUrut As Integer = -1
    Private mDate As Date = Today
    Private mPrimaryCust As Double = -1
    Private mPrimarySO As Double = -1
    Private mRemarks As String = ""
    Private mValidUntil As EnumValidUntil = EnumValidUntil.ETD
    Private msoftwaregenerate As Integer = 0
    Private mprimarybookingGIT As Double = -1
    ''' <summary>
    ''' ini digunakan hanya booking yang dibuat secara otomatis dari booking GIT dan 
    ''' secara default bernilai -1
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumValidUntil As Integer
        ETD
        one_days_after
        two_days_after
        three_days_after
    End Enum

    Public Property ValidUntil() As EnumValidUntil
        Get
            Return mValidUntil
        End Get
        Set(ByVal value As EnumValidUntil)
            Me.mValidUntil = value
        End Set
    End Property

    Public Property softwaregenerate() As Integer
        Get
            Return msoftwaregenerate
        End Get
        Set(ByVal value As Integer)
            msoftwaregenerate = value
        End Set
    End Property

    Public Property primary() As Double
        Get
            Return Me.mprimary
        End Get
        Set(ByVal value As Double)
            Me.mprimary = value
        End Set
    End Property

    Public Property primaryCustomer() As Double
        Get
            Return Me.mPrimaryCust
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryCust = value
        End Set
    End Property

    Public Property primaryBookingGIT() As Double
        Get
            Return Me.mprimarybookingGIT
        End Get
        Set(ByVal value As Double)
            Me.mprimarybookingGIT = value
        End Set
    End Property

    Public Property primarySO() As Double
        Get
            Return Me.mPrimarySO
        End Get
        Set(ByVal value As Double)
            Me.mPrimarySO = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mDate
        End Get
        Set(ByVal value As Date)
            Me.mDate = value
        End Set
    End Property

    Public Property BookingNo() As String
        Get
            Return Me.mBookingNo
        End Get
        Set(ByVal value As String)
            Me.mBookingNo = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            Me.mRemarks = value
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

    Public Function GetJmlBySO(ByVal vPrimarySO As Double) As Integer
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlByPrimarySO(vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_MaxUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_MaxUrutan(TanggalStart, TanggalEnd).Value + 1)
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
            Return NoUrut + "/" + "BOOK" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.BookingNo) = True Then
                    Me.mBookingNo = Me.GetNoBaru(Me.TanggalTransaksi)
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika sudah melewati batas penguncian transaksi

            '4. jika tidak ada item yang akan disimpan

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
      ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
      Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit invoice:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                    Return True
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        ' Dim DTA1 As New DseTransaksiBookingInvTableAdapters.tbltrans_aopTableAdapter 

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'data SO tidak boleh dihapus jika
            '1. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
                Return True
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. Jika sudah pernah dicetak, maka minta minta otorisasi sebelum menghapus invoice
            'msgbox("belum")

            '3. Tidak bisa dihapus jika sudah dirujuk oleh invoice lain sebagai item DP
            'If DTA1.ScalarQuery_GetJmlByPrimaryInvoiceDO(vPrimaryTransaksi).Value > 0 Then
            'If TampilkanPesan = True Then
            'MsgBox("This invoice is refered as down payment by: " + _
            'CStr(DTA1.ScalarQuery_GetNoInvoice(vPrimaryTransaksi)) + ", you cannot edit thisa data.", _
            ' MsgBoxStyle.Information, "Information")
            'End If
            'Return False
            'End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'DTA1.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vnoTrans As String) As Double
        Dim dta As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimary(vnoTrans).HasValue Then
                Return dta.ScalarQuery_GetPrimary(vnoTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vPrimary)
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

    Public Function CariDetail(ByVal vPrimaryTransaksi As Double, ByVal vDgvItem As DataGridView, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim TBL As New DseTransaksiBookingInv.tbltrans_booking_inv1DataTable
        Dim row As DseTransaksiBookingInv.tbltrans_booking_inv1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row

                    Me.mprimary = .Primary
                    mBookingNo = .No_Transaksi
                    mNoUrut = .Urutan
                    mDate = ._Date
                    mPrimaryCust = .Primary_Customer
                    mPrimarySO = .Primary_SO
                    mRemarks = .Remarks
                    mValidUntil = .Valid_Until
                    msoftwaregenerate = .Software_Generated
                    mprimarybookingGIT = .Primary_Booking_GIT

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                Dim objItem As New ClsTransaksiBookingInvItem
                objItem.IsiDGVCariDetail(vDgvItem, Me.mprimary, i)
                objItem = Nothing
                'nilai balikan
                Return True
            Else
                Me.ResetObject()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub ResetObject()
        Try

            mprimary = -1
            mBookingNo = ""
            mNoUrut = -1
            mDate = Today
            mPrimaryCust = -1
            mPrimarySO = -1
            mRemarks = ""
            mValidUntil = EnumValidUntil.ETD
            msoftwaregenerate = 0
            mprimarybookingGIT = -1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimeStamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimeStamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiBookingInv.DataTable2DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglSampaiDari(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustomerCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustomerNameLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_DescriptionLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                QueryResult = DTA.FillBy_SONumber(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(TBL.Columns.IndexOf(TBL.PrimaryColumn)).Visible = False
                DGV.Columns(TBL.Columns.IndexOf(TBL.DateColumn)).DefaultCellStyle.Format = "dd MMM yyyy"
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
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional _
       ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable21TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.GetJmlBy_NoTransaksiLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.GetJmlBy_TglSampaiDari(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.GetJmlBy_ProductCodeLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.GetJmlBy_DescriptionLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.GetJmlBy_SONumber(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.GetJmlBy_CustomerCodeLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.GetJmlBy_CustomerNameLike(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
        ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.DeleteQuery(CStr(vPrimaryTransaksi)) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim ObjItem As New ClsTransaksiBookingInvItem
        Dim UrutanFakturPajak As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True) = True Then
                'insert query
                Me.mNoUrut = CInt(Me.mBookingNo.Substring(0, 6))
                If DTA.Insert(Me.mBookingNo, Me.mDate, _
                 Me.mNoUrut, Me.mPrimaryCust, Me.mPrimarySO, Me.mRemarks, vUserCode, _
                Now, vUserCode, Now, Me.mValidUntil, Me.TimeStamp, _
                 Me.msoftwaregenerate, Me.mprimarybookingGIT) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mprimary = DTA.ScalarQuery_GetPrimary(Me.mBookingNo).Value
                    Me.mTimestamp = Me.GetTimestamp(Me.mprimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mprimary, i)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mBookingNo + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mBookingNo = Me.GetNoBaru(Me.mDate)
                Return Me.SimpanBaru(vdgvItem, vUserCode, i)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    ''' <summary>
    ''' mengeset aktif semua booking git yang merujuk pada packing list yang diberikan primary-nya 
    ''' pada parameter
    ''' </summary>
    ''' <param name="vPrimaryPackingList"></param>
    ''' <remarks></remarks>
    Public Sub SetBookingGITAktif(ByVal vPrimaryPackingList As Double)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable10TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetAktifBookingGIT(vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim ObjItem As New ClsTransaksiBookingInvItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False) = True Then
                'hitung level otorisasi yang dibutuhkan
                Me.mNoUrut = Me.mBookingNo.Substring(0, 6)
                If DTA.Update(Me.mBookingNo, Me.mDate, _
                Me.mNoUrut, Me.mPrimaryCust, Me.mPrimarySO, Me.mRemarks, _
                vUserCode, Now, Me.mValidUntil, Now, _
                Me.msoftwaregenerate, Me.mprimarybookingGIT, Me.primary, Me.TimeStamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    '   Me.mNoUrut = Me.mBookingNo.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mprimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mprimary, i)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub CreateBookingInventory_DenganGIT(ByVal vPrimaryStockInPurchase As Double, _
    ByVal vPrimaryPackingList As Double, ByVal vUser As String)
        Dim TBLDummy As New DseTransaksiBookingInv.DataTable9DataTable
        Dim rowDummy As DseTransaksiBookingInv.DataTable9Row
        Dim TBLStockIn As New DseTransaksiBookingInv.tbltransaksi_stockinout_itemDataTable
        Dim DTAStockIn As New DseTransaksiBookingInvTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBLBOokingGIT As New DseTransaksiBookingInv.DataTable10DataTable
        Dim DTABookingGIT As New DseTransaksiBookingInvTableAdapters.DataTable10TableAdapter
        Dim ListPrimaryPO As New List(Of Double)
        Dim JmlSOB, JmlGITAktif, JmlBookingInv, JmlPendingDO, JmlBooking As Double
        Dim DTABookingInv As New DseTransaksiBookingInvTableAdapters.DataTable7TableAdapter
        Dim TBLBookingInv As New DseTransaksiBookingInv.DataTable7DataTable
        Dim DTASOB As New DseTransaksiBookingInvTableAdapters.tbltrans_so_item1TableAdapter
        Dim TBLSOB As New DseTransaksiBookingInv.tbltrans_so_item1DataTable
        Dim DTAPendingDO As New DseTransaksiBookingInvTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLPendingDO As New DseTransaksiBookingInv.tbltrans_do_itemDataTable
        Dim Filter As String = ""
        Dim DTABookingGIT2 As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_gititemTableAdapter
        Dim TBLBookingGIT2 As New DseTransaksiBookingInv.tbltrans_booking_gititemDataTable
        Dim TBLHeader As New DseTransaksiBookingInv.tbltrans_booking_inv3DataTable
        Dim TBLItem As New DseTransaksiBookingInv.tbltrans_booking_invitemDataTable

        Try
            DTAStockIn.Connection.ConnectionString = Core_Module.GetConstring
            DTABookingGIT.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTABookingGIT2.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTABookingInv.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTASOB.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString

            'cara kerja koding:
            '00. set status booking git menjadi non aktif
            DTABookingGIT.UpdateQuery_NonAktifkanBookingGIT(vPrimaryPackingList)
            '0. Load database yang digunakan untuk perhitungan
            DTABookingInv.Fill(TBLBookingInv)
            DTASOB.Fill(TBLSOB)
            DTAPendingDO.Fill(TBLPendingDO)
            DTABookingGIT2.FillBy_PrimaryPackingList_Exclude(TBLBookingGIT2, vPrimaryPackingList)
            '1. Kumpulkan semua booking GIT yang merujuk ke dalam packing list sumber dari stock in ini
            DTABookingGIT.FillBy_PrimaryPackingList(TBLBOokingGIT, vPrimaryPackingList)
            '2. Kumpulkan semua item stock in purchase
            DTAStockIn.FillBy_PrimaryStockIn(TBLStockIn, vPrimaryStockInPurchase)
            '3.a. Kombinasikan item booking GIT dan item stock in untuk mencari item booking git mana saja 
            '   yang menjadi calon booking inv. Tabel ini dimasukkan mulai dari primary booking paling kecil dan 
            '   dibatasi oleh jumlah stock in dan jumlah booking git (mana yang paling kecil)
            Me.GenerateItemDummy(TBLBOokingGIT, TBLStockIn, TBLDummy)
            '4. Kumpulkan item po, sesuai dengan primary po yang ada didalam tabel dummy
            For Each rowDummy In TBLDummy.Rows
                Filter = "[Primary SO] =" + rowDummy.Primary_SO.ToString + _
                " AND [Primary Product] =" + rowDummy.Primary_Product.ToString + ""
                '5. untuk setiap item dummy yang ada, cari jumlah yang harus dibuatkan booking inv, jumlah ini adalah
                '   jumlah minimal dari:
                JmlBooking = rowDummy.Qty
                '       -. jmlsob-jmlpendingdo-jmlbookinginv-jmlbookinggitaktif
                JmlSOB = Me.GetJmlSOB(Filter, TBLSOB)
                JmlGITAktif = Me.GetJmlBookingGITAktif(Filter, TBLBookingGIT2)
                JmlBookingInv = Me.GetJmlBookingInv(Filter, TBLBookingInv)
                JmlPendingDO = Me.GetJmlPendingDO(Filter, TBLPendingDO)
                JmlSOB = JmlSOB - JmlGITAktif - JmlBookingInv - JmlPendingDO
                If JmlBooking > JmlSOB Then
                    JmlBooking = JmlSOB
                End If
                '       (pembatasan jumlah oleh stock in dan booking git sudah dilakukan ketika membuat item dummy ini)
                If JmlBooking <> 0 Then
                    '6. men-generate item booking dan headernya
                    Me.GenerateBookingInv2(TBLHeader, TBLItem, rowDummy, JmlBooking, vUser)
                End If
            Next
            '7. simpan data booking
            Me.SimpanBookingInv(TBLHeader, TBLItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy.Dispose()
            TBLStockIn.Dispose()
            DTAStockIn.Dispose()
            TBLBOokingGIT.Dispose()
            DTABookingGIT.Dispose()
            ListPrimaryPO = Nothing
            DTABookingInv.Dispose()
            TBLBookingInv.Dispose()
            DTASOB.Dispose()
            TBLSOB.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            DTABookingGIT2.Dispose()
            TBLBookingGIT2.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub GenerateItemDummy(ByVal vTBLBookingGIT As DseTransaksiBookingInv.DataTable10DataTable, _
    ByVal vTBLStockInPurchase As DseTransaksiBookingInv.tbltransaksi_stockinout_itemDataTable, _
    ByVal vTBLDummy As DseTransaksiBookingInv.DataTable9DataTable)
        Dim rowStockIn As DseTransaksiBookingInv.tbltransaksi_stockinout_itemRow
        Dim rowBookingGIT As DseTransaksiBookingInv.DataTable10Row
        Dim rowDummy As DseTransaksiBookingInv.DataTable9Row
        Dim QtyDummy As Double = 0
        Try
            'untuk setiap baris di tabel booking inv, selama masih ada baris stock in yang Primary Productnya sesuai
            'masukkan ke dalam item dummy
            For Each rowBookingGIT In vTBLBookingGIT.Rows
                For Each rowStockIn In vTBLStockInPurchase.Rows
                    If rowBookingGIT.Qty <> 0 Then
                        If rowStockIn.Qty <> 0 Then
                            rowDummy = vTBLDummy.NewDataTable9Row
                            rowDummy.Batch_No = rowStockIn.Batch_No
                            rowDummy.Delivery_Date = rowBookingGIT.Delivery_Date
                            rowDummy.Primary_Booking_GIT = rowBookingGIT.Primary
                            rowDummy.Primary_Customer = rowBookingGIT.Primary_Customer
                            rowDummy.Primary_IPO = rowBookingGIT.Primary_IPO
                            rowDummy.Primary_Product = rowBookingGIT.Primary_Product
                            rowDummy.Primary_SO = rowBookingGIT.PrimarySO
                            rowDummy.Primary_Storage_Area = rowStockIn.Primary_Storage_Area
                            rowDummy.Primary_Secondary_Units = rowStockIn.Primary_Secondary_Units
                            rowDummy.Primary_Units = rowStockIn.Primary_Units

                            If rowStockIn.Qty > rowBookingGIT.Qty Then
                                QtyDummy = rowBookingGIT.Qty
                                rowBookingGIT.Qty = 0
                                rowStockIn.Qty -= QtyDummy
                            Else
                                QtyDummy = rowStockIn.Qty
                                rowStockIn.Qty = 0
                                rowBookingGIT.Qty -= QtyDummy
                            End If
                            rowDummy.Qty = QtyDummy
                            rowDummy.Sec_Qty = rowStockIn.Secondary_Qty
                            rowDummy.Expiry_Date = rowStockIn.Expiry_Date

                            vTBLDummy.Rows.Add(rowDummy)
                        End If
                    Else
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CreateBookingInventory_TanpaGIT(ByVal vPrimaryStockInPurchase As Double, ByVal vUser As String)
        Dim DTAPO As New DseTransaksiBookingInvTableAdapters.tbltransaksi_stockinout_item1TableAdapter
        Dim TBLPO As New DseTransaksiBookingInv.tbltransaksi_stockinout_item1DataTable
        Dim rowPO As DseTransaksiBookingInv.tbltransaksi_stockinout_item1Row
        Dim DTAStockIn As New DseTransaksiBookingInvTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBLStockIn As New DseTransaksiBookingInv.tbltransaksi_stockinout_itemDataTable
        Dim rowStockIn As DseTransaksiBookingInv.tbltransaksi_stockinout_itemRow
        Dim DTAPOItem As New DseTransaksiBookingInvTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiBookingInv.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiBookingInv.tbltrans_po_itemRow
        Dim DTABookingGIT As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_gititemTableAdapter
        Dim TBLBookingGIT As New DseTransaksiBookingInv.tbltrans_booking_gititemDataTable
        Dim DTABookingInv As New DseTransaksiBookingInvTableAdapters.DataTable7TableAdapter
        Dim TBLBookingInv As New DseTransaksiBookingInv.DataTable7DataTable
        Dim DTASOB As New DseTransaksiBookingInvTableAdapters.tbltrans_so_item1TableAdapter
        Dim TBLSOB As New DseTransaksiBookingInv.tbltrans_so_item1DataTable
        Dim DTAPendingDO As New DseTransaksiBookingInvTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLPendingDO As New DseTransaksiBookingInv.tbltrans_do_itemDataTable
        Dim DTABookingInv2 As New DseTransaksiBookingInvTableAdapters.DataTable8TableAdapter
        Dim TBLBookingInv2 As New DseTransaksiBookingInv.DataTable8DataTable

        Dim JmlSOB, JmlGITAktif, JmlBookingInv, JmlPendingDO, JmlBooking As Double
        Dim Filter As String = ""

        Dim TBLHeader As New DseTransaksiBookingInv.tbltrans_booking_inv3DataTable
        Dim TBLItem As New DseTransaksiBookingInv.tbltrans_booking_invitemDataTable

        Try
            DTAStockIn.Connection.ConnectionString = Core_Module.GetConstring
            DTABookingGIT.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTABookingInv2.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTAPO.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTAPOItem.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTABookingInv.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTASOB.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTAStockIn.Connection.ConnectionString

            'cara kerja koding:
            '1. Kumpulkan semua primary po pada item stock in purchase
            '2. Untuk setiap primary po, kumpulkan item stock in yang merujuk pada primary ponya
            '3. untuk setiap item stock in yang ada, cari jumlah yang harus dibuatkan booking inv, jumlah ini adalah
            '   jumlah minimal dari:
            '       -. jmlstockin
            '       -. jmlsob-jmlpendingdo-jmlbookinginv-jmlbookinggitaktif
            '       -. jmlipo yang dipesan pada po tersebut - yang sudah terkirim(untuk ipo ybs) - 
            '          yang sudah dibooking (untuk ipo ybs)
            '4. men-generate item booking dan headernya
            '5. simpan data booking

            'isikan data yang akan digunakan untuk perhitungan
            DTABookingGIT.Fill(TBLBookingGIT)
            DTABookingInv.Fill(TBLBookingInv)
            DTASOB.Fill(TBLSOB)
            DTAPendingDO.Fill(TBLPendingDO)
            DTABookingInv2.Fill(TBLBookingInv2)

            'untuk setiap PO yang ada didalam stock in purchase:
            DTAPO.Fill(TBLPO, vPrimaryStockInPurchase)
            For Each rowPO In TBLPO.Rows
                'kumpulkan item stock in purchase yang merujuk pada po tersebut
                DTAStockIn.FillBy_PrimaryPO_PrimaryStockIn(TBLStockIn, rowPO.Primary_PO_Purchase, vPrimaryStockInPurchase)
                'kumpulkan item PO yang berasal dari IPO berdasarkan primary PO dan Primary Product
                'urutan mulai dari IPO yang paling pertama dimasukkan 
                '--> kemudian posisikan berdasarkan purchase orderbalance, diasumsikan barang yang masuk duluan, digunakan
                'untuk memenuhi IPO yang pertama
                DTAPOItem.FillBy_PrimaryPO(TBLPOItem, rowPO.Primary_PO_Purchase)
                Me.ModifikasiTBLPOItem(TBLPOItem, TBLStockIn, TBLBookingInv2)

                'mulai scanning dari baris pertama dari tabel stock in purhchase
                For Each rowStockIn In TBLStockIn.Rows
                    For Each rowPOItem In TBLPOItem.Rows
                        Filter = "[Primary SO] =" + rowPOItem.Primary_SO.ToString + _
                        " AND [Primary Product] =" + rowPOItem.Primary_Product.ToString + ""
                        If rowPOItem.Qty <> 0 And rowPOItem.Primary_Product = rowStockIn.Primary_Product Then
                            'jumlah yang masuk ke dalam booking adalah nilai minimum dari salah 1 dari 3 paramater dibawah:
                            '1. qty stock in purchase
                            JmlBooking = rowStockIn.Qty
                            '2. qty pemesanan ipo pada item po tsb
                            If JmlBooking > rowPOItem.Qty Then
                                JmlBooking = rowPOItem.Qty
                            End If
                            '3. qty sob - qty booking git aktif - qty booking lainnya - qty pending do
                            JmlSOB = Me.GetJmlSOB(Filter, TBLSOB)
                            JmlGITAktif = Me.GetJmlBookingGITAktif(Filter, TBLBookingGIT)
                            JmlBookingInv = Me.GetJmlBookingInv(Filter, TBLBookingInv)
                            JmlPendingDO = Me.GetJmlPendingDO(Filter, TBLPendingDO)
                            JmlSOB = JmlSOB - JmlGITAktif - JmlBookingInv - JmlPendingDO
                            If JmlBooking > JmlSOB Then
                                JmlBooking = JmlSOB
                            End If
                            'masukkan item booking sementara ke dalam satu tabel sementara
                            If JmlBooking <> 0 Then
                                Me.GenerateBookingInv(TBLHeader, TBLItem, rowStockIn, JmlBooking, rowPOItem.Primary_SO, _
                                rowPOItem.Primary_Customer, vUser, rowPOItem.Delivery_Date, rowPOItem.Primary_IPO)
                                'set nilai baru qty pada tbl stock in
                                rowStockIn.Qty -= JmlBooking
                                'set nilai baru qty pada tbl po item
                                rowPOItem.Qty -= JmlBooking
                                'tambahkan nilai qty booking inv yang terjadi pada tabel sementara
                                Me.TambahBookingInv(rowStockIn.Primary_Product, rowPOItem.Primary_SO, JmlBooking, _
                                TBLBookingInv, Filter)
                                Me.TambahBookingInv(rowStockIn.Primary_Product, rowPOItem.Primary_SO, JmlBooking, _
                                TBLBookingInv2, Filter, rowPOItem.Primary_IPO)
                            End If
                            'evaluasi, jika stock in tidak ada yang bisa dibooking maka
                            'keluar dari looping pada tbl po item, pindah ke bari selanjutnya pada
                            'tabel stock in purchase
                            If rowStockIn.Qty = 0 Then
                                Exit For
                            End If
                        End If
                    Next
                Next
            Next
            'simpan data booking inventory ke dalam basis data
            Me.SimpanBookingInv(TBLHeader, TBLItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPO.Dispose()
            TBLPO.Dispose()
            DTAStockIn.Dispose()
            TBLStockIn.Dispose()
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            DTABookingGIT.Dispose()
            TBLBookingGIT.Dispose()
            DTABookingInv.Dispose()
            TBLBookingInv.Dispose()
            DTASOB.Dispose()
            TBLSOB.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
            DTABookingInv2.Dispose()
            TBLBookingInv2.Dispose()
        End Try
    End Sub


    ''' <summary>
    ''' melakukan modifikasi pada tabel po item, asumsi yang digunakan booking dilakukan pada item2
    ''' IPO mulai dari yang paling pertama dimasukkan
    ''' </summary>
    ''' <param name="vTblPOItem"></param>
    ''' <remarks></remarks>
    Private Sub ModifikasiTBLPOItem(ByVal vTblPOItem As DseTransaksiBookingInv.tbltrans_po_itemDataTable, _
    ByVal vTblStockIn As DseTransaksiBookingInv.tbltransaksi_stockinout_itemDataTable, _
    ByVal vTblBookingInv2 As DseTransaksiBookingInv.DataTable8DataTable)
        Dim JmlTotal As Double = 0
        Dim rowPOItem As DseTransaksiBookingInv.tbltrans_po_itemRow
        Dim PrimaryProduct, counter, JmlTerkirim, POB, JmlStockIn, JmlBookingInvAktif, Dummy As Double
        Dim vListPrimaryProduct As New List(Of Double)

        Try
            'kumpulkan primary product yang ada
            For Each rowPOItem In vTblPOItem.Rows
                If vListPrimaryProduct.Contains(rowPOItem.Primary_Product) = False Then
                    vListPrimaryProduct.Add(rowPOItem.Primary_Product)
                End If
            Next
            'untuk setiap primary product 
            For Each PrimaryProduct In vListPrimaryProduct
                'hitung jumlah qty pesanan pada PO dan POB(Purchase Order Balance) untuk product tersebut
                'sekaligus membuang baris2 yang bukan berasal dari ipo (tidak perlu booking)
                counter = 0
                JmlTotal = 0
                While counter < vTblPOItem.Rows.Count
                    rowPOItem = vTblPOItem.Rows(counter)
                    If rowPOItem.Primary_Product = PrimaryProduct Then
                        JmlTotal += rowPOItem.Qty
                        POB = rowPOItem.POB
                        If rowPOItem.Primary_IPO = -1 Then
                            vTblPOItem.Rows.RemoveAt(counter)
                        Else
                            counter += 1
                        End If
                    Else
                        counter += 1
                    End If
                End While
                JmlStockIn = Me.GetJmlStockIn(vTblStockIn, PrimaryProduct)
                'hitung jumlah terkirim
                JmlTerkirim = JmlTotal - POB - JmlStockIn

                counter = 0
                While counter < vTblPOItem.Rows.Count
                    rowPOItem = vTblPOItem.Rows(counter)
                    If rowPOItem.Primary_Product = PrimaryProduct Then
                        'rowPOItem.Qty -= JmlBookingInvAktif
                        If rowPOItem.Qty > JmlTerkirim Then
                            'jika ada booking inv yang aktif maka ini juga akan menngurangi jumlah item po yang
                            'bisa dibooking
                            Dummy = rowPOItem.Qty - JmlTerkirim
                            JmlBookingInvAktif = Me.GetJmlBookingInv2(PrimaryProduct, _
                            rowPOItem.Primary_IPO, vTblBookingInv2)
                            If Dummy + JmlBookingInvAktif <= rowPOItem.Qty Then
                                rowPOItem.Qty = Dummy
                            Else
                                rowPOItem.Qty = rowPOItem.Qty - JmlTerkirim
                                If rowPOItem.Qty > JmlBookingInvAktif Then
                                    'masih ada ruangan untuk melakukan booking
                                    rowPOItem.Qty -= JmlBookingInvAktif
                                Else
                                    'booking yang aktif saat ini masih lebih banyak dari pada pemesanan ipo
                                    rowPOItem.Qty = 0
                                End If
                            End If
                            'jumlah terkirim = 0
                            JmlTerkirim = 0
                        Else
                            JmlTerkirim -= rowPOItem.Qty
                            rowPOItem.Qty = 0
                        End If
                        If rowPOItem.Qty = 0 Then
                            vTblPOItem.Rows.RemoveAt(counter)
                        Else
                            counter += 1
                        End If
                    Else
                        counter += 1
                    End If
                End While
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vListPrimaryProduct = Nothing
        End Try
    End Sub

    Private Function GetJmlBookingInv2(ByVal vPrimaryProduct As Double, ByVal vPrimaryIPO As Double, _
    ByVal vTBLBooking As DseTransaksiBookingInv.DataTable8DataTable) As Double
        Dim rowBooking As DseTransaksiBookingInv.DataTable8Row
        Dim retval As Double
        Try
            For Each rowBooking In vTBLBooking.Rows
                If rowBooking.Primary_IPO = vPrimaryIPO And rowBooking.Primary_Product = vPrimaryProduct Then
                    retval += rowBooking.QtyBookingInv
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlStockIn(ByVal vTBL As _
    DseTransaksiBookingInv.tbltransaksi_stockinout_itemDataTable, ByVal vPrimaryProduct As Double) As Double
        Dim row As DseTransaksiBookingInv.tbltransaksi_stockinout_itemRow
        Dim retval As Double

        Try
            For Each row In vTBL.Rows
                If row.Primary_Product = vPrimaryProduct Then
                    retval += row.Qty
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TambahBookingInv(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal vQtyBooking As Double, ByVal vTBLBoooking As DseTransaksiBookingInv.DataTable7DataTable, _
    ByVal vFilter As String)
        Dim rowBooking As DseTransaksiBookingInv.DataTable7Row

        Try
            vTBLBoooking.DefaultView.RowFilter = vFilter
            If vTBLBoooking.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowBooking = vTBLBoooking.NewDataTable7Row
                rowBooking.Primary_Product = vPrimaryProduct
                rowBooking.Primary_SO = vPrimarySO
                rowBooking.QtyBookingInv = vQtyBooking
                vTBLBoooking.Rows.Add(rowBooking)
            Else
                'tambahkan qty booking
                rowBooking = vTBLBoooking.DefaultView(0).Row
                rowBooking.QtyBookingInv += vQtyBooking
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahBookingInv(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal vQtyBooking As Double, ByVal vTBLBoooking As DseTransaksiBookingInv.DataTable8DataTable, _
    ByVal vFilter As String, ByVal vPrimaryIPO As Double)
        Dim rowBooking As DseTransaksiBookingInv.DataTable8Row
        Dim Dummy As String = vFilter + "AND [Primary IPO] =" + vPrimaryIPO.ToString

        Try
            vTBLBoooking.DefaultView.RowFilter = Dummy
            If vTBLBoooking.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowBooking = vTBLBoooking.NewDataTable8Row
                rowBooking.Primary_Product = vPrimaryProduct
                rowBooking.Primary_SO = vPrimarySO
                rowBooking.QtyBookingInv = vQtyBooking
                rowBooking.Primary_IPO = vPrimaryIPO
                vTBLBoooking.Rows.Add(rowBooking)
            Else
                'tambahkan qty booking
                rowBooking = vTBLBoooking.DefaultView(0).Row
                rowBooking.QtyBookingInv += vQtyBooking
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menyimpan ke database booking inv yang ada didalam tabel header dan tabel item
    ''' </summary>
    ''' <param name="vTBLHeader"></param>
    ''' <param name="vTblItem"></param>
    ''' <remarks></remarks>
    Private Sub SimpanBookingInv(ByVal vTBLHeader As DseTransaksiBookingInv.tbltrans_booking_inv3DataTable, _
    ByVal vTblItem As DseTransaksiBookingInv.tbltrans_booking_invitemDataTable)
        Dim DTAHeader As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv3TableAdapter
        Dim DTAItem As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim rowHeader As DseTransaksiBookingInv.tbltrans_booking_inv3Row
        Dim rowItem As DseTransaksiBookingInv.tbltrans_booking_invitemRow
        Dim PrimaryBaru As Double = -1

        Try
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTAHeader.Connection.ConnectionString

            For Each rowHeader In vTBLHeader.Rows
                With rowHeader
                    'cari nomor baru
                    rowHeader.No_Transaksi = Me.GetNoBaru(rowHeader._Date)
                    'simpan header ke basis data
                    DTAHeader.Insert(.No_Transaksi, .No_Transaksi.Substring(0, 6), ._Date, .Primary_Customer, .Primary_SO, _
                    .Remarks, .Created_By, .Created_Date, .Modified_By, .Modified_Date, .Timestamp, .Valid_Until, 1, rowHeader.Primary_Booking_GIT)
                    'cari primary baru
                    PrimaryBaru = Me.GetPrimary(.No_Transaksi)
                End With

                'simpan item ke basis data
                For Each rowItem In vTblItem.Rows
                    If rowItem.Primary_Booking = rowHeader.Primary Then
                        With rowItem
                            DTAItem.Insert(PrimaryBaru, .Primary_Product, .Qty, .Primary_Unit, .Batch_No, _
                            .Expiry_Date, .Remarks, .Sec_Qty, .Primary_Sec_Unit, .Valid_Until, .Primary_Location, _
                            .Primary_IPO)
                        End With
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAHeader.Dispose()
            DTAItem.Dispose()
        End Try
    End Sub

    Private Sub GenerateBookingInv2(ByVal vTBLHeader As DseTransaksiBookingInv.tbltrans_booking_inv3DataTable, _
    ByVal vTblItem As DseTransaksiBookingInv.tbltrans_booking_invitemDataTable, _
    ByVal rowDummy As DseTransaksiBookingInv.DataTable9Row, ByVal JmlBooking As Double, _
    ByVal vUser As String)
        Dim Index As Integer
        Dim rowHeader As DseTransaksiBookingInv.tbltrans_booking_inv3Row
        Dim rowItem As DseTransaksiBookingInv.tbltrans_booking_invitemRow
        Dim Timestamp As DateTime = Now
        Dim CreatedBy As String = ""

        Try
            Index = Me.GetPrimary(rowDummy.Primary_SO, vTBLHeader)
            If Index = -1 Then
                Dim ObjSO As New ClsTransaksiSalesOrder

                Index = vTBLHeader.Rows.Count + 1
                rowHeader = vTBLHeader.Newtbltrans_booking_inv3Row
                rowHeader.Primary = Index
                CreatedBy = ObjSO.Get_CreatedBy(rowDummy.Primary_SO)
                rowHeader.Created_By = CreatedBy
                rowHeader.Created_Date = Timestamp
                rowHeader.Modified_By = CreatedBy
                rowHeader.Modified_Date = Timestamp
                rowHeader.No_Transaksi = ""
                rowHeader.Primary_Booking_GIT = rowDummy.Primary_Booking_GIT
                rowHeader.Primary_Customer = rowDummy.Primary_Customer
                rowHeader.Primary_SO = rowDummy.Primary_SO
                rowHeader.Timestamp = Timestamp
                rowHeader.Urutan = 0
                'default 3 hari booking
                rowHeader.Valid_Until = 3
                rowHeader._Date = Today
                rowHeader.Remarks = ""
                rowHeader.Software_Generated = 1
                vTBLHeader.Rows.Add(rowHeader)

                ObjSO = Nothing
            End If

            'tambahkan item booking inv
            rowItem = vTblItem.Newtbltrans_booking_invitemRow
            With rowItem
                .Primary = vTblItem.Rows.Count + 1
                .Batch_No = rowDummy.Batch_No
                .Expiry_Date = rowDummy.Expiry_Date
                .Primary_Booking = Index
                .Primary_Location = rowDummy.Primary_Storage_Area
                .Primary_Product = rowDummy.Primary_Product
                .Primary_Sec_Unit = rowDummy.Primary_Secondary_Units
                .Primary_Unit = rowDummy.Primary_Units
                .Qty = JmlBooking
                .Remarks = ""
                .Sec_Qty = rowDummy.Sec_Qty
                If rowDummy.Delivery_Date <> #1/1/1950# Then
                    .Valid_Until = rowDummy.Delivery_Date
                Else
                    .Valid_Until = Today + TimeSpan.FromDays(3)
                End If
                .Primary_IPO = -1
            End With
            vTblItem.Rows.Add(rowItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menambahkan item booking dari baris stock in purchase sejumlah yang dibutuhkan
    ''' , dikelompokan berdasarkan primary SO
    ''' </summary>
    ''' <param name="vTBLHeader"></param>
    ''' <param name="vTblItem"></param>
    ''' <param name="rowStockIn"></param>
    ''' <param name="JmlBooking"></param>
    ''' <remarks></remarks>
    Private Sub GenerateBookingInv(ByVal vTBLHeader As DseTransaksiBookingInv.tbltrans_booking_inv3DataTable, _
    ByVal vTblItem As DseTransaksiBookingInv.tbltrans_booking_invitemDataTable, _
    ByVal rowStockIn As DseTransaksiBookingInv.tbltransaksi_stockinout_itemRow, ByVal JmlBooking As Double, _
    ByVal vPrimarySO As Double, ByVal vPrimaryCust As Double, ByVal vUser As String, _
    ByVal vEstimatedTimeOfDelivery As Date, ByVal vPrimaryIPO As Double)
        Dim Index As Integer
        Dim rowHeader As DseTransaksiBookingInv.tbltrans_booking_inv3Row
        Dim rowItem As DseTransaksiBookingInv.tbltrans_booking_invitemRow
        Dim Timestamp As DateTime = Now
        Dim CreatedBy As String = ""

        Try
            Index = Me.GetPrimary(vPrimarySO, vTBLHeader)
            If Index = -1 Then
                Dim ObjSO As New ClsTransaksiSalesOrder

                Index = vTBLHeader.Rows.Count + 1
                rowHeader = vTBLHeader.Newtbltrans_booking_inv3Row
                rowHeader.Primary = Index
                rowHeader.Created_Date = Timestamp
                rowHeader.Modified_Date = Timestamp
                rowHeader.No_Transaksi = ""
                rowHeader.Primary_Booking_GIT = -1
                rowHeader.Primary_Customer = vPrimaryCust
                rowHeader.Primary_SO = vPrimarySO
                CreatedBy = ObjSO.Get_CreatedBy(vPrimarySO)
                rowHeader.Created_By = CreatedBy
                rowHeader.Modified_By = CreatedBy
                rowHeader.Timestamp = Timestamp
                rowHeader.Urutan = 0
                'default 3 hari booking
                rowHeader.Valid_Until = 3
                rowHeader._Date = Today
                rowHeader.Remarks = ""
                rowHeader.Software_Generated = 1
                vTBLHeader.Rows.Add(rowHeader)

                ObjSO = Nothing
            End If

            'tambahkan item booking inv
            rowItem = vTblItem.Newtbltrans_booking_invitemRow
            With rowItem
                .Primary = vTblItem.Rows.Count + 1
                .Batch_No = rowStockIn.Batch_No
                .Expiry_Date = rowStockIn.Expiry_Date
                .Primary_Booking = Index
                .Primary_Location = rowStockIn.Primary_Storage_Area
                .Primary_Product = rowStockIn.Primary_Product
                .Primary_Sec_Unit = rowStockIn.Primary_Secondary_Units
                .Primary_Unit = rowStockIn.Primary_Units
                .Qty = JmlBooking
                .Remarks = ""
                .Sec_Qty = rowStockIn.Secondary_Qty
                If vEstimatedTimeOfDelivery > Today Then
                    .Valid_Until = vEstimatedTimeOfDelivery
                Else
                    .Valid_Until = Today + TimeSpan.FromDays(3)
                End If
                .Primary_IPO = vPrimaryIPO
            End With
            vTblItem.Rows.Add(rowItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimary(ByVal vPrimarySO As Double, _
    ByVal vTBLHeader As DseTransaksiBookingInv.tbltrans_booking_inv3DataTable) As Integer
        Dim row As DseTransaksiBookingInv.tbltrans_booking_inv3Row

        Try
            For Each row In vTBLHeader.Rows
                If row.Primary_SO = vPrimarySO Then
                    Return row.Primary
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlBookingInv(ByVal vFilter As String, _
    ByVal vTBL As DseTransaksiBookingInv.DataTable7DataTable) As Double
        Dim row As DseTransaksiBookingInv.DataTable7Row

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
    ByVal vTBL As DseTransaksiBookingInv.tbltrans_booking_gititemDataTable) As Double
        Dim row As DseTransaksiBookingInv.tbltrans_booking_gititemRow

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
    ByVal vTBL As DseTransaksiBookingInv.tbltrans_do_itemDataTable) As Double
        Dim row As DseTransaksiBookingInv.tbltrans_do_itemRow

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
    ByVal vTBL As DseTransaksiBookingInv.tbltrans_so_item1DataTable) As Double
        Dim row As DseTransaksiBookingInv.tbltrans_so_item1Row

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

    ''' <summary>
    ''' menghapuskan booking yang sudah dikirimkan pada DO
    ''' </summary>
    ''' <param name="vDGVDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HapusBooking(ByVal vDGVDO As DataGridView, ByVal ij As FrmTransaksiDO3) As Boolean
        Dim row As DataGridViewRow
        Dim TBL As New DseTransaksiBookingInv.tbltrans_booking_invitem_view2DataTable
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem_view2TableAdapter
        Dim rowBooking As DseTransaksiBookingInv.tbltrans_booking_invitem_view2Row
        Dim DTA1 As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim DTA2 As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim TBL3 As New DseTransaksiBookingInv.tbltrans_booking_inv_view0DataTable
        Dim DTA3 As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv_view0TableAdapter
        Dim row3 As DseTransaksiBookingInv.tbltrans_booking_inv_view0Row

        Dim Filter As String = ""
        Dim i As Double
        Dim Lanjut As Boolean = False
        Dim tambah_i As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.ClearBeforeFill = True
            i = 0
            'untuk setiap item DO
            While i <= vDGVDO.Rows.Count - 1
                row = vDGVDO.Rows(i)
                tambah_i = True

                'isikan tabel booking sesuai dengan no SO dan product code, berurutan 
                'berdasarkan created date dari transaksi, dan primary item
                If DTA.FillBy_PrimarySO_PrimaryProduct(TBL, row.Cells(ij.ColPrimarySO.Index).Value, _
                row.Cells(ij.ColPrimaryCodeNo.Index).Value) > 0 Then
                    'bentuk filter
                    Filter = ""
                    Lanjut = True
                    If row.Cells(6).Value <> "" Then
                        Filter = "[Batch No] ='" + CStr(row.Cells(6).Value) + "'"
                    End If
                    If IsNumeric(row.Cells(14).Value) = True Then
                        If row.Cells(14).Value <> -1 Then
                            If Filter = "" Then
                                Filter = "[Sec Qty] =" + CStr(row.Cells(6).Value) + ""
                            Else
                                Filter = Filter + " AND [Sec Qty] =" + CStr(row.Cells(6).Value) + ""
                            End If
                        End If
                    End If

                    'kosongkan booking yang sesuai dengan filter terlebih dahulu
                    If Filter <> "" Then
                        Lanjut = True
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count > 0 Then
                            While Lanjut = True And TBL.DefaultView.Count > 0
                                rowBooking = TBL.DefaultView(0).Row
                                If rowBooking.Qty < row.Cells(12).Value Then
                                    'kurangkan jumlah pada dgv
                                    row.Cells(12).Value = CDbl(row.Cells(12).Value) - rowBooking.Qty
                                    'qty booking kurang dari yang dikirimkan, hapus baris booking
                                    DTA1.DeleteQuery_ByPrimaryItem(rowBooking.Primary)
                                    TBL.Rows.Remove(rowBooking)
                                Else
                                    'qty booking lebih dari atau sama dengan qty DO, proses hapus
                                    'booking untuk item yang bersangkutan tidak dilanjutkan
                                    Lanjut = False
                                    If row.Cells(12).Value = rowBooking.Qty Then
                                        'jika qty booking = qty do sisa maka hapus booking
                                        DTA1.DeleteQuery_ByPrimaryItem(rowBooking.Primary)
                                        TBL.Rows.Remove(rowBooking)
                                    Else
                                        'jika qty do kurang dari qty booking maka kurangkan data booking
                                        rowBooking.Qty = rowBooking.Qty - CDbl(row.Cells(12).Value)
                                        'update data di database
                                        DTA1.UpdateQuery_SetQtyByPrimary(rowBooking.Qty, rowBooking.Primary)
                                    End If
                                    'hapuskan item do yang perlu dikurangkan dari dgv
                                    vDGVDO.Rows.Remove(row)
                                    tambah_i = False
                                End If
                            End While
                        End If
                    End If

                    'jika masih ada sisanya, lanjutkan dengan sisa booking yang bisa dihapuskan
                    If Lanjut = True Then
                        'jika filter tidak menemukan maka booking sesuai dengan kriteria item do
                        'maka item booking akan dihapuskan dengan urutan:
                        '1. Yang booking duluan
                        '2. yang batch numbernya paling deket dengan Z
                        '3. yang sec qty-nya paling besar
                        'jika cocok maka kurangkan  qty sejumlah yang dikirim atau sejumlah
                        'yang dibooking (mana yang lebih kecil)
                        While Lanjut = True And TBL.Rows.Count > 0
                            rowBooking = TBL.Rows(0)
                            If rowBooking.Qty < row.Cells(12).Value Then
                                'kurangkan jumlah pada dgv
                                row.Cells(12).Value = CDbl(row.Cells(12).Value) - rowBooking.Qty
                                'qty booking kurang dari yang dikirimkan, hapus baris booking
                                DTA1.DeleteQuery_ByPrimaryItem(rowBooking.Primary)
                                TBL.Rows.Remove(rowBooking)
                            Else
                                'qty booking lebih dari atau sama dengan qty DO, proses hapus
                                'booking untuk item yang bersangkutan tidak dilanjutkan
                                Lanjut = False
                                If row.Cells(12).Value = rowBooking.Qty Then
                                    'jika qty booking = qty do sisa maka hapus booking
                                    DTA1.DeleteQuery_ByPrimaryItem(rowBooking.Primary)
                                    TBL.Rows.Remove(rowBooking)
                                Else
                                    rowBooking.Qty = rowBooking.Qty - CDbl(row.Cells(12).Value)
                                    'update data di database
                                    DTA1.UpdateQuery_SetQtyByPrimary(rowBooking.Qty, rowBooking.Primary)
                                End If
                                'hapuskan item do yang perlu dikurangkan dari dgv
                                vDGVDO.Rows.Remove(row)
                                tambah_i = False
                            End If
                        End While
                    End If
                End If

                'catat perubahan yang akan dilakukan, no transaksi booking, dan user yang melakukan booking
                'filter bardasarkan yang ada di dgv DO
                'MsgBox("belum")
                'kirimkan email ke pada yang melakukan booking

                'pindah baris ke selanjutnya
                If tambah_i = True Then
                    i += 1
                End If
            End While

            'hapus semua booking inventory yang sudah tidak memiliki item lagi 
            DTA3.FillBy_BookingNol(TBL3)
            For Each row3 In TBL3.Rows
                DTA2.DeleteQuery(row3.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            TBL3.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
