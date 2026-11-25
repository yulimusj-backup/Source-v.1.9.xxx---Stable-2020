Option Explicit On
Option Strict Off
Public Class ClsTransaksiBookingInv
    Dim ObjCD As New Core_Module.ClsCekData

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
            Return CInt(DTA.ScalarQuery_GetJmlByPrimarySO(vPrimarySO).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
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
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_MaxUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_MaxUrutan(TanggalStart, TanggalEnd).Value + 1)
                If CDbl(NoUrut) <= 0 Then
                    NoUrut = "1"
                End If
            Else
                NoUrut = "1"
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
        Dim ObjUser As New Core_Module.ClsUser

        Try
            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit invoice:
            '0. Jika software generated maka tidak boleh edit
            If DTA.ScalarQuery_GetSoftwareGenerated(vPrimaryTransaksi).HasValue Then
                If DTA.ScalarQuery_GetSoftwareGenerated(vPrimaryTransaksi).Value = 1 Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction is software generated, you cannot edit this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'cek apakah masih ada booking inventory ybs
            If Me.ApakahSudahAda(Me.GetNoTransaksi(vPrimaryTransaksi)) = False Then
                MsgBox("Invalid Booking inventory data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiBookingInv_Override_Edit") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        ' Dim DTA1 As New DseTransaksiBookingInvTableAdapters.tbltrans_aopTableAdapter 
        Dim ObjUser As New Core_Module.ClsUser

        Try
            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'cek apakah masih ada booking inventory ybs
            If Me.ApakahSudahAda(Me.GetNoTransaksi(vPrimaryTransaksi)) = False Then
                MsgBox("Invalid Booking inventory data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'data SO tidak boleh dihapus jika
            '1. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiBookingInv_Override_Delete") = False Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
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
            ObjUser = Nothing
        End Try
    End Function

    Public Function GetPrimary(ByVal vnoTrans As String) As Double
        Dim dta As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Try
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

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

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
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_Specification(TBL, Criteria + "%")
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
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional _
       ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable21TableAdapter
        Try
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
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_SpecificationLike(Criteria + "%").Value

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

                    'cek berkaitan nomor sama untuk beda bulan
                    Me.CekNomor(DTA)

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

    Private Sub CekNomor(ByVal DTA As DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mNoUrut
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mDate)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mDate)

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
                    If PrimaryMax = Me.mprimary Then
                        Me.mBookingNo = Me.GetNoBaru(Me.mDate)
                        Me.mNoUrut = CInt(Me.mBookingNo.Substring(0, 6))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mBookingNo, Me.mNoUrut, Me.mprimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_inv1TableAdapter
        Dim ObjItem As New ClsTransaksiBookingInvItem
        Try
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
            If ex.Message = "Duplicate entry '" + Me.mBookingNo + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mBookingNo = Me.GetNoBaru(Me.mDate)
                Return Me.SimpanPerubahan(dgvItem, vUserCode, i)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
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
