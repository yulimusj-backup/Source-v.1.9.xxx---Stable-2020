Option Strict Off
Option Explicit On

Public Class ClsTransaksiBookingGIT
    Dim ObjCD As New Core_Module.ClsCekData

    Private mprimary As Double = -1
    Private mBookingNo As String = ""
    Private mNoUrut As Integer = -1
    Private mDate As Date = Today
    Private mPrimaryCust As Double = -1
    Private mPrimaryGIT As Double = -1
    Private mRemarks As String = ""
    Private mStatusAktif As Boolean = True
    Private msoftwaregenerate As Integer = 0
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

    Public Property StatusAktif() As Boolean
        Get
            Return mStatusAktif
        End Get
        Set(ByVal value As Boolean)
            mStatusAktif = value
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

    Public Property primaryGIT() As Double
        Get
            Return Me.mPrimaryGIT
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryGIT = value
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

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
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
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd).Value + 1)
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
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Try
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'tidak boleh edit bila:
            '1. software generated
            If DTA.ScalarQuery_GetSoftwareGenerated(vPrimaryTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetSoftwareGenerated(vPrimaryTransaksi).Value <> 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction is software generated, you cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                If TampilkanPesan = True Then
                    MsgBox("This transaction is software generated, you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. statusnya sudah non aktif
            If DTA.ScalarQuery_GetStatusAktif(vPrimaryTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetStatusAktif(vPrimaryTransaksi).Value = 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction is non-active, the computer had generate inventory booking from this transaction, you cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                If TampilkanPesan = True Then
                    MsgBox("This transaction is non-active, the computer had generate inventory booking from this transaction, you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
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
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Try
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. statusnya sudah non aktif
            If DTA.ScalarQuery_GetStatusAktif(vPrimaryTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetStatusAktif(vPrimaryTransaksi).Value = 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction is non-active, the computer had generate inventory booking from this transaction, you cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                If TampilkanPesan = True Then
                    MsgBox("This transaction is non-active, the computer had generate inventory booking from this transaction, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

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
        Dim dta As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
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
        Dim dta As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
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

    ''' <summary>
    ''' Untuk Semua No Transaksi Booking GIT yang di gunakan oleh GIT tertentu
    ''' </summary>
    ''' <param name="vPrimaryGIT"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoBookingGITAktifByPrimaryGIT(ByVal vPrimaryGIT As Double) As String
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Dim TBL As New DseTransaksiBookingGIT.tbltrans_booking_git1DataTable
        Dim row As DseTransaksiBookingGIT.tbltrans_booking_git1Row
        Dim KumpulanNoBooking As String = ""
        Try
            DTA.FillBy_PrimaryGIT(TBL, vPrimaryGIT)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If row.Status_Aktif = True Then
                        If KumpulanNoBooking = "" Then
                            KumpulanNoBooking = row.No_Transaksi
                        Else
                            KumpulanNoBooking = KumpulanNoBooking + ", " + row.No_Transaksi
                        End If
                    End If
                Next
            End If
            Return KumpulanNoBooking
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
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
            mRemarks = ""
            msoftwaregenerate = 0
            mPrimaryGIT = -1
            mStatusAktif = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
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
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiBookingGIT.DTable2DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DTable2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglsampaiDari(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustomerCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustomerNameLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_DescriptionLIke(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                QueryResult = DTA.FillBy_SOnumber(TBL, Criteria + "%")
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
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DTable21TableAdapter
        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.GetJmlBy_NoTransaksiLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.GetJmlBy_TglsampaiDari(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.GetJmlBy_ProductCodeLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.GetJmlBy_DescriptionLIke(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpec(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.GetJmlBy_SOnumber(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.GetJmlBy_CustomerCodeLike(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.GetJmllBy_CustomerNameLike(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
        ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(CStr(vPrimaryTransaksi)) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTransaksiBookingGITTableAdapters.tbltrans_booking_git1TableAdapter)
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
End Class
