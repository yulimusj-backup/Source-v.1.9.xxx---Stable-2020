Option Explicit On
Option Strict Off

Public Class ClsTransaksiGRN
    Dim ObjCD As New Core_Module.ClsCekData

    Private mprimary As Double = -1
    Private mnotransaksi As String = ""
    Private murutan As Integer = 0
    Private mprimarystockinpurchase As Double = -1
    Private mprimaryGIT As Double = -1
    Private mprimarySupplier As Double = -1
    Private mtanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    Private mSupplierDONumber As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property SupplierDONumber() As String
        Get
            Return mSupplierDONumber
        End Get
        Set(ByVal value As String)
            mSupplierDONumber = value
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

    Public Property NoTransaksi() As String
        Get
            Return Me.mnotransaksi
        End Get
        Set(ByVal value As String)
            Me.mnotransaksi = value
        End Set
    End Property

    Public Property primaryStockInPurchase() As Double
        Get
            Return Me.mprimarystockinpurchase
        End Get
        Set(ByVal value As Double)
            Me.mprimarystockinpurchase = value
        End Set
    End Property

    Public Property urutan() As Integer
        Get
            Return Me.murutan
        End Get
        Set(ByVal value As Integer)
            Me.murutan = value
        End Set
    End Property

    Public Property primaryGIT() As Double
        Get
            Return Me.mprimaryGIT
        End Get
        Set(ByVal value As Double)
            Me.mprimaryGIT = value
        End Set
    End Property

    Public Property primarySupplier() As Double
        Get
            Return Me.mprimarySupplier
        End Get
        Set(ByVal value As Double)
            Me.mprimarySupplier = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mtanggalTransaksi
        End Get
        Set(ByVal value As Date)
            Me.mtanggalTransaksi = value
        End Set
    End Property

    Public Property Remark() As String
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
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByGRN(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByGRN(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mnotransaksi) = True Then
                    Me.mnotransaksi = Me.GetNoBaru(Me.mtanggalTransaksi)
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika sudah melewati batas penguncian transaksi

            '4. jika tidak ada item yang akan disimpan

            '5. jika item yang disimpan mengakibatkan order balance minus

            '6. jika item yang disimpan nilainya lebih kecil dari yang sudah dibooking via IPO--> hati - hati ini

            '7. jika item yang disimpan nilainya lebih kecil dari yang sudah booking inv
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMAXUrutan(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMAXUrutan(TanggalStart, TanggalEnd).Value + 1)
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
            Return NoUrut + "/" + "GRN" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function



    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
          ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
          Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit PO:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
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

            'emang button edit-nya juga ga ada


            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryGRN As Double, _
   Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim ObjItemAOPItem As New ClsTransaksiAOPItem
        Try
            'Suatu transaksi do tidak bisa diedit jika:
            '1. sudah ada transaksi lanjutannya.
            '   a. stock out sales
            '            If DTA.ScalarQuery_GetPrimaryStockInpurchase(vPrimaryGRN).HasValue = True Then
            ' If DTA.ScalarQuery_GetPrimaryStockInpurchase(vPrimaryGRN).Value > 0 Then
            ' If TampilkanPesan = True Then
            ' MsgBox("There stock in refering to this transaction, cannot delete this data.", _
            ' MsgBoxStyle.Information, "Information")
            ' End If
            ' Return False
            ' End If
            ' End If
            '   b. Transaksi AOP
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If ObjItemAOPItem.GetJmlByGRN(vPrimaryGRN) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There sales invoice:" + ObjItemAOPItem.GetJmlByGRN(vPrimaryGRN) + _
                    " refering to this transaction, cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '   c. sales return 
            'If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).HasValue = True Then
            'If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).Value <= 0 Then
            'If TampilkanPesan = True Then
            'MsgBox("There are sales return refering to this transaction, cannot delete this data.", _
            ' MsgBoxStyle.Information, "Information")
            'End If
            ' Return False
            ' End If
            'End If

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            '4. msgbox("belum")

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItemAOPItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim ObjItem As New ClsTransaksiGRNItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True) = True Then
                'hitung level otorisasi yang dibutuhkan
                'insert query
                Me.murutan = CInt(Me.mnotransaksi.Substring(0, 6))
                If DTA.Insert(Me.NoTransaksi, Me.urutan, Me.primaryStockInPurchase, Me.primaryGIT, _
                Me.primarySupplier, Me.TanggalTransaksi, Me.Remark, vUserCode, Now, _
                vUserCode, Now, Now, Me.SupplierDONumber) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mprimary = DTA.ScalarQuery_GetPrimaryByNoTransaksi(Me.mnotransaksi).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mprimary)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 3" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mtanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim ObjItem As New ClsTransaksiGRNItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(False) = True Then
                'hitung level otorisasi yang dibutuhkan
                Me.murutan = Me.mnotransaksi.Substring(0, 6)
                If DTA.Update(Me.NoTransaksi, Me.urutan, Me.primaryStockInPurchase, Me.primaryGIT, Me.primarySupplier, _
                Me.TanggalTransaksi, Me.Remark, vUserCode, Now, Now, Me.SupplierDONumber, _
                Me.primary, Me.TimeStamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mprimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mprimary)
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

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim DTA2 As New DseTransaksiGRNTableAdapters.DataTable1TableAdapter

        Dim TBL As New DseTransaksiGRN.DataTable1DataTable
        Dim row As DseTransaksiGRN.DataTable1Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                DTA2.FillBy_PrimaryGRN(TBL, vPrimaryTransaksi)
                If TBL.Rows.Count > 0 Then
                    If MsgBox("If you remove this GRN all Stock In that related with this GRN will be remove.", MsgBoxStyle.OkCancel, "Confirm") = MsgBoxResult.Ok Then
                        Dim objStock As New Inventory_Module.clsTransaksiStockInOut
                        For Each row In TBL.Rows
                            objStock.HapusDariGRN(row.Primary_Stock_In_Purchase)
                        Next
                        objStock = Nothing
                        If DTA.Delete(vPrimaryTransaksi) > 0 Then
                            Return True
                        End If
                    Else
                        Return False
                    End If
                Else
                    If DTA.Delete(vPrimaryTransaksi) > 0 Then
                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiGRN.tbltrans_grn_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiGRNLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TgldariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SuppCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SuppNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierDONo Then
                QueryResult = DTA.FillBy_SuppDoNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                QueryResult = DTA.FillBy_NoPOLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierDONo Then
                QueryResult = DTA.FillBy_SuppDoNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                QueryResult = DTA.FillBy_StockInNo_LIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
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
            mprimary = -1
            mnotransaksi = ""
            murutan = 0
            mtanggalTransaksi = Now
            mprimarystockinpurchase = -1
            mprimaryGIT = -1
            mprimarySupplier = -1
            mRemarks = ""

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTimeStamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimeStamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryGRNByPrimaryStockIn(ByVal vPrimaryStockIn As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryByPrimaryStockIn(vPrimaryStockIn).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByPrimaryStockIn(vPrimaryStockIn).Value
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

    Public Function GetPrimaryStockIn(ByVal vPrimaryGRN As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryStockInpurchase(vPrimaryGRN).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryStockInpurchase(vPrimaryGRN).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimaryGRN As Double) As Date
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryGRN).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryGRN).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetPrimaryByNoTransaksi(vNoTrans).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimarySupplier(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimarySupplierByPrimaryGRN(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimarySupplierByPrimaryGRN(vPrimary).Value
            Else
                Return -1
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
        Dim DTA As New DseTransaksiGRNTableAdapters.TblTransGRN_DGVTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlTanggalTrans(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetjmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlProductDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                Return DTA.ScalarQuery_GetJmlBySuppNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierDONo Then
                Return DTA.ScalarQuery_GetJmlBySupplierDONOLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByStockInNo Then
                Return DTA.ScalarQuery_GetJmlByNoStockInLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                Return DTA.ScalarQuery_GetJmlByNoPOLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
   ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim TBL As New DseTransaksiGRN.tbltrans_grnDataTable
        Dim row As DseTransaksiGRN.tbltrans_grnRow
        Dim ObjItem As New ClsTransaksiGRNItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mnotransaksi = .No_Transaksi
                    Me.murutan = .Urutan
                    Me.mprimary = .Primary
                    Me.mprimarySupplier = .Primary_Supplier
                    Me.mprimaryGIT = .Primary_GIT
                    Me.mRemarks = .Remarks
                    Me.mtanggalTransaksi = .Tanggal_Transaksi
                    mprimarystockinpurchase = .Primary_Stock_In_Purchase

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV_Default(vDgvItem, Me.mprimary)
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

    ''' <summary>
    ''' di pake di purchase return 
    ''' </summary>
    ''' <param name="vTBL"></param>
    ''' <remarks></remarks>
    Public Sub Isikantblantaraviewgrnnpurchasereturn_view_Default(ByVal vTBL As DseTransaksiGRN.DataTable2DataTable, _
    ByVal vPrimaryStockGRN As Double, ByVal vPrimaryPurchaseReturn As Double)
        Dim DTA_GRN As New DseTransaksiGRNTableAdapters.DataTable2TableAdapter
        Dim row_GRN As DseTransaksiGRN.DataTable2Row
        Dim DTA_PR As New DseTransaksiGRNTableAdapters.DataTable4TableAdapter
        Dim TBL_PR As New DseTransaksiGRN.DataTable4DataTable
        Dim row_PR As DseTransaksiGRN.DataTable4Row
        Dim filter As String = ""

        Try
            DTA_GRN.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PR.Connection.ConnectionString = DTA_GRN.Connection.ConnectionString

            'isikan data GRN
            DTA_GRN.FillBy(vTBL, vPrimaryStockGRN)

            'kurangkan dengan semua purchase return yang merujuk pada GRN ybs, kecuali purchase return yang diberikan primarynya
            DTA_PR.Fill(TBL_PR, vPrimaryStockGRN)
            For Each row_PR In TBL_PR.Rows
                If row_PR.Primary_Purchase_Return <> vPrimaryPurchaseReturn Then
                    filter = "[Primary Product] = " + row_PR.Primary_Product.ToString + " AND [Primary PO]= " + row_PR.Primary_PO.ToString + _
                    " AND [Batch No] ='" + row_PR.Batch_No + "' AND [Secondary Qty] = " + row_PR.Secondary_Qty.ToString
                    vTBL.DefaultView.RowFilter = filter
                    If vTBL.DefaultView.Count > 0 Then
                        row_GRN = vTBL.DefaultView(0).Row
                        row_GRN.Total_Qty -= row_PR.Total_Qty
                        If row_GRN.Total_Qty <= 0 Then
                            vTBL.Rows.Remove(row_GRN)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_GRN.Dispose()
            DTA_PR.Dispose()
            TBL_PR.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub Isikantblantaraviewgrnnpurchasereturn_view(ByVal vTBL As DseTransaksiGRN.DataTable2_MultiDataTable, _
    ByVal vPrimaryStockGRN As Double, ByVal vPrimaryPurchaseReturn As Double, ByRef vStrRetur As String)
        Dim DTA_GRN As New DseTransaksiGRNTableAdapters.DataTable2_MultiTableAdapter
        Dim row_GRN As DseTransaksiGRN.DataTable2_MultiRow
        Dim DTA_PR As New DseTransaksiGRNTableAdapters.DataTable4_MultiTableAdapter
        Dim TBL_PR As New DseTransaksiGRN.DataTable4_MultiDataTable
        Dim row_PR As DseTransaksiGRN.DataTable4_MultiRow
        Dim filter As String = ""
        Dim objR As New ClsTransaksiPurchaseReturn

        Try
            DTA_GRN.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PR.Connection.ConnectionString = DTA_GRN.Connection.ConnectionString
            'isikan data GRN
            DTA_GRN.Fill(vTBL, vPrimaryStockGRN)

            Dim blnHabis As Boolean = False
            'kurangkan dengan semua purchase return yang merujuk pada GRN ybs, kecuali purchase return yang diberikan primarynya
            Dim ListNoRetur As New List(Of Double)
            Dim strNORetur As String = ""
            DTA_PR.Fill(TBL_PR, vPrimaryStockGRN)
            If TBL_PR.Rows.Count > 0 Then
                For Each row_PR In TBL_PR.Rows
                    If row_PR.Primary_Purchase_Return <> vPrimaryPurchaseReturn Then
                        filter = "[Primary Product] = " + row_PR.Primary_Product.ToString + " AND [Primary PO]= " + row_PR.Primary_PO.ToString + _
                                " AND [Batch No] ='" + row_PR.Batch_No + "' AND [Secondary Qty] = " + row_PR.Secondary_Qty.ToString
                        vTBL.DefaultView.RowFilter = filter
                        If vTBL.DefaultView.Count > 0 Then
                            row_GRN = vTBL.DefaultView(0).Row
                            row_GRN.Total_Qty -= row_PR.Total_Qty

                            ''@rosadinnn
                            ''@MUaT
                            'If row_GRN.Total_Qty <= 0 Then
                            '    vTBL.Rows.Remove(row_GRN)
                            'End If
                            row_GRN.Total_Qty_Multi -= row_PR.Total_Qty_Multi

                            If ListNoRetur.Contains(row_PR.Primary_Purchase_Return) = False Then
                                ListNoRetur.Add(row_PR.Primary_Purchase_Return)
                                If strNORetur = "" Then
                                    strNORetur &= objR.GetNoTransaksi(row_PR.Primary_Purchase_Return)
                                Else
                                    strNORetur &= ", " & objR.GetNoTransaksi(row_PR.Primary_Purchase_Return)
                                End If
                            End If


                            If row_GRN.Total_Qty_Multi <= 0 Then
                                vTBL.Rows.Remove(row_GRN)
                            End If
                        End If
                    End If
                Next
                Dim strNoGRN As String = Me.GetNoGRN(vPrimaryStockGRN)

                If vTBL.Rows.Count = 0 Then
                    vStrRetur = strNORetur
                End If
            End If
           

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_GRN.Dispose()
            DTA_PR.Dispose()
            TBL_PR.Dispose()
        End Try
    End Sub

    Public Function IsiTBLGRNBelumAdaAOP(ByVal vPrimarySupplier As Double, _
    ByVal vTBL As DseTransaksiGRN.DataTable8DataTable) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.DataTable8TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL, vPrimarySupplier)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah grn yang belum di AOP
    ''' </summary>
    ''' <param name="vTanggalSampai"></param>
    ''' <param name="vTanggalDari"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlBelumAOP(ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlGRNBelumAOP(vTanggalSampai, vTanggalDari).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoStockIn(ByVal vNoGRN As String) As String
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoStockIn(vNoGRN)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoStockIn(vNoGRN)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoGRN_BelumAOP(ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date) As String
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_view1TableAdapter
        Dim TBL As New DseTransaksiGRN.tbltrans_grn_view1DataTable
        Dim row As DseTransaksiGRN.tbltrans_grn_view1Row
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_GRNBelumAOP_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari) > 0 Then
                For Each row In TBL.Rows
                    If Retval = "" Then
                        Retval = Me.GetNoStockIn(row.GRN_Number)
                    Else
                        Retval = Retval + ", " + Me.GetNoStockIn(row.GRN_Number)
                    End If
                Next

                Return Retval
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

    Public Function GetNoGRN(ByVal Vprimary As Double) As String
        Dim dta As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoGRN(Vprimary)) = False Then
                Return dta.ScalarQuery_GetNoGRN(Vprimary)
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

    Public Sub GetInfo(ByVal VPrimaryStockIn As Double, ByRef vgrnNo As String, _
       ByRef vPrimaryGRN As String, ByRef vAOPNo As String, ByRef vPRimarySupplier As String, _
       ByRef vSupplierCode As String, ByRef vSupplierName As String, ByRef vSuplierDoName As String)
        Dim dta As New DseTransaksiGRNTableAdapters.DataTable11TableAdapter
        Dim TBL As New DseTransaksiGRN.DataTable11DataTable
        Dim row As DseTransaksiGRN.DataTable11Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            vgrnNo = ""
            vPrimaryGRN = "-1"
            vAOPNo = ""
            vPRimarySupplier = "-1"
            vSupplierCode = ""
            vSupplierName = ""
            vSuplierDoName = ""
            dta.FillBy_PrimaryStockIn(TBL, VPrimaryStockIn)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vgrnNo = row.No_GRN
                vPrimaryGRN = row.PrimaryGRN
                vAOPNo = row.No_AOP
                vPRimarySupplier = row.Primary_Supplier
                vSupplierCode = row.Supplier_Code
                vSupplierName = row.Supplier_Name
                vSuplierDoName = row._Supplier_s_DO_Number
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Function GetNoGRNByPrimaryStockInOut(ByVal vPrimaryStockInOut As Double) As String
        Dim dta As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(dta.ScalarQuery_getNoTransaksiByPrimaryStokInOut(vPrimaryStockInOut)) = False Then
                Return dta.ScalarQuery_getNoTransaksiByPrimaryStokInOut(vPrimaryStockInOut)
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

    Public Function GetPriamryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
