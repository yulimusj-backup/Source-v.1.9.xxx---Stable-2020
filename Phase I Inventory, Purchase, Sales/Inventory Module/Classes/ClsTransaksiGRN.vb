Option Explicit On
Option Strict Off

Public Class ClsTransaksiGRN
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
                    Return True
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            'tidak ada batasan lain untuk edit transaksi karena, pada dasarnya batasan edit
            'sebenarnya adalah pada item, jadi kalo rajin mungkin bisa ditambahin:
            'JIKA SEMUA ITEM TIDAK BISA DIEDIT MAKA TRANSAKSI TIDAK BISA DIEDIT
            'msgbox("belum")

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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Overridable Function SimpanBaruLewatTransaksiStockInPurchase(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim ObjItem As New ClsTransaksiGRNItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
                ObjItem.SimpanLewatTransaksiStockInOut_Multi(vdgvItem, Me.mprimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 3" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mtanggalTransaksi)
                Return Me.SimpanBaruLewatTransaksiStockInPurchase(vdgvItem, vUserCode)
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring

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
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function HapusDariStockInPurchase(ByVal vPrimaryStockInPurchase As Double) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'hapus data header, data item otomatis dihapuskan oleh server
            If DTA.DeleteQuery_ByPrimaryStockInPuchase(vPrimaryStockInPurchase) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
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

    Public Function GetPrimary_ByPrimaryStockInPurchase(ByVal vPrimaryStockInPurchase As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary_ByPrimaryStockIn(vPrimaryStockInPurchase).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary_ByPrimaryStockIn(vPrimaryStockInPurchase).Value
            Else
                Return -1
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

    Public Function GetJmlAOPbyPrimaryStockInOut(ByVal vPrimaryStock As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.DataTable11TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_getJmlAOPByPrimaryStockInOut(vPrimaryStock).HasValue Then
                Return DTA.ScalarQuery_getJmlAOPByPrimaryStockInOut(vPrimaryStock).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoAOP(ByVal vPrimaryStock As Double) As String
        Dim DTA As New DseTransaksiGRNTableAdapters.DataTable11TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoAOPByPrimaryStock(vPrimaryStock)) = False Then
                Return DTA.ScalarQuery_GetNoAOPByPrimaryStock(vPrimaryStock)
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

    ''' <summary>
    ''' mengembalikan no transaksi purchase return yang merujuk ke stock in purchase yang diberikan primary-nya
    ''' </summary>
    ''' <param name="vPrimaryStockInPurchase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoPurchaseReturn(ByVal vPrimaryStockInPurchase As Double) As String
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBL As New DseTransaksiGRN.tbltrans_purchasereturnDataTable
        Dim row As DseTransaksiGRN.tbltrans_purchasereturnRow
        Dim dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryStockInPurchase)
            For Each row In TBL.Rows
                If dummy = "" Then
                    dummy = row.PR_Number
                Else
                    dummy = dummy + "," + row.PR_Number
                End If
            Next

            Return dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
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
                ObjItem.IsiDGV(vDgvItem, Me.mprimary)
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
        End Try
    End Function

    Public Function GetPrimaryStockInPurchase(ByVal vPrimaryGRN As Double) As Double
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryStockInpurchase(vPrimaryGRN)
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
End Class
