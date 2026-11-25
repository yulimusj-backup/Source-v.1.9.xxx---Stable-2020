Option Explicit On
Option Strict Off

Public Class ClsTransaksiDOLoss
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mUrutan As Integer = 0
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""

    Private mPrimaryStockoutAdjust As Double = -1
    Private mPrimaryCustomer As Double = -1
    Private mAddressIdentifier As String = ""
    Dim mReplacementDOLoss As EnumTipeDOLoss = EnumTipeDOLoss.Replacement

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumTipeDOLoss As Integer
        Replacement
        Sample_Promotion
    End Enum

    Public Property ReplacementDOLoss() As EnumTipeDOLoss
        Get
            Return mReplacementDOLoss
        End Get
        Set(ByVal value As EnumTipeDOLoss)
            mReplacementDOLoss = value
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

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property PrimaryStockOutAdjust() As Double
        Get
            Return mPrimaryStockoutAdjust
        End Get
        Set(ByVal value As Double)
            mPrimaryStockoutAdjust = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return mPrimaryCustomer
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomer = value
        End Set
    End Property

    Public Property AddressIdentifier() As String
        Get
            Return mAddressIdentifier
        End Get
        Set(ByVal value As String)
            mAddressIdentifier = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
       ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
       Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim ObjStockAdjust As New Inventory_Module.clsTransaksiStockInOut
        Dim NoTransaksi As String = ""

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

            '2. Jika sudah ada stock out adjust yang merujuk pada transaksi DO Loss yang bersangkutan maka tidak bisa
            '   diedit
            NoTransaksi = ObjStockAdjust.GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(vPrimaryTransaksi, _
            Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
            If NoTransaksi <> "" Then
                MsgBox("This data referred by stock adjust: " + NoTransaksi + ", you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
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
            ObjStockAdjust = Nothing
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
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
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
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            TanggalEnd = Core_Module.ModCekData.GetLasDateOfTheMonth(vTanggalTransaksi)
            TanggalStart = Core_Module.ModCekData.GetFirstDateofTheMonth(vTanggalTransaksi)

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
            Return NoUrut + "/" + "DOLoss" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
      ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim TBL As New DseTransaksiDOLoss.tbltrans_do_lossDataTable
        Dim row As DseTransaksiDOLoss.tbltrans_do_lossRow
        Dim ObjItem As New ClsTransaksiDOLossItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                row = TBL.Rows(0)
                'salin data header

                With row
                    mNoTransaksi = .No_Transaksi
                    mUrutan = .Urutan
                    mTanggalTransaksi = .Tanggal_Transaksi
                    mRemarks = .Remarks
                    Me.mPrimary = .Primary

                    mPrimaryStockoutAdjust = .Primary_Stock_Out_Adjust
                    mPrimaryCustomer = .Primary_Customer
                    mAddressIdentifier = .Address_Identifier
                    Me.mReplacementDOLoss = .Replacement_DO_Loss

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary)
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mUrutan = 0
            mPrimaryStockoutAdjust = -1
            mPrimaryCustomer = -1
            mAddressIdentifier = ""
            mTanggalTransaksi = Now
            mRemarks = ""

            Dim mReplacementDOLoss As EnumTipeDOLoss = EnumTipeDOLoss.Replacement
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim ObjItem As New ClsTransaksiDOLossItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True) = True Then
                'insert query
                If DTA.Insert(Me.NoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.PrimaryStockOutAdjust, Me.PrimaryCustomer, Me.AddressIdentifier, _
                Me.TanggalTransaksi, Me.ReplacementDOLoss, Me.Remarks, vUserCode, Now, _
                vUserCode, Now, Now) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimaryBynoTransaksi(Me.mNoTransaksi).Value
                    Me.mUrutan = CInt(Me.mNoTransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary)
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

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim ObjItem As New ClsTransaksiDOLossItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False) = True Then
                'hitung level otorisasi yang dibutuhkan
                If DTA.Update(Me.NoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.PrimaryStockOutAdjust, Me.PrimaryCustomer, Me.AddressIdentifier, _
                Me.TanggalTransaksi, Me.ReplacementDOLoss, Me.Remarks, vUserCode, Now, Now, Me.Primary, Me.TimeStamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mUrutan = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary)
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

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Dim ObjStockAdjust As New Inventory_Module.clsTransaksiStockInOut
        Dim NoTransaksi As String = ""

        Try
            'msgbox("belum")
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            '2. Jika sudah ada stock out adjust yang merujuk pada transaksi DO Loss yang bersangkutan maka tidak bisa
            '   diedit
            NoTransaksi = ObjStockAdjust.GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(vPrimaryTransaksi, _
            Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
            If NoTransaksi <> "" Then
                MsgBox("This data referred by stock adjust: " + NoTransaksi + ", you delete edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjStockAdjust = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
           UserCode) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryBynoTransaksi(vNoTrans).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryBynoTransaksi(vNoTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_lossTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoDOLoss(vPrimary)) = False Then
                Return DTA.ScalarQuery_getNoDOLoss(vPrimary)
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

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetjmlByNotransaksiLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTgldarisampai(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%").value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.ScalarQuery_GetJmlBySONoLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustomerCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustomerName(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiDOLoss.tbltrans_doloss_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_view0TableAdapter
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
                QueryResult = DTA.FillBy_Tanggaldarisampai(TBL, vTanggalDari, vTanggalSampai)

            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescriptionLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustomerCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustomerNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                QueryResult = DTA.FillBy_SONumberLIKE(TBL, Criteria + "%")
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
End Class
