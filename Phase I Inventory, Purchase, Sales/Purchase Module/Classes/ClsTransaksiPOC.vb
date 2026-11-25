Option Explicit On
Option Strict Off

Public Class ClsTransaksiPOC
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mprimaryPO As Double = -1
    Private mNoTransaksi As String = ""
    Private mUrutan As Integer = 0
    Private mTanggalTransaksi As Date = Today
    Private mPrimaryPurchaseStaffAuthorization As Double = -1
    Private mPurchaseStaffAuthorizationDate As Date = Now
    Private mPrimaryManagerStaffAuthorization As Double = -1
    Private mPurchaseManagerAuthorizationDate As Date = Now
    Private mPrimaryDirectorStaffAuthorization As Double = -1
    Private mPurchaseDirectorAuthorizationDate As Date = Now
    Private mAuthorizationlevel As Integer = 0
    Private mAuthorizationNeeded As Integer = 0
    Private mCancellationReason As String = ""
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now


    Public Property primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property primaryPO() As Double
        Get
            Return Me.mprimaryPO
        End Get
        Set(ByVal value As Double)
            mprimaryPO = value
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

    Public Property urutan() As Double
        Get
            Return Me.mUrutan
        End Get
        Set(ByVal value As Double)
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

    Public Property PrimaryPurchaseStaffAuthorization() As Double
        Get
            Return Me.mPrimaryPurchaseStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryPurchaseStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseStaffAuthorizationDate() As Date
        Get
            Return Me.mPurchaseManagerAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseStaffAuthorizationDate = value
        End Set
    End Property

    Public Property PrimaryManagerStaffAuthorization() As Double
        Get
            Return Me.mPrimaryManagerStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryManagerStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseManagerAuthorizationDate() As Date
        Get
            Return Me.mPurchaseManagerAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseManagerAuthorizationDate = value
        End Set
    End Property

    Public Property PrimaryDirectorStaffAuthorization() As Double
        Get
            Return Me.mPrimaryDirectorStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryDirectorStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseDirectorAuthorizationDate() As Date
        Get
            Return Me.mPurchaseDirectorAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseDirectorAuthorizationDate = value
        End Set
    End Property

    Public Property Authorizationlevel() As Integer
        Get
            Return Me.mAuthorizationlevel
        End Get
        Set(ByVal value As Integer)
            mAuthorizationlevel = value
        End Set
    End Property

    Public Property AuthorizationNeeded() As Integer
        Get
            Return Me.mAuthorizationNeeded
        End Get
        Set(ByVal value As Integer)
            mAuthorizationNeeded = value
        End Set
    End Property

    Public Property CancellationReason() As String
        Get
            Return mCancellationReason
        End Get
        Set(ByVal value As String)
            mCancellationReason = value
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

    '' dipakai di PO
    Public Function GetJmlByPO(ByVal vPrimaryPO As Double) As Integer
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).Value
            Else
                Return 0
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
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetjmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTglDariSampai(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDesc(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                Return DTA.ScalarQuery_GetJmlBySupplierCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                Return DTA.ScalarQuery_GetJmlSupplierNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                Return DTA.ScalarQuery_GetJmlByNoPOLIKe(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim objCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = objCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = objCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

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
            Return NoUrut + "/POC/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            objCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
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
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim TBL As New DseTransaksiPOC.tbltrans_poc2DataTable
        Dim row As DseTransaksiPOC.tbltrans_poc2Row
        Dim ObjItem As New ClsTransaksiPOCItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'fil basis data
            If DTA.FillBy_PrimaryPOC(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimary = .Primary
                    Me.mprimaryPO = .Primary_PO
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mUrutan = .Urutan
                    Me.TanggalTransaksi = .Tanggal_Transaksi
                    Me.mPrimaryPurchaseStaffAuthorization = .Primary_Purchase_Staff_Authorization
                    Me.mPurchaseStaffAuthorizationDate = .Purchase_Staff_Authorization_Date
                    Me.mPrimaryManagerStaffAuthorization = .Primary_Purchase_Manager_Authorization
                    Me.mPurchaseManagerAuthorizationDate = .Purchase_Manager_Authorization_Date
                    Me.mPrimaryDirectorStaffAuthorization = .Primary_Director_Authorization
                    Me.mPurchaseDirectorAuthorizationDate = .Director_Authorization_Date
                    Me.mAuthorizationlevel = .Authorization_Level
                    Me.mAuthorizationNeeded = .Authorization_Level_Needed
                    Me.mCancellationReason = .Cancellation_Reason
                    Me.mRemarks = .Remarks

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
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPOC.tbltrans_poc_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_view0TableAdapter
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
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SupplierCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryPOC As Double, _
    ByVal StrictEditting As Boolean, ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'packing list tidak boleh diedit jika:
            '0. Untuk strict editting harus lihat user edit apakah sama dengan yang menciptakan
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryPOC)
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
            '2. sudah diotorisasi
            If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryPOC).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This purchase order cancellation has been authorized, please cancel authorization before editting data.", _
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


    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryPOC As Double, _
    ByVal StrictEditting As Boolean, ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjPO As New Purchase_Module.ClsTransaksiPO
        Dim CreatedBy As String

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'packing list tidak boleh diedit jika:
            '1. batas tutup transaksi purchase sudah lewat
            'msgbox("belum")

            '0. Untuk strict editting harus lihat user edit apakah sama dengan yang menciptakan
            If StrictEditting = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryPOC)
                If CreatedBy = vUser Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. sudah diotorisasi
            If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryPOC).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This purchase order cancellation has been authorized, please cancel authorization before deleting data.", _
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
            ObjPO = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Try
            If vCekDataBenar = False Then
                Return True
            End If
            'msgbox("belum")

            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim ObjItem As New ClsTransaksiPOCItem
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim ObjPOItem As New ClsTransaksiPOItem

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, vCekDataBenar) = True Then
                'set level otorisasi yang dibutuhkan
                Me.mAuthorizationNeeded = 3
                'insert query
                If DTA.Insert(Me.primaryPO, Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mTanggalTransaksi, _
                 Me.PrimaryPurchaseStaffAuthorization, Me.PurchaseStaffAuthorizationDate, Me.PrimaryManagerStaffAuthorization, _
                 Me.PurchaseManagerAuthorizationDate, Me.PrimaryDirectorStaffAuthorization, Me.PurchaseManagerAuthorizationDate, _
                 Me.Authorizationlevel, Me.AuthorizationNeeded, Me.CancellationReason, Me.mRemarks, _
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
                    ObjItem.Simpan_Multi(vdgvItem, Me.mPrimary)
                    If vFromCutOff = False Then
                        'set purchase order balance
                        ObjPOItem.SetOrderBalance(Me.mprimaryPO)
                        'cek nomor
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
                Return Me.SimpanBaru(vdgvItem, vUserCode, vCekDataBenar, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjPOItem = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter)
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
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mNoTransaksi, Me.mUrutan, Me.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim ObjItem As New ClsTransaksiPOCItem
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim prim_po_lama As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(False) = True Then
                'set purchase order balance
                prim_po_lama = Me.GetPrimaryPO(Me.mPrimary)

                If DTA.Update(Me.mprimaryPO, Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mTanggalTransaksi, _
                 Me.PrimaryPurchaseStaffAuthorization, Me.PurchaseStaffAuthorizationDate, _
                 Me.PrimaryManagerStaffAuthorization, Me.PurchaseManagerAuthorizationDate, Me.PrimaryDirectorStaffAuthorization, _
                 Me.PurchaseDirectorAuthorizationDate, Me.Authorizationlevel, Me.AuthorizationNeeded, Me.CancellationReason, _
                 Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mUrutan = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan_Multi(dgvItem, Me.mPrimary)
                    ObjPOItem.SetOrderBalance(prim_po_lama)
                    If Me.mprimaryPO <> prim_po_lama Then
                        ObjPOItem.SetOrderBalance(Me.mprimaryPO)
                    End If
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
            ObjPOItem = Nothing
        End Try
    End Function

    Public Function GetPrimaryPO(ByVal vPrimaryPOC As Double) As Double
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If IsNothing(DTA.ScalarQuery_GetPrimaryPO(vPrimaryPOC).Value) = True Then
                Return -1
            Else
                Return DTA.ScalarQuery_GetPrimaryPO(vPrimaryPOC).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim Prim_PO As Double = -1
        Dim ObjPOItem As New ClsTransaksiPOItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vPrimaryTransaksi, False, "", TampilkanPesan) = True Then
                Prim_PO = Me.GetPrimaryPO(vPrimaryTransaksi)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    ObjPOItem.SetOrderBalance(Prim_PO)
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = 0
            mprimaryPO = 0
            mNoTransaksi = ""
            mTanggalTransaksi = Today
            mUrutan = 0
            mPrimaryPurchaseStaffAuthorization = -1
            mPurchaseStaffAuthorizationDate = Now
            mPrimaryManagerStaffAuthorization = -1
            mPurchaseManagerAuthorizationDate = Now
            mPrimaryDirectorStaffAuthorization = -1
            mPurchaseDirectorAuthorizationDate = Now
            ''msgbox("belum")
            mAuthorizationlevel = 0
            mAuthorizationNeeded = 3

            mCancellationReason = ""
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

    Public Function IsiChecklistBoxOtorisasi(ByVal vLevel As Integer, _
    ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiPOC.tbltrans_poc3DataTable, _
    Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'mengisikan semua quotation yang belum diotorisasi
            If vLevel = 0 Then
                If StrictSalesStaff = True Then
                    DTA.FillBy_UserNAuthoLevel(TBL, SalesStaff, vLevel)
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                End If
            ElseIf vLevel = 2 Then
                DTA.FillBy_MaxAuthorization(TBL, 2)
            Else
                DTA.FillBy_AuthorizationLevel(TBL, vLevel)
            End If

            'tampilkan pada checklist box
            clb.DataSource = TBL.DefaultView
            clb.DisplayMember = "No Transaksi"
            clb.Refresh()

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiChecklistBoxBatalOtorisasi(ByVal vLevel As Integer, _
    ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiPOC.tbltrans_poc3DataTable, _
    Optional ByVal StrictSalesStaff As Boolean = True, _
    Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'mengisikan semua quotation yang belum diotorisasi
            If StrictSalesStaff = True Then
                If vLevel = 1 Then
                    DTA.FillBy_CancelPurchaseStaffAuthorization(TBL, 1, SalesStaff)
                ElseIf vLevel = 2 Then
                    DTA.FillBy_CancelManagerAuthorization(TBL, 2, SalesStaff)
                ElseIf vLevel = 3 Then
                    DTA.FillBy_CancelDirectorAuthorization(TBL, 3, SalesStaff)
                End If
            Else
                DTA.FillBy_AuthorizationLevel(TBL, vLevel)
            End If

            'tampilkan pada checklist box
            clb.DataSource = TBL.DefaultView
            clb.DisplayMember = "No Transaksi"
            clb.Refresh()

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function CancelAuthorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
    ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim row As DseTransaksiPOC.tbltrans_poc3Row
        Dim rowview As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each rowview In clbAutho.CheckedItems
                row = rowview.Row
                If vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_StaffAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_ManagerAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                ElseIf vAuthorizationLevel = 3 Then
                    DTA.UpdateQuery_DirectorAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Authorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
    ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc2TableAdapter
        Dim row As DseTransaksiPOC.tbltrans_poc3Row
        Dim rowview As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each rowview In clbAutho.CheckedItems
                row = rowview.Row
                If vAuthorizationLevel = 0 Then
                    DTA.UpdateQuery_StaffAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                ElseIf vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_ManagerAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
