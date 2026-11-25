Option Explicit On
Option Strict Off

Public Class ClsTransaksiPackingList
    Dim ObjCD As New Core_Module.ClsCekData

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

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim DTA1 As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter

        Try
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
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaSupplier Then
                Return DTA1.ScalarQuery_GetJmlBySupplierNameLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksi(vPrimary)) = False Then
                Return DTA.ScalarQuery_getNoTransaksi(vPrimary)
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

    Public Function GetSupplierPackingListNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierPackingListNo(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetSupplierPackingListNo(vPrimary)
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

    Public Function GetSupplierCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getSupplierCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_getSupplierCode(vPrimary)
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
    ''' mengembalikan no transaksi stock in purchase yang merujuk pada packing list
    ''' yang diberikan primarynya
    ''' </summary>
    ''' <param name="vPrimary"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoStockInPurchase(ByVal vPrimary As Double) As String
        Dim Obj As New clsTransaksiStockInOut

        Try
            Return Obj.GetStockInNumberByPrimaryPackingList(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Obj = Nothing
        End Try
    End Function

    Public Sub GetInfoSupplierNameNCodeNTransNo(ByVal vPrimary As Double, ByRef vSupName As String, _
    ByRef vSupCode As String, ByRef vNoTrans As String)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_view0DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_view0Row
        Try
            DTA.FillBy_primaryTransaksi(TBL, vPrimary)
            vSupName = ""
            vSupCode = ""
            vNoTrans = ""
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vSupName = row.Supplier_Name
                vSupCode = row.Supplier_Code
                vNoTrans = row.Packing_List_Number
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Try
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
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
            Return NoUrut + "/DO/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
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
                ObjItem.IsiDGV2_Multi(vDgvItem, Me.mPrimary)
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

    'Public Sub IsiTBLPackingList(ByVal vPrimarySupplier As Double, ByVal vPrimaryPackingList As Double, _
    'ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view1DataTable)
    '    Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view1TableAdapter

    '    Try
    '        TBL.Rows.Clear()
    '        DTA.FillBy_PrimaryPackingList_PrimarySupplier(TBL, vPrimarySupplier, vPrimaryPackingList)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLPackingList(ByVal vPrimarySupplier As Double, ByVal vPrimaryPackingList As Double, _
    ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view1MultiDataTable)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view1MultiTableAdapter

        Try
            TBL.Rows.Clear()
            DTA.FillBy_PrimaryPackingList_PrimarySupplier(TBL, vPrimarySupplier, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ' Public Sub IsiTBLPackingList_PO(ByVal vPrimaryPackingList As Double, _
    'ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view1DataTable, ByVal vPrimaryPO As Double)
    '     Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view1TableAdapter

    '     Try
    '         TBL.Rows.Clear()
    '         DTA.FillBy_PrimaryPO_PrimaryPackingList(TBL, _
    '         vPrimaryPO, vPrimaryPackingList)
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         DTA.Dispose()
    '     End Try
    ' End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLPackingList_PO(ByVal vPrimaryPackingList As Double, _
   ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view1MultiDataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view1MultiTableAdapter

        Try
            TBL.Rows.Clear()
            DTA.FillBy_PrimaryPO_PrimaryPackingList(TBL, _
            vPrimaryPO, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPackingList.tbltrans_packinglist_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryPackingList As Double, _
    ByVal StrictEditting As Boolean, ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut

        Try
            'packing list tidak boleh diedit jika:
            '0. Untuk strict editting harus lihat user edit apakah sama dengan yang menciptakan
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryPackingList)
                If CreatedBy = vUser Then
                    Return True
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
            '     Dim NoStockIn As String = ObjStockInOut.GetStockInNumberByPrimaryPackingList(vPrimaryPackingList)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjStockInOut = Nothing
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

    Public Sub UpdateEstimateDelivery(ByVal vPrimaryPackingList As Double, ByVal vNewETA As Date)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetETA(vNewETA, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vPrimaryPackingList As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml(vPrimaryPackingList, vPrimaryCodeNo, vPrimaryPO)
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

    Public Function Get_EstimateArrival(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEstimateArrival(vPrimary)
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

    Public Function ApakahSudahStockIn(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter
        Dim JmlStockIn As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            JmlStockIn = DTA.ScalarQuery_GetJmlStockIn(clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, vPrimary)
            If JmlStockIn.HasValue = True Then
                If JmlStockIn.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
