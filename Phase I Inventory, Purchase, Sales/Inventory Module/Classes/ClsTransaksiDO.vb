Option Explicit On
Option Strict Off

Public Class ClsTransaksiDO
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""

    Private mPrimaryStockOutSales As Double = -1
    Private mPrimaryCust As Double = -1
    Private mCustAddIdentifier As String = ""
    Private mCetakKwitansi As Integer = 0
    Private mCetakFaktur As Integer = 0
    Private mAuthorizationLevel As Integer = 0
    Private mAuthorizationLevelNeeded As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Property PrimaryStockOutSales() As Integer
        Get
            Return Me.mPrimaryStockOutSales
        End Get
        Set(ByVal value As Integer)
            mPrimaryStockOutSales = value
        End Set
    End Property

    Public Property AuthorizationLevel() As Integer
        Get
            Return Me.mAuthorizationLevel
        End Get
        Set(ByVal value As Integer)
            mAuthorizationLevel = value
        End Set
    End Property

    Public Property AuthorizationLevelNeeded() As Integer
        Get
            Return Me.mAuthorizationLevelNeeded
        End Get
        Set(ByVal value As Integer)
            mAuthorizationLevelNeeded = value
        End Set
    End Property

    Public Property CetakFaktur() As Boolean
        Get
            If Me.mCetakFaktur = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mCetakFaktur = 1
            Else
                Me.mCetakFaktur = 0
            End If
        End Set
    End Property

    Public Property CetakKwitansi() As Boolean
        Get
            If Me.mCetakKwitansi = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mCetakKwitansi = 1
            Else
                Me.mCetakKwitansi = 0
            End If
        End Set
    End Property

    Public Property CustAddIdentifier() As String
        Get
            Return Me.mCustAddIdentifier
        End Get
        Set(ByVal value As String)
            mCustAddIdentifier = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return Me.mPrimaryCust
        End Get
        Set(ByVal value As Double)
            mPrimaryCust = value
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

    Public Function SetPrimaryStockOut(ByVal vPrimaryDO As Double, ByVal vPrimaryStockOut As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'set nilai primary stock out & nilai authorization level
            If vPrimaryStockOut = -1 Then
                DTA.UpdateQuery_UpdateStatusDO(vPrimaryStockOut, 0, vPrimaryDO)
            Else
                DTA.UpdateQuery_UpdateStatusDO(vPrimaryStockOut, 1, vPrimaryDO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaInvoice(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlItemInvoice(vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaSalesReturn(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
       ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_do3DataTable
        Dim row As DseTransaksiDO.tbltrans_do3Row
        Dim ObjItem As New ClsTransaksiDOItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi

                    mPrimaryStockOutSales = .Primary_Stock_Out_Sales
                    mPrimaryCust = .Primary_Customer
                    mCustAddIdentifier = .Address_Identifier
                    mCetakKwitansi = .Cetak_Kwitansi
                    mCetakFaktur = .Cetak_Invoice_N_Faktur_Pajak
                    mAuthorizationLevel = .Authorization_Level
                    mAuthorizationLevelNeeded = .Authorization_Level_Needed

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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimaryStockOutSales = -1
            mCustAddIdentifier = ""
            mCetakKwitansi = 0
            mCetakFaktur = 0
            mAuthorizationLevel = 0
            mAuthorizationLevelNeeded = 0
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mPrimaryCust = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPrimaryStockOut(ByVal vPrimaryDO As Double) As Integer
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).Value
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

    Public Function GetJml2_Default(ByVal vPrimariDO As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml_Default(vPrimariDO, vPrimaryCodeNo, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml2_Multi(ByVal vPrimariDO As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml_Multi(vPrimariDO, vPrimaryCodeNo, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml3_Default(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, ByVal vPrimaryCodeNo As Double, _
    ByVal vLotNo As String, ByVal vSecQTy As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml2_Default(vLotNo, vPrimarySO, vPrimaryCodeNo, vPrimaryDO, vSecQTy).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml3_Multi(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, ByVal vPrimaryCodeNo As Double, _
    ByVal vLotNo As String, ByVal vSecQTy As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml2_Multi(vLotNo, vPrimarySO, vPrimaryCodeNo, vPrimaryDO, vSecQTy).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoDO As String) As Double
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoDO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vNoDO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustomerPrimary(ByVal vPrimaryDO As Double) As Double
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable4TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCustomer(vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCustomer(vPrimaryDO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

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

    Public Function GetUserModifiedBy(ByVal vPrimaryDo As Double) As String
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable4TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetModifiedBy(vPrimaryDo)) = False Then
                Return DTA.ScalarQuery_GetModifiedBy(vPrimaryDo)
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

    Public Function GetInvoiceNo(ByVal vPrimaryDO As Double) As String
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoInvoice(vPrimaryDO)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoInvoice(vPrimaryDO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalDO(ByVal vPrimaryDO As Double) As Date
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTanggal(vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggal(vPrimaryDO).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDoNo(ByVal vPrimaryDO As Double) As String
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable4TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDONo(vPrimaryDO)) = False Then
                Return DTA.ScalarQuery_GetDONo(vPrimaryDO)
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

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
      ByVal DGV As DataGridView, ByVal TBL As DseTransaksiDO.DataTable4DataTable, _
      ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal Non_InvoicedOnly As Boolean = False, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable4TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Non_InvoicedOnly = False Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
                ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTA.FillBy_NoSOLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                    QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                    QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%")
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_NoTransaksiLIKE_NonInvoiceDOOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA.FillBy_TanggalSampaiDari_NonInvoiceDOOnly(TBL, vTanggalSampai, vTanggalDari)
                ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                    QueryResult = DTA.FillBy_NoSOLIKE_NonInvoiceDOOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                    QueryResult = DTA.FillBy_CustCodeLIKE_NonInvoiceOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    QueryResult = DTA.FillBy_ProductCodeLIKE_NonInvoiceDOOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    QueryResult = DTA.FillBy_ProductDescLIKE_NonInvoiceDOOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = DTA.FillBy_ProductSpecLIKE_NonInvoiceDOOnly(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                    QueryResult = DTA.FillBy_CustNameLIKE_NonInvoiceDOInly(TBL, Criteria + "%")
                End If
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

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable4TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalTransaksi(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.ScalarQuery_GetJmlByNoSOLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimaryDO As Double) As Date
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimaryDO).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vdgvItem As DataGridView, _
    ByVal ij As FrmTransaksiDO3, ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_so_itemDataTable
        Dim row As DseTransaksiDO.tbltrans_so_itemRow
        Dim row_dgv As DataGridViewRow
        Dim PrimaryMin As Double = Double.MaxValue
        Dim cmd As String = ""
        Dim Dummy As String = ""
        Dim QtySOB, QtyDOLama, QtyDO, QtyOnHand, QtyBookingInv, QtyPendingDO, QtyBookingInv_Sendiri As Decimal
        Dim List As New List(Of String)
        Dim ListMinus As New List(Of String)
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Dim TBL_DO As New DseTransaksiDO.tbltrans_do_itemDataTable
        Dim DTA_DO As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Dim DTA_On_Hand As New DseTransaksiDOTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBL_On_Hand As New DseTransaksiDO.tblmaster_productqtyDataTable
        Dim row_On_Hand As DseTransaksiDO.tblmaster_productqtyRow
        Dim DTA_Booking_Inv As New DseTransaksiDOTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBL_Booking_Inv As New DseTransaksiDO.tbltrans_booking_invitemDataTable
        Dim row_Booking_Inv As DseTransaksiDO.tbltrans_booking_invitemRow
        Dim DTA_Pending As New DseTransaksiDOTableAdapters.tbltrans_do_item7TableAdapter
        Dim TBL_Pending As New DseTransaksiDO.tbltrans_do_item7DataTable
        Dim row_Pending As DseTransaksiDO.tbltrans_do_item7Row
        Dim TBL_Booking_Sendiri As New DseTransaksiDO.DataTable5DataTable
        Dim DTA_Booking_Sendiri As New DseTransaksiDOTableAdapters.DataTable5TableAdapter
        Dim row_Booking_Sendiri As DseTransaksiDO.DataTable5Row

        Dim Primary_Currency_Acuan As Double = -1
        Dim Tipe_Pajak_Acuan As Boolean = False
        Dim Persen_Pajak_Acuan As Double = 0
        Dim DTA_SO As New DseTransaksiDOTableAdapters.tbltrans_so1TableAdapter
        Dim TBL_SO As New DseTransaksiDO.tbltrans_so1DataTable
        Dim row_SO As DseTransaksiDO.tbltrans_so1Row


        Try
            DTA.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Booking_Inv.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Booking_Sendiri.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_DO.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_On_Hand.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Pending.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_SO.Connection.ConnectionString = DTA.Connection.ConnectionString

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("DO number duplication, please change DO number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("DO number duplication, please change DO number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Untuk simpan perubahan, simulasikan terlebih dahulu data setelah simpan, 
            '   pastikan tidak ada sales order balance data yang negatif

            '4. Dalam 1 do bisa ada beberapa SO dengan syarat:
            'cari primary SO minimum
            For Each row_dgv In vdgvItem.Rows
                If row_dgv.Cells(ij.ColPrimarySO.Index).Value < PrimaryMin Then
                    PrimaryMin = row_dgv.Cells(ij.ColPrimarySO.Index).Value
                End If
            Next
            DTA_SO.Fill(TBL_SO, PrimaryMin)
            row_dgv = vdgvItem.Rows(0)
            List.Add(row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString)
            cmd = "[Primary SO] = " + row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString
            TBL_SO.DefaultView.RowFilter = cmd
            row_SO = TBL_SO.DefaultView(0).Row
            Primary_Currency_Acuan = row_SO.Primary_Currency
            Tipe_Pajak_Acuan = row_SO.Use_Tax
            Persen_Pajak_Acuan = row_SO.Tax

            For Each row_dgv In vdgvItem.Rows
                If List.Contains(row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString) = False Then
                    List.Add(row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString)
                    '   a. % pajak yang digunakan sama
                    '   b. jenis faktur pajak yang digunakan sama
                    '   c. primary currency sama
                    cmd = "[Primary SO] = " + row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString
                    TBL_SO.DefaultView.RowFilter = cmd
                    row_SO = TBL_SO.DefaultView(0).Row
                    If Tipe_Pajak_Acuan <> row_SO.Use_Tax Then
                        MsgBox("There are sales order item conflict with each other. " + _
                        "For several sales order item written in one delivery order their currency and tax must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If Persen_Pajak_Acuan <> row_SO.Tax Then
                        MsgBox("There are sales order item conflict with each other. " + _
                        "For several sales order item written in one delivery order their currency and tax must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If Primary_Currency_Acuan <> row_SO.Primary_Currency Then
                        MsgBox("There are sales order item conflict with each other. " + _
                        "For several sales order item written in one delivery order their currency and tax must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            '5. Pastikan tidak ada sales order balance negatif
            'isikan data so
            List.Clear()
            cmd = ""

            DTA.Fill(TBL, PrimaryMin)
            'isikan data do lama jika simpan perubahan 
            If vBaru = False Then
                DTA_DO.FillBy_PrimaryDO(TBL_DO, vPrimaryDO)
            End If
            For Each row_dgv In vdgvItem.Rows
                cmd = "[Primary SO] =" + row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString + _
                " AND [Primary Product] =" + row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value.ToString

                If List.Contains(cmd) = False Then
                    'untuk 1 so dan 1 product, hanya perlu dihitung 1x
                    List.Add(cmd)

                    'hitung sales order balance
                    'qty sob
                    TBL.DefaultView.RowFilter = cmd
                    If TBL.DefaultView.Count = 0 Then
                        QtySOB = 0
                    Else
                        QtySOB = 0
                        For i As Integer = 0 To TBL.DefaultView.Count - 1
                            row = TBL.DefaultView(i).Row
                            QtySOB += row.SOB_Multi + row.Tolerance_Up
                        Next
                    End If
                    'qty do lama
                    If vBaru = True Then
                        QtyDOLama = 0
                    Else
                        QtyDOLama = Me.GetQtyDO_Multi(row_dgv.Cells(ij.ColPrimarySO.Index).Value, _
                        row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value, TBL_DO)
                    End If
                    'qty do
                    QtyDO = Me.GetQtyDGV_Multi(ij, row_dgv.Cells(ij.ColPrimarySO.Index).Value, _
                    row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value, vdgvItem)
                    'pastikan sales order balance yang baru tidak minus
                    If QtySOB + QtyDOLama - QtyDO < 0 Then
                        ListMinus.Add("Product Code: " + row_dgv.Cells(ij.ColCodeNo.Index).Value + _
                        "; SO Number: " + row_dgv.Cells(ij.ColSONO.Index).Value)
                    End If
                End If
            Next
            If ListMinus.Count > 0 Then
                Dummy = "The delivery order cannot be saved, the following product will have minus sales order balance: " + vbCrLf
                QtySOB = 1
                For Each cmd In ListMinus
                    Dummy = Dummy + QtySOB.ToString + ". " + cmd + vbCrLf
                    QtySOB += 1
                Next
                MsgBox(Dummy, MsgBoxStyle.Information, "Information")

                Return False
            End If

            '6. Pastikan tidak ada available to send yang negatif
            'isikan data:
            'a. On hand
            DTA_On_Hand.Fill(TBL_On_Hand)
            'b. Booking Inv
            DTA_Booking_Inv.Fill(TBL_Booking_Inv)
            'c. Pending DO
            DTA_Pending.Fill(TBL_Pending)
            'd. DO lama --> sudah diisikan diatas, hanya ada isinya kalo simpan perubahan
            'e. Booking Sendiri
            DTA_Booking_Sendiri.Fill(TBL_Booking_Sendiri)

            List.Clear()
            ListMinus.Clear()
            For Each row_dgv In vdgvItem.Rows
                If List.Contains(row_dgv.Cells(ij.ColCodeNo.Index).Value) = False Then
                    'untuk 1 product code hanya perlu dicek 1 kali aja, ga peduli so-nya
                    List.Add(row_dgv.Cells(ij.ColCodeNo.Index).Value)
                    cmd = "[Primary Product] = " + row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value.ToString
                    'isikan data on qty on hand
                    TBL_On_Hand.DefaultView.RowFilter = cmd
                    If TBL_On_Hand.DefaultView.Count > 0 Then
                        row_On_Hand = TBL_On_Hand.DefaultView(0).Row
                        QtyOnHand = row_On_Hand.TotalQty
                    Else
                        QtyOnHand = 0
                    End If
                    'isikan data booking inventory
                    TBL_Booking_Inv.DefaultView.RowFilter = cmd
                    If TBL_Booking_Inv.DefaultView.Count > 0 Then
                        row_Booking_Inv = TBL_Booking_Inv.DefaultView(0).Row
                        QtyBookingInv = row_Booking_Inv.TotalBookingInv
                    Else
                        QtyBookingInv = 0
                    End If
                    'isikan data pending do lama
                    TBL_Pending.DefaultView.RowFilter = cmd
                    If TBL_Pending.DefaultView.Count > 0 Then
                        row_Pending = TBL_Pending.DefaultView(0).Row
                        QtyPendingDO = row_Pending.TotalPending
                    Else
                        QtyPendingDO = 0
                    End If
                    'isikan data booking sendiri
                    cmd = cmd + " AND [Primary SO] =" + row_dgv.Cells(ij.ColPrimarySO.Index).Value.ToString
                    TBL_Booking_Sendiri.DefaultView.RowFilter = cmd
                    If TBL_Booking_Sendiri.DefaultView.Count > 0 Then
                        row_Booking_Sendiri = TBL_Booking_Sendiri.DefaultView(0).Row
                        QtyBookingInv_Sendiri = row_Booking_Sendiri.TotalBookingInv
                    Else
                        QtyBookingInv_Sendiri = 0
                    End If
                    'kurangkan data pending do lama dengan qty do lama
                    QtyPendingDO = QtyPendingDO - Me.GetQtyDO(row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value, TBL_DO)
                    'perhitungan data baru yang akan disimpan
                    QtyDO = Me.GetQtyDGV(ij, row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value, vdgvItem)
                    'perhitungkan data booking inventory dia sendiri

                    'pastikan available to send tidak minus
                    If QtyOnHand - QtyBookingInv - QtyPendingDO - QtyDO + QtyBookingInv_Sendiri < 0 Then
                        If ObjProduct.ApakahGoods(row_dgv.Cells(ij.ColPrimaryCodeNo.Index).Value) = True Then
                            ListMinus.Add(row_dgv.Cells(ij.ColCodeNo.Index).Value.ToString)
                        End If
                    End If
                End If
            Next
            If ListMinus.Count > 0 Then
                Dummy = "The following product will have minus qty available to deliver: " + vbCrLf
                QtySOB = 1
                For Each cmd In ListMinus
                    Dummy = Dummy + QtySOB.ToString + ". " + cmd + vbCrLf
                    QtySOB += 1
                Next
                Dummy = Dummy + "Data cannot be saved."

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")

                Return False
            End If


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            DTA.Dispose()
            List = Nothing
            TBL_DO.Dispose()
            DTA_DO.Dispose()
            ListMinus = Nothing
            DTA_On_Hand.Dispose()
            TBL_On_Hand.Dispose()
            DTA_Booking_Inv.Dispose()
            TBL_Booking_Inv.Dispose()
            DTA_Pending.Dispose()
            TBL_Pending.Dispose()
            ObjProduct = Nothing
            DTA_Booking_Sendiri.Dispose()
            DTA_SO.Dispose()
        End Try
    End Function

    Private Function GetQtyDGV(ByVal vFrm As FrmTransaksiDO3, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
   ByVal vDGV As DataGridView) As Double
        Dim row_DGV As DataGridViewRow
        Dim RetVal As Double = 0

        Try
            For Each row_DGV In vDGV.Rows
                If row_DGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row_DGV.Cells(vFrm.ColPrimarySO.Index).Value = vPrimarySO Then
                    RetVal += CDbl(row_DGV.Cells(vFrm.ColTotalQty.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyDGV(ByVal vFrm As FrmTransaksiDO3, ByVal vPrimaryProduct As Double, _
   ByVal vDGV As DataGridView) As Double
        Dim row_DGV As DataGridViewRow
        Dim RetVal As Double = 0

        Try
            For Each row_DGV In vDGV.Rows
                If row_DGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    RetVal += CDbl(row_DGV.Cells(vFrm.ColTotalQty.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyDGV_Multi(ByVal vFrm As FrmTransaksiDO3, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vDGV As DataGridView) As Double
        Dim row_DGV As DataGridViewRow
        Dim RetVal As Double = 0

        Try
            For Each row_DGV In vDGV.Rows
                If row_DGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row_DGV.Cells(vFrm.ColPrimarySO.Index).Value = vPrimarySO Then
                    RetVal += CDbl(row_DGV.Cells(vFrm.ColTotalQtyMulti.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyDO(ByVal vPrimaryProduct As Double, _
   ByVal vTBL As DseTransaksiDO.tbltrans_do_itemDataTable) As Double
        Dim cmd As String = "[Primary Product] =" + vPrimaryProduct.ToString
        Dim DataView As Data.DataRowView
        Dim row As DseTransaksiDO.tbltrans_do_itemRow
        Dim Retval As Double = 0

        Try
            vTBL.DefaultView.RowFilter = cmd
            For Each DataView In vTBL.DefaultView
                row = DataView.Row
                Retval += row.Total_Qty
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyDO(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
  ByVal vTBL As DseTransaksiDO.tbltrans_do_itemDataTable) As Double
        Dim cmd As String = "[Primary SO] =" + vPrimarySO.ToString + _
        " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim DataView As Data.DataRowView
        Dim row As DseTransaksiDO.tbltrans_do_itemRow
        Dim Retval As Double = 0

        Try
            vTBL.DefaultView.RowFilter = cmd
            For Each DataView In vTBL.DefaultView
                row = DataView.Row
                Retval += row.Total_Qty
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyDO_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vTBL As DseTransaksiDO.tbltrans_do_itemDataTable) As Double
        Dim cmd As String = "[Primary SO] =" + vPrimarySO.ToString + _
        " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim DataView As Data.DataRowView
        Dim row As DseTransaksiDO.tbltrans_do_itemRow
        Dim Retval As Double = 0

        Try
            vTBL.DefaultView.RowFilter = cmd
            For Each DataView In vTBL.DefaultView
                row = DataView.Row
                Retval += row.Total_Qty_Multi
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
   ByVal ij As FrmTransaksiDO3, Optional ByVal AskEmailNotification As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, ij, AskEmailNotification)
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
    ByVal ij As FrmTransaksiDO3, Optional ByVal AskEmailNotification As Boolean = True, Optional ByVal vSetOrderBalance As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim ObjItem As New ClsTransaksiDOItem
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim ObjBookingInv As New ClsTransaksiBookingInv
        Dim Urutan As Double = -1
        'Dim ObjEmail As New E_mail_Module.ClsTransEmail

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vdgvItem, ij, -1) = True Then
                If ObjLS.IsInventoryInstalled = True Then
                    Me.mAuthorizationLevelNeeded = 1
                    Me.mAuthorizationLevel = 0
                End If

                'insert query
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, Me.mTanggalTransaksi)

                If DTA.Insert(Me.mNoTransaksi, Urutan, Me.mPrimaryStockOutSales, _
                Me.mPrimaryCust, Me.mCustAddIdentifier, Me.mTanggalTransaksi, Me.mCetakFaktur, Me.mCetakKwitansi, _
                Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, Me.mAuthorizationLevel, Me.mAuthorizationLevelNeeded) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.SimpanItemDO(vdgvItem, Me.mPrimary, vSetOrderBalance)
                    'hapus data booking
                    ObjBookingInv.HapusBooking(vdgvItem, ij)
                    'kirim email
                    'If AskEmailNotification = True Then
                    '    If ObjEmail.ApakahAdaYgDikirimiDOReady = True Then
                    '        If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '            Me.SendDOReadyNotificationEmail(Me.mNoTransaksi, Me.mModifiedBy)
                    '        End If
                    '    End If
                    'End If

                    'cek berkaitan nomor sama untuk beda bulan
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                        Me.CekNomor1(DTA)
                    End If
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru1(vdgvItem, vUserCode, ij)
                Else
                    MsgBox("DO Number duplication, please change DO number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjBookingInv = Nothing
            'ObjEmail = Nothing
        End Try
    End Function

    Public Function SendDOReadyNotificationEmail(ByVal DoNo As String, ByVal User As String) As Boolean
        'Dim ObjEmail As New E_mail_Module.ClsTransEmail
        'Dim ListUser As New List(Of String)
        'Dim DTA As New E_mail_Module.DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        'Dim TBL As New E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoDataTable
        'Dim row As E_mail_Module.DsTransaksiEmail.tbltrans_email_automated_sendtoRow

        'Try
        '    E_mail_Module.SetConstring_DTAEmail(DTA)
        '    With ObjEmail
        '        .Content = "Delivery order:  " + CStr(DoNo) + " is ready." + vbCrLf + _
        '        "Please prepare the goods refering to this transaction and arrange shipment as soon as possible." + vbCrLf + _
        '        vbCrLf + vbCrLf + "Regards, " + vbCrLf _
        '        + vbCrLf + CStr(User)
        '        .CreatedBy = User
        '        .TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
        '        .Subject = "DO " + CStr(DoNo) + " - Ready to send."
        '    End With

        '    DTA.FillBy_PrimaryEmail(TBL, 1)
        '    For Each row In TBL.Rows
        '        ListUser.Add(row.User_Code)
        '    Next

        '    If ObjEmail.SendEmail(ListUser) = True Then
        '        Return True
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Function

    Private Sub CekNomor1(ByVal DTA As DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mNoUrut
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
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        Me.mNoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, _
                        Me.mTanggalTransaksi)
                    End If
                End If
                LoopCount += 1
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

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
   ByVal ij As FrmTransaksiDO3) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, ij)
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
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal ij As FrmTransaksiDO3) As Boolean
        '    Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        '    Dim ObjItem As New ClsTransaksiDOItem
        '    Dim Urutan As Double = -1
        '    Dim ObjEmail As New E_mail_Module.ClsTransEmail

        '    Try
        '        DTA.Connection.ConnectionString = Core_Module.GetConstring
        '        If Me.ApakahDataBenar(False, dgvItem, ij, Me.mPrimary) = True Then
        '            Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
        '            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, Me.mTanggalTransaksi)

        '            If DTA.Update(Me.mNoTransaksi, Urutan, Me.mPrimaryCust, _
        '            Me.mCustAddIdentifier, Me.mTanggalTransaksi, Me.mCetakFaktur, Me.mCetakKwitansi, _
        '            Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
        '                'update data yang berubah setelah simpan perubahan
        '                Me.mNoUrut = Urutan
        '                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
        '                Me.mModifiedDate = Me.mTimestamp
        '                Me.mModifiedBy = vUserCode
        '                'hapus data item lama
        '                ObjItem.Hapus(Me.mPrimary, Me)
        '                'simpan data item baru
        '                ObjItem.SimpanItemDO(dgvItem, Me.mPrimary)

        '                If ObjEmail.ApakahAdaYgDikirimiDOReady = True Then
        '                    If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
        '                        Me.SendDOReadyNotificationEmail(Me.mNoTransaksi, Me.mModifiedBy)
        '                    End If
        '                End If

        '                'nilai balikan
        '                Return True
        '            Else
        '                Return False
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
        '            'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
        '            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
        '                Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
        '                Return Me.SimpanPerubahan1(dgvItem, vUserCode, ij)
        '            Else
        '                MsgBox("DO number duplication, please change DO number", MsgBoxStyle.Information, "Information")
        '                Return False
        '            End If
        '        Else
        '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '        End If
        '    Finally
        '        DTA.Dispose()
        '        ObjItem = Nothing
        '        ObjEmail = Nothing
        '    End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
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
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_Hapus, UserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntri.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, UserCode, TampilkanPesan)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim vListPrimary As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrder

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, UserCode, TampilkanPesan) = True Then
                Me.GetPrimarySO(vPrimaryTransaksi, vListPrimary)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    ObjSOItem.Set_Order_Balance(vListPrimary)
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetPrimarySO(ByVal vPrimaryDO As Double, ByVal vListPrimarySO As List(Of Double))
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item4TableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_do_item4DataTable
        Dim row As DseTransaksiDO.tbltrans_do_item4Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryDO)
            vListPrimarySO.Clear()
            For Each row In TBL.Rows
                vListPrimarySO.Add(row.Primary_SO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryDO As Double, ByVal vUserCode As String, _
   Optional ByVal vStrictEdit As Boolean = True, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim ObjItemInvoice As New ClsTransaksiSalesInvoiceItem
        Dim TanggalDO As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Suatu transaksi do tidak bisa diedit jika:
            '0. jika strict edit maka harus dicek siapa yang buat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryDO)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryDO)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '1. sudah ada transaksi lanjutannya.
            '   a. stock out sales
            If DTA.ScalarQuery_GetJmlStockOut(vPrimaryDO).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There stock out refering to this transaction, cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '   b. invoice
            If ObjItemInvoice.GetJmlByDO(vPrimaryDO) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There sales invoice:" + ObjItemInvoice.GetSalesInvoiceNo(vPrimaryDO) + _
                    " refering to this transaction, cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '   c. sales return 
            If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).HasValue = True Then
                If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).Value > 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("There are sales return refering to this transaction, cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItemInvoice = Nothing
            ObjCD = Nothing
            ObjCL = Nothing
        End Try
    End Function

    Public Function ApakahQuick(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_salesinvoiceDataTable
        Dim row As DseTransaksiDO.tbltrans_salesinvoiceRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryDO)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.Quick_Invoice = 1 Then
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
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryDO As Double, _
 ByVal vUserCode As String, Optional ByVal vStrictEdit As Boolean = True, _
 Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim ObjItemInvoice As New ClsTransaksiSalesInvoiceItem
        Dim CreatedBy As String = ""
        Dim TanggalDO As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosingItem

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '0. jika strict edit maka cek dulu siapa yang membuat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryDO)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryDO)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'Suatu transaksi do tidak bisa diedit jika:
            '1. sudah ada transaksi lanjutannya.
            '   a. stock out sales
            'If DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).HasValue = True Then
            ' If DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).Value > 0 Then
            ' If TampilkanPesan = True Then
            ' MsgBox("There are stock out refering to this transaction, cannot edit this data.", _
            ' MsgBoxStyle.Information, "Information")
            ' End If
            ' Return False
            ' End If
            'End If
            '   b. invoice
            If ObjItemInvoice.GetJmlByDO(vPrimaryDO) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There are sales invoice: " + ObjItemInvoice.GetSalesInvoiceNo(vPrimaryDO) + _
                    " refering to this transaction, cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '   c. sales return 
            If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).HasValue = True Then
                If DTA.ScalarQuery_GetJmlSalesReturn(vPrimaryDO).Value > 0 Then
                    If TampilkanPesan = True Then
                        MsgBox("There are sales return refering to this transaction, cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            '   d. stock out sales
            If DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).HasValue = True Then
                If DTA.ScalarQuery_GetPrimaryStockOutSales(vPrimaryDO).Value <> -1 Then
                    If TampilkanPesan = True Then
                        MsgBox("There are stock out sales refering to this transaction, cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItemInvoice = Nothing
            ObjCD = Nothing
            ObjCL = Nothing
        End Try
    End Function

    Public Function ApakahSudahAdaStockOutSales(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlDOSudahStockOut(vPrimaryDO).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CetakDOCMSI_Panjanga(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal vPrinterName As String, ByVal vUserCode As String) As Boolean
        Dim report As New CrpCetakTransaksiDOa
        Dim DS As New DseTransaksiDO
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim ObjCustom As New CustomSMESS.ClsTransaksiDo

        Try
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool) = "1" Then
                Return ObjCustom.CetakDO(CRV, vprimaryDO, vUserCode)
            End If
            Dim DTAHeader As New DseTransaksiDOTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAHeader.FillBy_PrimaryDO(DS.TblCetakHeader, vprimaryDO)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiDOTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Connection.ConnectionString = Core_Module.GetConstring
            DTAIsi.FillBy_PrimaryDO(DS.TblCetakIsi, vprimaryDO)
            DTAIsi.Dispose()

            'Isikan No SO
            Dim rowIsi As DseTransaksiDO.TblCetakIsiRow
            Dim rowHeader As DseTransaksiDO.TblCetakHeaderRow
            Dim NoSO As String = ""
            Dim LIst As New List(Of String)
            For Each rowIsi In DS.TblCetakIsi.Rows
                If LIst.Count = 0 Then
                    LIst.Add(rowIsi.Cust_PO_No)
                Else
                    If LIst.Contains(rowIsi.Cust_PO_No) = False Then
                        LIst.Add(rowIsi.Cust_PO_No)
                    End If
                End If
            Next
            Dim i As String
            For Each i In LIst
                If NoSO = "" Then
                    NoSO = i
                Else
                    NoSO += ", " + i
                End If
            Next
            LIst = Nothing
            For Each rowHeader In DS.TblCetakHeader.Rows
                rowHeader.No_SO_Buatan_Sendiri = NoSO
                If rowHeader.Line_2 <> "" Then
                    If rowHeader.Line_1 = "" Then
                        rowHeader.Line_1 = rowHeader.Line_2
                    Else
                        rowHeader.Line_1 = rowHeader.Line_1 + vbCrLf + rowHeader.Line_2
                    End If
                End If
                If rowHeader.Line_3 <> "" Then
                    If rowHeader.Line_1 = "" Then
                        rowHeader.Line_1 = rowHeader.Line_3
                    Else
                        rowHeader.Line_1 = rowHeader.Line_1 + vbCrLf + rowHeader.Line_3
                    End If
                End If
                If rowHeader.Line_4 <> "" Then
                    If rowHeader.Line_1 = "" Then
                        rowHeader.Line_1 = rowHeader.Line_4
                    Else
                        rowHeader.Line_1 = rowHeader.Line_1 + vbCrLf + rowHeader.Line_4
                    End If
                End If
                If rowHeader.Line_5 <> "" Then
                    If rowHeader.Line_1 = "" Then
                        rowHeader.Line_1 = rowHeader.Line_5
                    Else
                        rowHeader.Line_1 = rowHeader.Line_1 + vbCrLf + rowHeader.Line_5
                    End If
                End If
            Next
            '
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_City.ParameterFieldName, ObjCompany.GetCompany_City)

            CRV.ReportSource = report
            CRV.Show()
            If vPrinterName <> "" Then
                report.PrintOptions.PrinterName = vPrinterName
                report.PrintToPrinter(1, False, 1, 100)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjSetting = Nothing
            ObjCompany = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Function CetakDOCMSI_Pendeka(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal vPrinterName As String) As Boolean
        Dim report As New CrpCetakTransaksiDO2a
        Dim DS As New DseTransaksiDO
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1

        Try
            Dim DTAHeader As New DseTransaksiDOTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAHeader.FillBy_PrimaryDO(DS.TblCetakHeader, vprimaryDO)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiDOTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Connection.ConnectionString = Core_Module.GetConstring
            DTAIsi.FillBy_PrimaryDO(DS.TblCetakIsi, vprimaryDO)
            DTAIsi.Dispose()

            'Isikan No SO
            Dim rowIsi As DseTransaksiDO.TblCetakIsiRow
            Dim rowHeader As DseTransaksiDO.TblCetakHeaderRow
            Dim NoSO As String = ""
            Dim LIst As New List(Of String)
            For Each rowIsi In DS.TblCetakIsi.Rows
                If LIst.Count = 0 Then
                    LIst.Add(rowIsi.Cust_PO_No)
                Else
                    If LIst.Contains(rowIsi.Cust_PO_No) = False Then
                        LIst.Add(rowIsi.Cust_PO_No)
                    End If
                End If
            Next
            Dim i As String
            For Each i In LIst
                If NoSO = "" Then
                    NoSO = i
                Else
                    NoSO += ", " + i
                End If
            Next
            LIst = Nothing
            For Each rowHeader In DS.TblCetakHeader.Rows
                rowHeader.No_SO_Buatan_Sendiri = NoSO
            Next
            '
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)

            CRV.ReportSource = report
            CRV.Show()

            If vPrinterName <> "" Then
                report.PrintOptions.PrinterName = vPrinterName
                report.PrintToPrinter(1, False, 1, 100)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub

    Public Function ResetItemDO(ByVal vPrimaryStockOutSales As Double) As Boolean
        Dim dta As New DseTransaksiDOTableAdapters.tbltrans_do_item3TableAdapter
        Dim tbl As New DseTransaksiDO.tbltrans_do_item3DataTable
        Dim tbl2 As New DseTransaksiDO.tbltrans_do_item3DataTable
        Dim row As DseTransaksiDO.tbltrans_do_item3Row
        Dim row2 As DseTransaksiDO.tbltrans_do_item3Row
        Dim vPrimaryDo As Double = -1
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, vPrimaryStockOutSales)
            If tbl.Count <= 0 Then
                Return False
            End If
            For Each row In tbl
                With row
                    vPrimaryDo = .Primary_DO
                    tbl2.DefaultView.RowFilter = "[Primary Product] = " & .Primary_Product.ToString & " and [Batch No] = '" & .Batch_No & "'"
                    If tbl2.DefaultView.Count > 0 Then
                        row2 = tbl2.DefaultView(0).Row
                        row2.Total_Qty += .Total_Qty
                        row2.Total_Qty_Multi += .Total_Qty_Multi
                    Else
                        row2 = tbl2.Newtbltrans_do_item3Row
                        For i As Integer = 0 To tbl.Columns.Count - 1
                            row2.Item(i) = row.Item(i)
                        Next
                        tbl2.Rows.Add(row2)
                    End If
                End With
            Next
            dta.DeleteQuery(vPrimaryDo)
            For Each row2 In tbl2
                With row2
                    dta.InsertQuery(.Primary_DO, .Primary_Product, .Primary_SO, .Batch_No, .Expiry_Date, .Qty_Master, _
                        .Qty_Per_Master, .Primary_Master_Units, .Qty, .Total_Qty, .Primary_Units, .Secondary_Qty, _
                        .Primary_Secondary_Units, .Remarks, .Total_Qty_Multi, .Primary_Units_Multi, .Ratio_Reference, _
                        .Ratio_Default)
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally

        End Try
    End Function
End Class
