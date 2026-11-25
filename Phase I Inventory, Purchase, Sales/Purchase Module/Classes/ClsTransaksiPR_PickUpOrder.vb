Option Explicit On
Option Strict Off
Public Class ClsTransaksiPR_PickUpOrder
    '''member pada form sendiri
    Private mPrimary As Double = -1
    Private mPrimaryRefParty As Double = -1
    Private mNoTransaksi As String = ""
    Private mDate As Date = Now
    Private mUrutan As Double
    Private mRemarks As String = ""
    Private mPrimaryStockOutAdjust As Double = -1
    Private mPrimaryStockInPurchase As Double = -1

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Dim ObjPR As New ClsTransaksiPR
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPR_Pickup_Item As New ClsTransaksiPR_PickUpOrder_Item

    '''khusus untuk antrian
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianPurchaseRequest
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    '''Dummy
    Dim TBL_Dummy As New DseTransaksiPR_PickUpOrder.TBLSetBalanceDataTable
    Dim Row_Dummy As DseTransaksiPR_PickUpOrder.TBLSetBalanceRow

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryRefParty() As Double
        Get
            Return mPrimaryRefParty
        End Get
        Set(ByVal value As Double)
            mPrimaryRefParty = value
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

    Public Property Tanggal() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
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

    Public Property PrimaryStockInPurchase() As Double
        Get
            Return Me.mPrimaryStockInPurchase
        End Get
        Set(ByVal value As Double)
            mPrimaryStockInPurchase = value
        End Set
    End Property

    Public Property PrimaryStockOutAdjust() As Double
        Get
            Return Me.mPrimaryStockOutAdjust
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryStockOutAdjust = value
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

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
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

    '''cek kebenaran data 
    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vFrm As FrmTransaksiPR_PickUpOrder, ByVal vPrimary As Double) As Boolean
        Try
            '''cek data yang salah, double dalam kode
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mDate)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mDate)
                    End If
                End If
            End If

            ''cek kesesuain data pr pickup dengan pr 
            Dim row As DataGridViewRow
            For Each row In vFrm.DgvItem.Rows
                If ObjPR.GetJmlByPrimaryCode(CDec(row.Cells(vFrm.ColItem_PrimaryPR.Index).Value), CDec(row.Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value)) <= 0 Then
                    MsgBox("Product code is nothing this PR Number", MsgBoxStyle.Information, "Information ")
                    Return False
                End If
            Next

            If Me.ApakahDataItemBenar(vFrm) = False Then
                Return False
            End If

            If ApakahQtySudahBenar(vFrm, vPrimary) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function ApakahDataItemBenar(ByVal vfrm As FrmTransaksiPR_PickUpOrder) As Boolean
        Try
            If vfrm.DgvItem.Rows.Count = 0 Then
                MsgBox("Please input product code", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '''apakahQty sudah benar sesuai balance 
    Public Function ApakahQtySudahBenar(ByVal vFrm As FrmTransaksiPR_PickUpOrder, ByVal vPrimary As Double) As Boolean
        Dim row As DataGridViewRow
        Try
            For Each row In vFrm.DgvItem.Rows
                Dim OB As Decimal = ObjPR.GetOB(row.Cells(vFrm.ColItem_PrimaryPR.Index).Value, row.Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value)
                Dim ItemDB As Decimal = Me.GetQty(row.Cells(vFrm.ColItem_PrimaryPR.Index).Value, vPrimary, row.Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value)
                Dim DefPrec As Decimal = Me.ObjProduct.GetAngkaDecimal(row.Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value)
                Dim Maximumnya As Decimal = 0
                Dim OB_StlhEdit As Decimal = 0

                OB_StlhEdit = OB + ItemDB - CDec(row.Cells(vFrm.ColItem_Qty.Index).Value)
                If OB_StlhEdit < 0 Then
                    Maximumnya = CDec(row.Cells(vFrm.ColItem_Qty.Index).Value) + OB_StlhEdit
                    MsgBox("Maximum value you can enter is" & " " & Me.ObjCD.DigitGrouping(Maximumnya, DefPrec) & _
                    " " & row.Cells(vFrm.ColItem_Unit.Index).Value & " ", MsgBoxStyle.OkOnly, "Information")
                    vFrm.DgvItem.ClearSelection()
                    vFrm.DgvItem.Rows(row.Index).Selected = True
                    vFrm.DgvItem.Item(vFrm.ColItem_Qty.Index, row.Index).Value = Me.ObjCD.DigitGrouping(Maximumnya, DefPrec)
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '''mengecek nilai OB di pickup_order disamain dengan ob di pr ada kemungkinan jika ob sudah berubah
    '''sebelum input data
    Public Function ApakahBalanceSudahBenar(ByVal vOB As Decimal, ByVal vPrimaryPR As Double, ByVal vPrimaryCode As Double, ByVal vfrm As FrmTransaksiPR_PickUpOrder) As Boolean
        Dim row As DataGridViewRow
        '''cari nilai Qty di dgvnya dulu
        Dim Qty_Di_DGV As Decimal = 0

        Try
            For Each row In vfrm.DgvItem.Rows
                If row.Cells(vfrm.ColItem_PrimaryPR.Index).Value = vPrimaryPR Then
                    If row.Cells(vfrm.ColItem_PrimaryCodeNo.Index).Value = vPrimaryCode Then
                        '''kalo udah dapet nilai qty berdasarkan primarypr dan primarycode
                        '''tinggal tambahin aja nilai qty didgvnya dengan nilai ob di textbox
                        vOB += CDec(row.Cells(vfrm.ColItem_Qty.Index).Value)
                    End If
                End If
            Next

            'Dim Nilai_OB_diPurchase_Order As Decimal = ObjPR.GetOB(vPrimaryPR, vPrimaryCode)
            'If Nilai_OB_diPurchase_Order <> vOB Then
            '    MsgBox("Please try again because balance already updating in purchase request", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '''mengecek tanggal di pickup order tidak lebih kecil dengan tanggal pembuatan prnya sendiri
    Public Function ApakahTanggalSudahBenar(ByVal vPrimaryPR As Double, ByVal vTglDiPickupOrder As Date) As Boolean
        Try
            Dim Tgl_Di_Purchase_Request As Date = ObjPR.GetTanggalTransaksi(vPrimaryPR)
            If vTglDiPickupOrder.Date < Tgl_Di_Purchase_Request.Date = True Then
                MsgBox("Dating in pickup order must be > or same with date in purchase request", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    '''cek di Transaksi sendiri
    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Try
            If Me.GetJmlByNoTransaksi(vNoTransaksi) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, _
                                        ByVal vSalesStaff As String, ByVal Frm As FrmTransaksiPR_PickUpOrder, _
                                        Optional ByVal vTampilkanPesan As Boolean = True) As Boolean

        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
        Dim ObjStockOut As New Inventory_Module.clsTransaksiStockInOut
        Dim vPrimaryStockOutAdjust As Nullable(Of Double)

        Try
            If Created_By <> vSalesStaff Then
                If vTampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'jika ada pick up order, maka tidak boleh delete
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vPrimaryStockOutAdjust = DTA.ScalarQuery_GetPrimaryStockOutAdjust(vPrimary)
            If vPrimaryStockOutAdjust.HasValue = True Then
                If vPrimaryStockOutAdjust.Value <> -1 Then
                    If vTampilkanPesan = True Then
                        MsgBox("To edit this data please delete stock out adjust: " + ObjStockOut.GetNoTransaksi(vPrimaryStockOutAdjust.Value) + ".", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjStockOut = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal frm As FrmTransaksiPR_PickUpOrder, ByVal vPrimary As Double, _
    ByVal vSalesStaff As String, ByVal vStrNumber As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
        Dim DummyStockOutNo As String = ""
        Dim vPrimaryStockOutAdjust As Nullable(Of Double)
        Dim ObjStockOut As New Inventory_Module.clsTransaksiStockInOut

        Try
            If Me.ApakahSudahAda(vStrNumber) = False Then
                If vTampilkanPesan = True Then
                    MsgBox("The purchase pick up order: " + vStrNumber + " has been deleted by another user, please choose another purchase request.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If


            If Created_By <> vSalesStaff Then
                If vTampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'jika ada pick up order, maka tidak boleh edit
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vPrimaryStockOutAdjust = DTA.ScalarQuery_GetPrimaryStockOutAdjust(vPrimary)
            If vPrimaryStockOutAdjust.HasValue = True Then
                If vPrimaryStockOutAdjust.Value <> -1 Then
                    If vTampilkanPesan = True Then
                        MsgBox("To edit this data please delete stock out adjust: " + ObjStockOut.GetNoTransaksi(vPrimaryStockOutAdjust.Value) + ".", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjStockOut = Nothing
        End Try
    End Function

    Public Function GetPrimaryStockOutAdjust(ByVal vPrimaryPickUpOrder As Double) As Double
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryStockOutAdjust(vPrimaryPickUpOrder)
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

    Public Function GetPrimaryRefParty(ByVal vPrimaryPickUpOrder As Double) As Double
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryRefParty(vPrimaryPickUpOrder).HasValue = True Then
                Primary = DTA.ScalarQuery_GetPrimaryRefParty(vPrimaryPickUpOrder).Value
            End If

            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue Then
                Primary = DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
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

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNoTransaksi(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQty(ByVal vPrimaryPR As Double, ByVal vPrimaryPRPickup As Double, ByVal vPrimaryCode As Double) As Decimal
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Dim Retval As Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetQty(vPrimaryPR, vPrimaryPRPickup, vPrimaryCode)
            If Retval.HasValue = True Then
                Return Retval
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetJmlByNoTransaksi(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi)

            If Retval.HasValue = True Then
                If String.IsNullOrEmpty(Retval) = False Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                 Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Double
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order1TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlNoTransaksiLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                RetVal = DTA.ScalarQuery_GetTglDariSampai(vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByRequestBy Then
                RetVal = DTA.ScalarQuery_GetJmlCodeLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlRemarksLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByNoStockInPurchase Then
                RetVal = DTA.ScalarQuery_GetJmlByStockInPurchase(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByNoStockOutAdjust Then
                RetVal = DTA.ScalarQuery_GetJmlByStockOutAdjustNo(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductCodeNo Then
                RetVal = DTA.ScalarQuery_GetJmlByProductCode(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductDesc Then
                RetVal = DTA.ScalarQuery_GetJmlByProductDesc(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByPRNumber Then
                RetVal = DTA.ScalarQuery_GetJmlByPRNo(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim TanggalAwal, TanggalAkhir As Date
        Dim NoUrut As String
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            TanggalAwal = ObjCD.GetFirstDateofTheMonth(vTanggalTransaksi)
            TanggalAkhir = ObjCD.GetLasDateOfTheMonth(vTanggalTransaksi)
            If DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).HasValue Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If
            If IsNumeric(NoUrut) Then
                Me.mUrutan = CDbl(NoUrut)
            End If
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While
            Return NoUrut + "/" + "PRP" + "/" + CStr(Format(vTanggalTransaksi, "MMM").ToUpper) + "/" + CStr(Format(vTanggalTransaksi, "yyyy"))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
                                ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_PickUpOrder, ByVal vBaru As Boolean) As Boolean

        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ApakahDataBenar(vBaru, vFrm, Me.Primary) = False Then Exit Function

            Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))

            If DTA.Insert(Me.mPrimaryRefParty, Me.mNoTransaksi, _
                          Me.mDate, Me.mUrutan, Me.mRemarks, _
                          vUserCode, Now, vUserCode, _
                          Now, Now, Me.mPrimaryStockOutAdjust, Me.mPrimaryStockInPurchase) > 0 Then

                Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data item
                Me.TBL_Dummy.Rows.Clear()
                Me.ObjPR_Pickup_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm, Me.TBL_Dummy)

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

    Public Function IsiPRDummy(ByVal vPrimaryPickUp As Double) As Boolean
        Dim DTA_Item As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_item3TableAdapter
        Dim TBL_Item As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item3DataTable
        Dim Row_Item As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item3Row

        Try
            DTA_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Item.Fill(TBL_Item, vPrimaryPickUp)
            TBL_Dummy.Clear()
            For Each Row_Item In TBL_Item.Rows
                Row_Dummy = TBL_Dummy.NewTBLSetBalanceRow
                With Row_Item
                    Row_Dummy.Primary_Transaksi = vPrimaryPickUp
                    Row_Dummy.Primary_PR = .primary_pr
                End With
                TBL_Dummy.AddTBLSetBalanceRow(Row_Dummy)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Item.Dispose()
        End Try
    End Function

    Public Function SetBalanceSesuaiDummy() As Boolean
        Dim ListPrimaryPR As New List(Of Double)

        Try
            For Each Row_Dummy In TBL_Dummy
                With Row_Dummy
                    If ListPrimaryPR.Contains(.Primary_PR) = False Then
                        ObjPR.SetOrderBalance1(.Primary_PR, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                        ListPrimaryPR.Add(.Primary_PR)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPR = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
                                                    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_PickUpOrder, _
                                                    ByVal vBaru As Boolean) As Boolean
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim DTA1 As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Dim ObjPR As New ClsTransaksiPR
        Dim PrimaryPRLama As Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ApakahDataBenar(vBaru, vFrm, Me.Primary) = False Then Exit Function

            Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
            PrimaryPRLama = DTA1.ScalarQuery_GetPrimaryPR(mPrimary)

            If DTA.Update(Me.mPrimaryRefParty, Me.mNoTransaksi, _
                          Me.mDate, Me.mUrutan, Me.mRemarks, _
                          vUserCode, Now, Now, Me.mPrimaryStockOutAdjust, Me.mPrimaryStockInPurchase, mPrimary) > 0 Then

                'update data yang berubah setelah simpan perubahan
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                '''isi dulu ke dummy sebelum di hapus itemnya
                Me.TBL_Dummy.Rows.Clear()
                IsiPRDummy(Me.mPrimary)
                'hapus data item lama
                ObjPR_Pickup_Item.HapusItem(Me.mPrimary)
                'simpan data item baru
                Me.ObjPR_Pickup_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm, Me.TBL_Dummy)
                '''setelah di hapus update balancenya ambil data di dummy
                'SetBalanceSesuaiDummy()

                'For Each row In vFrm.DgvItem.Rows
                '    If CDbl(PrimaryPRLama.Value) <> CDbl(row.Cells(vFrm.ColItem_PrimaryPR.Index).Value) Then
                '        ObjPR.SetOrderBalance(PrimaryPRLama.Value, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                '    End If
                'Next
                'nilai balikan
                Return True
            Else
                Return False
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function


    Public Function SimpanBaru(ByVal vDGVItem As DataGridView, _
                                    ByVal vUserCode As String, ByVal Frm As FrmTransaksiPR_PickUpOrder, _
                                    ByVal vBaru As Boolean) As Boolean

        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier.ToString, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_PickUp_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vDGVItem, vUserCode, Frm, vBaru)
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


    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_PickUpOrder, _
                                                ByVal vBaru As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_PickUp_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, vFrm, vBaru)
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
            'ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimary As Double, ByVal vSalesStaff As String, ByVal vFrm As FrmTransaksiPR_PickUpOrder) As Boolean
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'hapus data header, data item otomatis dihapuskan oleh s                    
            Me.IsiPRDummy(vPrimary)

            If Me.ApakahBolehHapus(vPrimary, vSalesStaff, vFrm, True) = True Then
                If DTA.Delete(vPrimary) > 0 Then
                    Me.SetBalanceSesuaiDummy()
                    Return True
                End If
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal usercode As String, ByVal vFrm As FrmTransaksiPR_PickUpOrder) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_PickUp_Hapus, usercode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, usercode, vFrm)
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
            'ObjAntri = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
                                   ByVal vObjCD As Core_Module.ClsCekData, ByVal vFrm As FrmTransaksiPR_PickUpOrder) As Boolean

        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim TBL As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_orderDataTable
        Dim row As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_orderRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_orderRow)
                With row
                    Me.mPrimary = .primary
                    Me.mPrimaryRefParty = .primary_ref_party
                    Me.mNoTransaksi = .no_transaksi
                    Me.mDate = .tanggal
                    Me.mUrutan = .urutan
                    Me.mRemarks = .remarks
                    Me.mPrimaryStockOutAdjust = .primary_stock_out_adjust
                    Me.mPrimaryStockInPurchase = .primary_stock_in_purchase

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                ObjPR_Pickup_Item.IsiDGVItem(vPrimary, vDGVItem, vFrm, vObjCD)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
                                        ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order1DataTable, _
                                        ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
                                        Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                        Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean

        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariTglSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRequestBy Then
                QueryResult = DTA.FillBy_ReqByLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                QueryResult = DTA.FillBy_StockInPurchaseNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockOutAdjust Then
                QueryResult = DTA.FillBy_StockOutAdjustNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPRNumber Then
                QueryResult = DTA.FillBy_PRNo(TBL, Criteria + "%")
            End If

            'update DGV
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(3).Visible = False
                DGV.Columns(1).HeaderText = "Document No"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(4).HeaderText = "Request By"
                DGV.Columns(5).HeaderText = "Remarks"

                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(1).MinimumWidth = 100
                DGV.Columns(2).MinimumWidth = 130
                DGV.Columns(3).MinimumWidth = 80
                DGV.Columns(4).MinimumWidth = 80
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

    Public Function IsiTblPR(ByVal vTBLItem As DseTransaksiPR_PickUpOrder.TblPRDataTable, ByVal vPrimaryRefParty As Double, _
    ByVal vPrimaryPickUpOrder As Double, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiPR_PickUpOrder, _
    ByVal vTBLPR As DseTransaksiPR_PickUpOrder.TblPR1DataTable) As Boolean
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.TblPRTableAdapter
        Dim row As DseTransaksiPR_PickUpOrder.TblPRRow
        Dim DTAItem As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_item1TableAdapter
        Dim TBLItem As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item1DataTable
        Dim rowItem As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item1Row
        Dim rowDGV As DataGridViewRow
        Dim rowPR As DseTransaksiPR_PickUpOrder.TblPR1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(vTBLItem, ClsTransaksiPR.EnumReadyForPurchase.Yes, vPrimaryRefParty)
            'tambah db
            DTAItem.Fill(TBLItem, vPrimaryPickUpOrder)
            For Each rowItem In TBLItem.Rows
                vTBLItem.DefaultView.RowFilter = "primary = " + rowItem.primary_pr.ToString + " AND [primary code no] = " + rowItem.primary_code_no.ToString
                If vTBLItem.DefaultView.Count = 0 Then
                    'tambahkan baris
                    row = vTBLItem.NewTblPRRow
                    row.OB = rowItem.Qty
                    row.primary = rowItem.primary_pr
                    row.primary_code_no = rowItem.primary_code_no
                    row.no_transaksi = rowItem.no_transaksi
                    row.primary_unit = rowItem.primary_unit
                    row.Product_Code = rowItem.Product_Code
                    row.Product_Desc = rowItem.Description
                    row.Unit_Code = rowItem.Unit_Code
                    row._Decimal = rowItem._Decimal
                    row.tanggal = rowItem.tanggal
                    vTBLItem.Rows.Add(row)
                Else
                    'tambahkan qty
                    row = vTBLItem.DefaultView(0).Row
                    row.OB += rowItem.Qty
                End If
            Next

            'minus dgv
            For Each rowDGV In vDGV.Rows
                vTBLItem.DefaultView.RowFilter = "primary = " + CStr(rowDGV.Cells(vFrm.ColItem_PrimaryPR.Index).Value).ToString + _
                " AND [primary code no] = " + CStr(rowDGV.Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value).ToString

                If vTBLItem.DefaultView.Count = 0 Then
                    'ga mungkin, kalo sampe ketemu, maka coding salah
                    MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'kurangi qty
                    row = vTBLItem.DefaultView(0).Row
                    row.OB -= CDec(rowDGV.Cells(vFrm.ColItem_Qty.Index).Value)
                    If row.OB <= 0 Then
                        vTBLItem.Rows.Remove(row)
                    End If
                End If
            Next

            'pindahkan tbl pr ke data table yang akan digunakan
            vTBLPR.Rows.Clear()
            For Each row In vTBLItem.Rows
                vTBLPR.DefaultView.RowFilter = "[Primary PR] = " + row.primary.ToString
                If vTBLPR.DefaultView.Count = 0 Then
                    'tambahkan barisnya
                    rowPR = vTBLPR.NewTblPR1Row
                    rowPR.Primary_PR = row.primary
                    rowPR.No_Transaksi = row.no_transaksi
                    rowPR.Tgl = row.tanggal
                    vTBLPR.Rows.Add(rowPR)
                End If
            Next
            vTBLPR.DefaultView.RowFilter = ""

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Function IsiTblProduct(ByVal vTBL As DseTransaksiPR_PickUpOrder.TblProductDataTable, ByVal vPrimaryPR As Double) As Boolean
        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.TblProductTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryPRAndStatus(vTBL, vPrimaryPR, ClsTransaksiPR.EnumStatus.Purchase_On_Progress)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''''Perhitungan Nilai OB
    'Public Function HitungNilaiOB(ByVal vTBL As DseTransaksiPR_PickUpOrder.TblProductDataTable, _
    '                              ByVal vPrimaryPR As Double, ByVal vPrimaryCode As Double, _
    '                              ByVal vDGV_Item As DataGridView, ByVal vFrm As FrmTransaksiPR_PickUpOrder) As Boolean

    '    Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.TblProductTableAdapter
    '    Dim row As DseTransaksiPR_PickUpOrder.TblProductRow
    '    Dim Nilai_OB_SetelahEdit As Decimal = 0
    '    Dim Nilai_Maximumnya As Decimal = 0

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring

    '        DTA.FillBy_PrimaryPRAndStatus(vTBL, vPrimaryPR, ClsTransaksiPR.EnumStatus.Purchase_On_Progress)

    '        For ij As Integer = 0 To vDGV_Item.Rows.Count - 1
    '            Dim DefPrec As Decimal = ObjProduct.GetAngkaDecimal(vFrm.ColItem_PrimaryCodeNo.Index, ij)
    '            Dim Nilai_Dari_DGV As Decimal = vFrm.DgvItem.Item(vFrm.ColItem_Qty.Index, ij).Value
    '            Dim PrimaryPR As Double = vFrm.DgvItem.Item(vFrm.ColItem_PrimaryPR.Index, ij).Value
    '            Dim PrimaryCode As Double = vFrm.DgvItem.Item(vFrm.ColItem_PrimaryCodeNo.Index, ij).Value
    '            Dim Nilai_OB As Decimal = Me.ObjPR.GetOB(PrimaryPR, PrimaryCode)
    '            Dim Nilai_ItemDB As Decimal = ObjCD.DigitGrouping(Me.GetQty(PrimaryPR, vFrm.TxtPrimary.Text, PrimaryCode), DefPrec)
    '            Dim index As Integer = 0

    '            While index < vTBL.Rows.Count
    '                row = vTBL.DefaultView(index).Row
    '                If row.Primary_Code = PrimaryCode Then
    '                    Nilai_OB_SetelahEdit = Nilai_OB + Nilai_ItemDB - Nilai_Dari_DGV
    '                    If Nilai_OB_SetelahEdit <= 0 Then
    '                        vTBL.Rows.RemoveAt(index)
    '                    End If
    '                End If
    '                index += 1
    '            End While

    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function RefreshCbo(ByVal vTBL As DseTransaksiPR_PickUpOrder.TblProductDataTable, _
                               ByVal vPrimaryPR As Double, ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiPR_PickUpOrder) As Boolean

        Dim DTA As New DseTransaksiPR_PickUpOrderTableAdapters.TblProductTableAdapter
        Dim index As Integer
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryPRAndStatus(vTBL, vPrimaryPR, ClsTransaksiPR.EnumStatus.Purchase_On_Progress)
            While vTBL.Rows.Count - index = 0

                index = +1
            End While


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CetakPR_PickUp_Order(ByVal vPrimary As Double, _
    ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, Optional ByVal vCreatePDF As Boolean = False) As Boolean

        Dim Frm As New FrmTransaksiPR_PickUpOrder_Cetak
        Dim Report As New CRPLaporanPR_Pickup_Order
        Dim DS As New DseTransaksiPR_PickUpOrder

        Dim DTACetak_Header As New DseTransaksiPR_PickUpOrderTableAdapters.TBLCetak_HeaderTableAdapter
        Dim TBLCetak_Header As DseTransaksiPR_PickUpOrder.TBLCetak_HeaderDataTable = DS.TBLCetak_Header
        Dim RowCetak_Header As DseTransaksiPR_PickUpOrder.TBLCetak_HeaderRow

        Dim DTACetak_Item As New DseTransaksiPR_PickUpOrderTableAdapters.TBLCetak_ItemTableAdapter
        Dim TBLCetak_Item As DseTransaksiPR_PickUpOrder.TBLCetak_ItemDataTable = DS.TBLCetak_Item
        Dim RowCetak_Item As DseTransaksiPR_PickUpOrder.TBLCetak_ItemRow

        Dim TBLHeader As DseTransaksiPR_PickUpOrder.TblHeaderDataTable = DS.TblHeader
        Dim TBLItem As DseTransaksiPR_PickUpOrder.TblItemDataTable = DS.TblItem

        Dim rowHeader As DseTransaksiPR_PickUpOrder.TblHeaderRow
        Dim rowItem As DseTransaksiPR_PickUpOrder.TblItemRow

        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim BufferNoPR As String = ""

        Try
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBLHeader.Rows.Clear()
            TBLItem.Rows.Clear()

            '''Fill semua akan di cetak
            DTACetak_Header.Fill(TBLCetak_Header, vPrimary)
            DTACetak_Item.Fill(TBLCetak_Item, vPrimary)
            '''Jika data header ada
            If TBLCetak_Header.DefaultView.Count > 0 Then
                TBLCetak_Header.DefaultView.Sort = "tanggal, [no transaksi]"

                '''looping Header
                For Each DV1 In TBLCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    With RowCetak_Header
                        '''Isi Header
                        rowHeader = TBLHeader.NewTblHeaderRow
                        rowHeader.Primary = .primary
                        rowHeader.No_Transaksi = .no_transaksi
                        rowHeader.Tanggal = Format(.tanggal, "dd MMMM yyyy").ToString
                        rowHeader.Request_By = ObjRefParty.GetCode(.primary_ref_party)
                        rowHeader.Remarks = .remarks
                        TBLHeader.Rows.Add(rowHeader)
                        BufferNoPR = rowHeader.No_Transaksi
                    End With

                    '''Cek dan Looping data Item
                    TBLCetak_Item.DefaultView.RowFilter = Nothing
                    TBLCetak_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak_Header.primary.ToString
                    If TBLCetak_Item.DefaultView.Count > 0 Then
                        For Each DV2 In TBLCetak_Item.DefaultView
                            RowCetak_Item = DV2.Row
                            With RowCetak_Item
                                '''isi item
                                rowItem = TBLItem.NewTblItemRow
                                rowItem.Primary = .primary
                                rowItem.Primary_Transaksi = .primary_transaksi
                                rowItem.Primary_PR = .PrimaryPR
                                rowItem.PR_Number = .PR_Number
                                'If BufferNoPR <> .PR_Number Then
                                '    rowItem.PR_Number = .PR_Number
                                '    BufferNoPR = .PR_Number
                                'Else
                                '    rowItem.PR_Number = ""
                                'End If
                                rowItem.Product_Code = .Product_Code & vbNewLine & .Description
                                rowItem._Decimal = ._Decimal
                                rowItem.Remarks = .remarks
                                rowItem.Qty = Me.ObjCD.DigitGrouping(.qty, rowItem._Decimal) & " " & .Unit_Code
                                TBLItem.Rows.Add(rowItem)
                            End With
                        Next
                    End If
                Next
            End If

            Report.SetDataSource(DS)
            CRV.ReportSource = Report
            CRV.Zoom(95)
            CRV.Refresh()
            CRV.Show()

            If vCreatePDF = True And BufferNoPR <> "" Then
                BufferNoPR = BufferNoPR.Replace("/", "-")
                If My.Computer.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + _
                "\Pick Up Order - " + BufferNoPR + ".pdf") = True Then
                    'hapus filenya
                    My.Computer.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + _
                    "\Pick Up Order - " + BufferNoPR + ".pdf", FileIO.UIOption.OnlyErrorDialogs, _
                    FileIO.RecycleOption.DeletePermanently)
                End If

                Report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, _
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Pick Up Order - " + _
                BufferNoPR + ".pdf")
                CRV.ParentForm.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Header.Dispose()
            DTACetak_Item.Dispose()
            TBLCetak_Header.Dispose()
            TBLCetak_Item.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub
End Class
