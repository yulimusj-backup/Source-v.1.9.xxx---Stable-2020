Option Explicit On
Option Strict Off

Public Class clsMasterProduct_PriceList
    Private mPrimary As Double = -1
    Private mCurrencyCode As String = ""

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property CurrencyCode() As String
        Get
            Return mCurrencyCode
        End Get
        Set(ByVal value As String)
            mCurrencyCode = value
        End Set
    End Property


    Public Sub IsiTBLCurrency(ByVal TBL As DseMasterProduct_EditPriceList.tblmaster_currencyDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_currencyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub FillMasterProductPriceList_Cust_Group(ByVal vTblProd_PC_Cust_Group As DseMasterProduct_EditPriceList.tblmaster_product_pricelist_groupDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTblProd_PC_Cust_Group)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub


    Public Sub IsiTBLCustomerGroup(ByVal TBL As DseMasterProduct_EditPriceList.tblmaster_customer_groupDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_customer_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub


    Public Sub IsiTBLCustomerSalesArea(ByVal TBL As DseMasterProduct_EditPriceList.tblmaster_customer_sales_areaDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_customer_sales_areaTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLSegmentationCode(ByVal TBL As DseMasterProduct_EditPriceList.tblmaster_customer_segmentationDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_customer_segmentationTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLTypeCode(ByVal TBL As DseMasterProduct_EditPriceList.tblmaster_customer_typeDataTable)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_customer_typeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiSemua(ByVal vDGVItem As DataGridView, ByVal vDef_Prec_Disc As Double, _
    ByVal i As frmMasterProduct_PriceList, ByVal vDefPrec As Integer)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.DataTable3TableAdapter
        Dim TBLHasil As New DseMasterProduct_EditPriceList.DataTable3DataTable
        Dim RowHasil, RowHasilNext, rowHasilBaru As DseMasterProduct_EditPriceList.DataTable3Row

        Dim RowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim Buffer As String = ""
        Dim ObjCD As New ClsCekData
        Dim ijk As Double = 0

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'Connection
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'isi koding-nya
            DTA.FillBy_PrimaryCurrency(TBLHasil, i.TxtKodeProduk.Text & "%", i.txtPrimaryCurrency.Text)

            'Koding ini untuk menentukan posisi berdasarkan Max Qty, pindahkan nol ke posisi paling bawah
            ijk = 0
            While ijk < TBLHasil.Rows.Count - 1
                RowHasil = TBLHasil.Rows(ijk)
                RowHasilNext = TBLHasil.Rows(ijk + 1)

                If RowHasil.max_qty = 0 Then
                    If RowHasil.Primary_Product = RowHasilNext.Primary_Product And RowHasil.Primary_Unit = RowHasilNext.Primary_Unit Then
                        'turunkan indexnya
                        rowHasilBaru = TBLHasil.NewDataTable3Row
                        For jk As Integer = 0 To TBLHasil.Columns.Count - 1
                            rowHasilBaru.Item(jk) = RowHasil.Item(jk)
                        Next
                        TBLHasil.Rows.Remove(RowHasil)
                        TBLHasil.Rows.InsertAt(rowHasilBaru, ijk + 1)
                        ijk += 1
                    Else
                        ijk += 1
                    End If
                Else
                    ijk += 1
                End If
            End While


            vDGVItem.Rows.Clear()
            If TBLHasil.Rows.Count > 0 Then
                For Each RowHasil In TBLHasil.Rows
                    i.DGVHarga.Rows.Add()
                    RowDGV = i.DGVHarga.Rows(i.DGVHarga.Rows.Count - 1)
                    With RowHasil
                        RowDGV.Cells(i.ColProductCode.Index).Value = .prod_code
                        RowDGV.Cells(i.ColDescription.Index).Value = .desc_code
                        RowDGV.Cells(i.ColUnit.Index).Value = .unit_code
                        RowDGV.Cells(i.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.unit_price, vDefPrec)
                        RowDGV.Cells(i.ColDiscount.Index).Value = ObjCD.DigitGrouping(.discount, vDef_Prec_Disc)
                        If .max_qty <= 0 Then
                            RowDGV.Cells(i.ColMaxQty.Index).Value = "-"
                        Else
                            RowDGV.Cells(i.ColMaxQty.Index).Value = ObjCD.DigitGrouping(.max_qty, ._Decimal)
                        End If
                        RowDGV.Cells(i.ColPrimaryUnit.Index).Value = .Primary_Unit
                        RowDGV.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDGV.Cells(i.ColPrimary.Index).Value = .primary
                        RowDGV.Cells(i.ColDefrec.Index).Value = .Default_Precision
                        RowDGV.Cells(i.ColDecimal.Index).Value = ._Decimal
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBLHasil.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub IsiByPrimaryProduct(ByVal vDGVItem As DataGridView, ByVal vDef_Prec_Disc As Double, _
   ByVal i As frmMasterProduct_PriceList, ByVal vDefPrec As Integer, ByVal vPrimaryProduct As Double, ByVal vPrimaryCurrency As Double)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.DataTable3TableAdapter
        Dim TBLHasil As New DseMasterProduct_EditPriceList.DataTable3DataTable
        Dim RowHasil, RowHasilNext, rowHasilBaru As DseMasterProduct_EditPriceList.DataTable3Row

        Dim RowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim Buffer As String = ""
        Dim ObjCD As New ClsCekData
        Dim ijk As Double = 0

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'Connection
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'isi koding-nya
            DTA.FillBy_PrimaryProduct(TBLHasil, vPrimaryCurrency, vPrimaryProduct)

            'Koding ini untuk menentukan posisi berdasarkan Max Qty, pindahkan nol ke posisi paling bawah
            ijk = 0
            While ijk < TBLHasil.Rows.Count - 1
                RowHasil = TBLHasil.Rows(ijk)
                RowHasilNext = TBLHasil.Rows(ijk + 1)

                If RowHasil.max_qty = 0 Then
                    If RowHasil.Primary_Product = RowHasilNext.Primary_Product And RowHasil.Primary_Unit = RowHasilNext.Primary_Unit Then
                        'turunkan indexnya
                        rowHasilBaru = TBLHasil.NewDataTable3Row
                        For jk As Integer = 0 To TBLHasil.Columns.Count - 1
                            rowHasilBaru.Item(jk) = RowHasil.Item(jk)
                        Next
                        TBLHasil.Rows.Remove(RowHasil)
                        TBLHasil.Rows.InsertAt(rowHasilBaru, ijk + 1)
                        ijk += 1
                    Else
                        ijk += 1
                    End If
                Else
                    ijk += 1
                End If
            End While


            vDGVItem.Rows.Clear()
            If TBLHasil.Rows.Count > 0 Then
                For Each RowHasil In TBLHasil.Rows
                    vDGVItem.Rows.Add()
                    RowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowHasil
                        RowDGV.Cells(i.ColProductCode.Index).Value = .prod_code
                        RowDGV.Cells(i.ColDescription.Index).Value = .desc_code
                        RowDGV.Cells(i.ColUnit.Index).Value = .unit_code
                        RowDGV.Cells(i.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.unit_price, vDefPrec)
                        RowDGV.Cells(i.ColDiscount.Index).Value = ObjCD.DigitGrouping(.discount, vDef_Prec_Disc)
                        If .max_qty <= 0 Then
                            RowDGV.Cells(i.ColMaxQty.Index).Value = "-"
                        Else
                            RowDGV.Cells(i.ColMaxQty.Index).Value = ObjCD.DigitGrouping(.max_qty, ._Decimal)
                        End If
                        RowDGV.Cells(i.ColPrimaryUnit.Index).Value = .Primary_Unit
                        RowDGV.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDGV.Cells(i.ColPrimary.Index).Value = .primary
                        RowDGV.Cells(i.ColDefrec.Index).Value = .Default_Precision
                        RowDGV.Cells(i.ColDecimal.Index).Value = ._Decimal
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBLHasil.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Private Sub PindahkanTabelGroupKeGeneral(ByVal TBLGeneral As DseMasterProduct_EditPriceList.DataTable3DataTable, _
    ByVal TBLGroup As DseMasterProduct_EditPriceList.tbl_pricelist_groupDataTable)
        Dim DummyKombinasi As String = ""
        Dim DummyIndexAcuan As Integer = 0
        Dim RowGeneral As DseMasterProduct_EditPriceList.DataTable3Row
        Dim RowGroup As DseMasterProduct_EditPriceList.tbl_pricelist_groupRow

        Try
            While TBLGroup.Rows.Count > 0
                RowGroup = TBLGroup.Rows(0)
                If DummyKombinasi <> RowGroup.Primary_Product.ToString + " - " + RowGroup.Primary_Unit.ToString Then
                    'cari index acuan
                    DummyIndexAcuan = 0
                    While DummyIndexAcuan <= TBLGeneral.Rows.Count - 1
                        RowGeneral = TBLGeneral.Rows(DummyIndexAcuan)
                        If RowGeneral.Primary_Product = RowGroup.Primary_Product And _
                        RowGeneral.Primary_Unit = RowGroup.Primary_Unit Then
                            DummyKombinasi = RowGroup.Primary_Product.ToString + " - " + RowGroup.Primary_Unit.ToString
                            Exit While
                        Else
                            DummyIndexAcuan += 1
                        End If
                    End While

                    'just in case tidak ada di general
                    DummyKombinasi = RowGroup.Primary_Product.ToString + " - " + RowGroup.Primary_Unit.ToString

                    'hapuskan datanya pada tabel general
                    TBLGeneral.DefaultView.RowFilter = "[Primary Product] = " + RowGroup.Primary_Product.ToString + _
                    " AND [Primary Unit] = " + RowGroup.Primary_Unit.ToString
                    While TBLGeneral.DefaultView.Count > 0
                        RowGeneral = TBLGeneral.DefaultView(0).Row
                        TBLGeneral.Rows.Remove(RowGeneral)
                    End While
                End If

                'pindahkan datanya ke tabel general
                RowGeneral = TBLGeneral.NewDataTable3Row
                With RowGeneral
                    .Default_Precision = RowGroup.Default_Precision
                    .desc_code = RowGroup.desc_code
                    .discount = RowGroup.discount
                    .max_qty = RowGroup.max_qty
                    .primary = RowGroup.primary
                    .Primary_Currency = RowGroup.Primary_Currency
                    .Primary_Product = RowGroup.Primary_Product
                    .Primary_Unit = RowGroup.Primary_Unit
                    .prod_code = RowGroup.prod_code
                    .Remarks = RowGroup.Remarks
                    .unit_code = RowGroup.unit_code
                    .unit_price = RowGroup.unit_price
                    ._Decimal = RowGroup._Decimal
                End With

                TBLGeneral.Rows.InsertAt(RowGeneral, DummyIndexAcuan)

                'majukan index acuan
                DummyIndexAcuan += 1
                TBLGroup.Rows.Remove(RowGroup)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PindahkanTabelCustKeGeneral(ByVal TBLGeneral As DseMasterProduct_EditPriceList.DataTable3DataTable, _
   ByVal TBLCust As DseMasterProduct_EditPriceList.tbl_pricelist_customerDataTable)
        Dim DummyKombinasi As String = ""
        Dim DummyIndexAcuan As Integer = 0
        Dim RowGeneral As DseMasterProduct_EditPriceList.DataTable3Row
        Dim RowCust As DseMasterProduct_EditPriceList.tbl_pricelist_customerRow

        Try
            While TBLCust.Rows.Count > 0
                RowCust = TBLCust.Rows(0)
                If DummyKombinasi <> RowCust.Primary_Product.ToString + " - " + RowCust.Primary_Unit.ToString Then
                    'cari index acuan
                    DummyIndexAcuan = 0
                    While DummyIndexAcuan <= TBLGeneral.Rows.Count - 1
                        RowGeneral = TBLGeneral.Rows(DummyIndexAcuan)
                        If RowGeneral.Primary_Product = RowCust.Primary_Product And _
                        RowGeneral.Primary_Unit = RowCust.Primary_Unit Then
                            DummyKombinasi = RowCust.Primary_Product.ToString + " - " + RowCust.Primary_Unit.ToString
                            Exit While
                        Else
                            DummyIndexAcuan += 1
                        End If
                    End While

                    'just in case tidak ada di general
                    DummyKombinasi = RowCust.Primary_Product.ToString + " - " + RowCust.Primary_Unit.ToString

                    'hapuskan datanya pada tabel general
                    TBLGeneral.DefaultView.RowFilter = "[Primary Product] = " + RowCust.Primary_Product.ToString + _
                    " AND [Primary Unit] = " + RowCust.Primary_Unit.ToString
                    While TBLGeneral.DefaultView.Count > 0
                        RowGeneral = TBLGeneral.DefaultView(0).Row
                        TBLGeneral.Rows.Remove(RowGeneral)
                    End While
                End If

                'pindahkan datanya ke tabel general
                RowGeneral = TBLGeneral.NewDataTable3Row
                With RowGeneral
                    .Default_Precision = RowCust.Default_Precision
                    .desc_code = RowCust.desc_code
                    .discount = RowCust.discount
                    .max_qty = RowCust.max_qty
                    .primary = RowCust.primary
                    .Primary_Currency = RowCust.Primary_Currency
                    .Primary_Product = RowCust.Primary_Product
                    .Primary_Unit = RowCust.Primary_Unit
                    .prod_code = RowCust.prod_code
                    .Remarks = RowCust.Remarks
                    .unit_code = RowCust.unit_code
                    .unit_price = RowCust.unit_price
                    ._Decimal = RowCust._Decimal
                End With

                TBLGeneral.Rows.InsertAt(RowGeneral, DummyIndexAcuan)

                'majukan index acuan
                DummyIndexAcuan += 1
                TBLCust.Rows.Remove(RowCust)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub isiCustomerGroup(ByVal vKodeProduct As String, ByVal vPrimaryCurrency As Double, _
    ByVal vPrimaryGroup As Double, ByVal vDGVItem As DataGridView, ByVal vDef_Prec_Disc As Double, _
    ByVal i As frmMasterProduct_PriceList, ByVal vDefPrec As Integer)
        Dim DTAGeneral As New DseMasterProduct_EditPriceListTableAdapters.DataTable3TableAdapter
        Dim TBLHasil As New DseMasterProduct_EditPriceList.DataTable3DataTable
        Dim RowHasil, RowHasilNext, rowHasilBaru As DseMasterProduct_EditPriceList.DataTable3Row
        Dim DTAGroup As New DseMasterProduct_EditPriceListTableAdapters.tbl_pricelist_groupTableAdapter
        Dim TBLGroup As New DseMasterProduct_EditPriceList.tbl_pricelist_groupDataTable
        Dim RowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim ijk As Double = 0

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAGeneral.Connection.ConnectionString = Core_Module.GetConstring
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring

            DTAGeneral.FillBy_PrimaryCurrency(TBLHasil, vKodeProduct & "%", vPrimaryCurrency)
            DTAGroup.FillBy_Group(TBLGroup, vKodeProduct & "%", vPrimaryCurrency, vPrimaryGroup)

            'gabungkan data per group dengan data general
            Me.PindahkanTabelGroupKeGeneral(TBLHasil, TBLGroup)
            'Koding ini untuk menentukan posisi berdasarkan Max Qty, pindahkan nol ke posisi paling bawah
            ijk = 0
            While ijk < TBLHasil.Rows.Count - 1
                RowHasil = TBLHasil.Rows(ijk)
                RowHasilNext = TBLHasil.Rows(ijk + 1)

                If RowHasil.max_qty = 0 Then
                    If RowHasil.Primary_Product = RowHasilNext.Primary_Product And RowHasil.Primary_Unit = RowHasilNext.Primary_Unit Then
                        'turunkan indexnya
                        rowHasilBaru = TBLHasil.NewDataTable3Row
                        For jk As Integer = 0 To TBLHasil.Columns.Count - 1
                            rowHasilBaru.Item(jk) = RowHasil.Item(jk)
                        Next
                        TBLHasil.Rows.Remove(RowHasil)
                        TBLHasil.Rows.InsertAt(rowHasilBaru, ijk + 1)
                        ijk += 1
                    Else
                        ijk += 1
                    End If
                Else
                    ijk += 1
                End If
            End While

            'Masukan data kedalam DGV Item
            vDGVItem.Rows.Clear()
            If TBLHasil.Rows.Count > 0 Then
                For Each RowHasil In TBLHasil.Rows
                    vDGVItem.Rows.Add()
                    RowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowHasil
                        RowDGV.Cells(i.ColProductCode.Index).Value = .prod_code
                        RowDGV.Cells(i.ColDescription.Index).Value = .desc_code
                        RowDGV.Cells(i.ColUnit.Index).Value = .unit_code
                        RowDGV.Cells(i.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.unit_price, vDefPrec)
                        RowDGV.Cells(i.ColDiscount.Index).Value = ObjCD.DigitGrouping(.discount, vDef_Prec_Disc)
                        If .max_qty = 0 Then
                            RowDGV.Cells(i.ColMaxQty.Index).Value = "-"
                        Else
                            RowDGV.Cells(i.ColMaxQty.Index).Value = ObjCD.DigitGrouping(.max_qty, ._Decimal)
                        End If
                        RowDGV.Cells(i.ColPrimaryUnit.Index).Value = .Primary_Unit
                        RowDGV.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDGV.Cells(i.ColPrimary.Index).Value = .primary
                        RowDGV.Cells(i.ColDefrec.Index).Value = .Default_Precision
                        RowDGV.Cells(i.ColDecimal.Index).Value = ._Decimal
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGeneral.Dispose()
            DTAGroup.Dispose()
            TBLGroup.Dispose()
            TBLHasil.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub IsiCustomer(ByVal vKodeProduct As String, ByVal vPrimaryCurrency As Double, _
    ByVal vPrimaryCustomer As Double, ByVal vDGVItem As DataGridView, ByVal vDef_Prec_Disc As Double, _
    ByVal i As frmMasterProduct_PriceList, ByVal vTipeGrup As frmMasterProduct_PriceList.EnumCustomerGroupTipe, ByVal vDefPrec As Integer)
        Dim DTAGeneral As New DseMasterProduct_EditPriceListTableAdapters.DataTable3TableAdapter
        Dim TBLHasil As New DseMasterProduct_EditPriceList.DataTable3DataTable
        Dim RowHasil, RowHasilNext, rowHasilBaru As DseMasterProduct_EditPriceList.DataTable3Row
        Dim DTAGroup As New DseMasterProduct_EditPriceListTableAdapters.tbl_pricelist_groupTableAdapter
        Dim TBLGroup As New DseMasterProduct_EditPriceList.tbl_pricelist_groupDataTable
        Dim rowGroup As DseMasterProduct_EditPriceList.tbl_pricelist_groupRow

        Dim DTACustomer As New DseMasterProduct_EditPriceListTableAdapters.tbl_pricelist_customerTableAdapter
        Dim TBLCustomer As New DseMasterProduct_EditPriceList.tbl_pricelist_customerDataTable
        Dim ROWCustomer As DseMasterProduct_EditPriceList.tbl_pricelist_customerRow

        Dim RowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData

        Dim ListKombinasi As New List(Of String)
        Dim ijk As Double = 0
        Dim vPrimaryGroup As Double = -1
        Dim ObjCust As New clsMasterCustomer

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAGeneral.Connection.ConnectionString = Core_Module.GetConstring
            DTAGroup.Connection.ConnectionString = DTAGeneral.Connection.ConnectionString
            DTACustomer.Connection.ConnectionString = Core_Module.GetConstring

            DTAGeneral.FillBy_PrimaryCurrency(TBLHasil, vKodeProduct & "%", vPrimaryCurrency)
            DTACustomer.FillBy_Customer(TBLCustomer, vPrimaryCurrency, vKodeProduct & "%", vPrimaryCustomer)
            If vTipeGrup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group Then
                vPrimaryGroup = ObjCust.GetPrimary_CustGroup(vPrimaryCustomer)
            ElseIf vTipeGrup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Sales_Area Then
                vPrimaryGroup = ObjCust.GetPrimary_CustSalesArea(vPrimaryCustomer)
            ElseIf vTipeGrup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Segmentation Then
                vPrimaryGroup = ObjCust.GetPrimary_CustSegmentasi(vPrimaryCustomer)
            ElseIf vTipeGrup = frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Type Then
                vPrimaryGroup = ObjCust.GetPrimary_CustType(vPrimaryCustomer)
            End If
            DTAGroup.FillBy_Group(TBLGroup, vKodeProduct & "%", vPrimaryCurrency, vPrimaryGroup)

            'aturan main menampilkan data harga
            '1. jika ada di customer, maka gunakan data harga customer ybs
            '2. jika tidak ada di customer, maka gunakan harga yg di grup
            '3. jika tidak ada maka gunakan harga general
            'untuk bs mencapai hal diatas, maka berikut hal yang dilakukan:
            'a. buang data per group yang sama dengan data per customer
            'b. pindahkan data per group ke dalam general
            'c. pindahkan data per customer ke dalam general
            'eksekusi step a
            For Each ROWCustomer In TBLCustomer.Rows
                If ListKombinasi.Contains(ROWCustomer.Primary_Product.ToString + " - " + ROWCustomer.Primary_Unit.ToString) = False Then
                    ListKombinasi.Add(ROWCustomer.Primary_Product.ToString + " - " + ROWCustomer.Primary_Unit.ToString)
                    TBLGroup.DefaultView.RowFilter = "[Primary Product] = " + ROWCustomer.Primary_Product.ToString + _
                    " AND [Primary Unit] = " + ROWCustomer.Primary_Unit.ToString
                    While TBLGroup.DefaultView.Count > 0
                        rowGroup = TBLGroup.DefaultView(0).Row
                        TBLGroup.Rows.Remove(rowGroup)
                    End While
                End If
            Next
            'eksekusi step b
            Me.PindahkanTabelGroupKeGeneral(TBLHasil, TBLGroup)
            'eksekusi step c
            Me.PindahkanTabelCustKeGeneral(TBLHasil, TBLCustomer)

            'Koding ini untuk menentukan posisi berdasarkan Max Qty, pindahkan nol ke posisi paling bawah
            ijk = 0
            While ijk < TBLHasil.Rows.Count - 1
                RowHasil = TBLHasil.Rows(ijk)
                RowHasilNext = TBLHasil.Rows(ijk + 1)

                If RowHasil.max_qty = 0 Then
                    If RowHasil.Primary_Product = RowHasilNext.Primary_Product And RowHasil.Primary_Unit = RowHasilNext.Primary_Unit Then
                        'turunkan indexnya
                        rowHasilBaru = TBLHasil.NewDataTable3Row
                        For jk As Integer = 0 To TBLHasil.Columns.Count - 1
                            rowHasilBaru.Item(jk) = RowHasil.Item(jk)
                        Next
                        TBLHasil.Rows.Remove(RowHasil)
                        TBLHasil.Rows.InsertAt(rowHasilBaru, ijk + 1)
                        ijk += 1
                    Else
                        ijk += 1
                    End If
                Else
                    ijk += 1
                End If
            End While

            'Masukan data kedalam DGV Item
            vDGVItem.Rows.Clear()
            If TBLHasil.Rows.Count > 0 Then
                For Each RowHasil In TBLHasil.Rows
                    vDGVItem.Rows.Add()
                    RowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowHasil
                        RowDGV.Cells(i.ColProductCode.Index).Value = .prod_code
                        RowDGV.Cells(i.ColDescription.Index).Value = .desc_code
                        RowDGV.Cells(i.ColUnit.Index).Value = .unit_code
                        RowDGV.Cells(i.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.unit_price, vDefPrec)
                        RowDGV.Cells(i.ColDiscount.Index).Value = ObjCD.DigitGrouping(.discount, vDef_Prec_Disc)
                        If .max_qty = ObjCD.DigitGrouping(0, ._Decimal) Then
                            RowDGV.Cells(i.ColMaxQty.Index).Value = "-"
                        Else
                            RowDGV.Cells(i.ColMaxQty.Index).Value = ObjCD.DigitGrouping(.max_qty, ._Decimal)
                        End If
                        RowDGV.Cells(i.ColPrimaryUnit.Index).Value = .Primary_Unit
                        RowDGV.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDGV.Cells(i.ColPrimary.Index).Value = .primary
                        RowDGV.Cells(i.ColDefrec.Index).Value = .Default_Precision
                        RowDGV.Cells(i.ColDecimal.Index).Value = ._Decimal
                        RowDGV.Cells(i.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ' vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        Finally
            DTAGeneral.Dispose()
            DTACustomer.Dispose()
            TBLCustomer.Dispose()
            TBLHasil.Dispose()
            ObjCD = Nothing
            ListKombinasi = Nothing
            ObjCust = Nothing
            DTAGroup.Dispose()
            TBLGroup.Dispose()
        End Try
    End Sub


    Public Sub SimpanGeneral(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, ByVal vMaxqty As Double, _
    ByVal vPrimarycurrency As Double, ByVal vPrimaryunit As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelistTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryunit, vRemarks)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Sub


    Public Function HapusGeneral(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelistTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdateGeneral(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, _
    ByVal vMaxqty As Double, ByVal vPrimarycurrency As Double, ByVal vPrimaryunit As Double, ByVal vPrimary As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelistTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Update(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryunit, vRemarks, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Sub

    Public Sub SimpanCustomer(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, _
    ByVal vMaxqty As Double, ByVal vPrimarycurrency As Double, ByVal vPrimaryCustomer As Double, ByVal vPrimaryunit As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryCustomer, vPrimaryunit, vRemarks)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Sub

    Public Sub HapusItemCustomer(ByVal vPrimary As Double)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function HapusCustomer(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusSatuanCustomer(ByVal vPrimaryCurrency As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByProdCustUnit(vPrimaryProduct, vPrimaryUnit, vPrimaryCurrency)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdateCustomer(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, ByVal vMaxqty As Double, _
    ByVal vPrimarycurrency As Double, ByVal vPrimaryCustomer As Double, ByVal vPrimaryunit As Double, ByVal vprimary As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Update(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryCustomer, vPrimaryunit, vRemarks, vprimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Sub

    Public Sub SimpanGroupCustomer(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, _
    ByVal vMaxqty As Double, ByVal vPrimarycurrency As Double, ByVal vPrimaryGroup As Double, ByVal vTipeGroup As Integer, ByVal vPrimaryunit As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryGroup, vTipeGroup, vPrimaryunit, vRemarks)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Sub

    Public Function HapusGroupCustomer(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByPrimaryGroup(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HapusSatuanGroupCustomer(ByVal vPrimaryCurrency As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_produnitcur(vPrimaryProduct, vPrimaryUnit, vPrimaryCurrency)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub HapusItemGroup(ByVal vPrimary As Double)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateGroupCustomer(ByVal vprimaryproduct As Double, ByVal vUnitprice As Decimal, ByVal vDiscount As Decimal, _
    ByVal vMaxqty As Double, ByVal vPrimarycurrency As Double, ByVal vPrimaryGroup As Double, ByVal vTipeGroup As Integer, _
    ByVal vPrimaryunit As Double, ByVal vPrimary As Double, ByVal vRemarks As String)
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Update(vprimaryproduct, vUnitprice, vDiscount, vMaxqty, vPrimarycurrency, vPrimaryGroup, vTipeGroup, vPrimaryunit, vRemarks, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetMaxPrimaryGeneral() As Double
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelistTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMaxPrimary()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetmaxPrimaryGroup() As Double
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMaxPrimary()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetMaxPrimaryCustomer() As Double
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMaxPrimary()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDefPrecCurrency(ByVal vPrimaryCurrency As Double) As Integer
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_currencyTableAdapter
        Dim temp As Nullable(Of Long)
        Try
            temp = DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryCurrency)
            If temp.HasValue = True Then
                Return temp
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        End Try
    End Function

    Public Function ApakahAdaCust(ByVal vPrimaryProduct As Double, ByVal vPrimaryCurrency As Double, ByVal vPrimaryCust As Double, ByVal vPrimaryUnit As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_customerTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimaryCurrency, vPrimaryUnit, vPrimaryCust)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
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

    Public Function ApakahAdaCustGroup(ByVal vPrimaryProduct As Double, ByVal vPrimaryCurrency As Double, ByVal vPrimaryCustGroup As Double, _
    ByVal vPrimaryUnit As Double) As Boolean
        Dim DTA As New DseMasterProduct_EditPriceListTableAdapters.tblmaster_product_pricelist_groupTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimaryCurrency, vPrimaryUnit, vPrimaryCustGroup)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
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
End Class
