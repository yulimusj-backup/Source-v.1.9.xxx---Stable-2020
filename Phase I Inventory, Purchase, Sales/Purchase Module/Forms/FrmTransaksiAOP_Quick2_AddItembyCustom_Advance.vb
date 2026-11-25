Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTransaksiAOP_Quick2_AddItembyCustom_Advance
   
    Public PrimaryTransaksi As Double = -1
    Public PrimaryLocation As Double = -1
    Public TglTransaksi As Date = #1/1/1950#

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public NoAOP As String = ""
    Public CurrencyCode As String = ""
    Public PrimaryCurrency As Double = -1
    Public PrimaryCurrTax As Double = -1
    Public PrimaryTransaksiAOP As Double = -1
    Public PrimarySupplier As Double = -1
    Public SupplierCode As String = ""
    Public baru As Boolean = True
    Public TanggalTransaksi As Date = Now
    Public UserCode As String = ""
    Public TaxExchageRate As String = ""
    Public CanEditForm As Boolean = True
    Public PrimaryPO As Double = -1
    Public TblItemCustomInOut As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Public TblItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable
    Public CalculateByIndex As Integer = 0


    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProdCust As New CustomSMESS.ClsMasterProduct
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim ObjStockInOutItem As New clsTransaksiStockInOutItem
    Dim ObjItemInOut As New ClsTransaksiStockInOut
    Dim ObjReference As New ClsTransaksiStockInPurchase
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim objAopItem As New ClsTransaksiAOP
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjPO As New ClsTransaksiPO
    Dim tblCboUnit As New DseTransaksiPO.CboUnitDataTable
    Dim rowCboUnit As DseTransaksiPO.CboUnitRow
    Dim tblMultiUnit As New DseTransaksiPO.tblmaster_product_item_multiunitDataTable
    Dim rowMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitRow

    Dim DefaultPrecision As Integer = 0
    Dim Buffer As String = ""
    Dim BufferDetail As String = ""
    Dim FromOthBtn As Boolean = False
    Dim cancelClose As Boolean = False
    Dim AktifkanSelectedDgvItem As Boolean = False
    Dim buffer2 As String = ""
    Dim GetHistoricalPrice As Boolean = True
    Enum EnumEditCol As Integer
        Edit_QtyMulti
        Edit_QtyDef
        Edit_QtyRoll
    End Enum

    Private Sub FrmTransaksiAOP_Quick2_AddItembyCustom_Advance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If cancelClose = True Then
                cancelClose = False
                e.Cancel = True
                Exit Sub
            End If

            If FromOthBtn = False Then
                Dim hasil As MsgBoxResult = MsgBox("Add item before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Exit Sub
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FromOthBtn = False
        End Try
    End Sub
    Private Function ApakahDataBenar() As Boolean
        Dim DTAItem As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBLItem As New DseTransaksiAOP.tbltrans_aop_itemDataTable
        Dim rowItemD As DseTransaksiAOP.tbltrans_aop_itemRow
        Dim primaryStockIn As Double = -1
        Dim ObjGRN As New ClsTransaksiGRN
        Dim RowDummy As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Try
            'For Each rd As DataGridViewRow In Me.dgvItem.Rows
            '    If CDec(rd.Cells(Me.ColInQtyMult.Index).Value) > CDec(rd.Cells(Me.ColOrderBalanceQty.Index).Value) Then
            '        Me.dgvItem.ClearSelection()
            '        rd.Selected = True
            '        MsgBox("In qty multi can't be greater than order balance qty.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            '    If CDec(rd.Cells(Me.ColInQtyMulti.Index).Value) > 0 And CDec(rd.Cells(Me.ColInQtyDefault.Index).Value) <= 0 Then
            '        Me.dgvItem.ClearSelection()
            '        rd.Selected = True
            '        MsgBox("You can't add this data because In qty multi you entered make qty default=0.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                With rd
                    If rd.Cells(Me.ColIsKain.Index).Value = True Then
                        If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                            MsgBox("Please specify price by " & rd.Cells(Me.ColUnit.Index).Value & " for row: " & rd.Index + 1, MsgBoxStyle.Information, "Information")
                            rd.Selected = True
                            Return False
                        End If
                    End If
                End With
            Next
            Dim RowItem As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
            Dim ListProd As New List(Of Double)
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If ListProd.Contains(rd.Cells(Me.ColPrimaryProd.Index).Value) = False Then
                    If rd.Cells(Me.ColIsLocation.Index).Value = 0 Then
                        Me.TblItemCustom.DefaultView.RowFilter = "[Primary Product] = " & CDbl(rd.Cells(Me.ColPrimaryProd.Index).Value).ToString
                        For Each dv As DataRowView In Me.TblItemCustom.DefaultView
                            RowItem = dv.Row
                            RowItem.Primary_Location = -1
                            RowItem.Storage_Area_Code = ""
                        Next
                    End If

                    ListProd.Add(rd.Cells(Me.ColPrimaryProd.Index).Value)
                End If
            Next
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.PrimaryTransaksiAOP)
            If TBLItem.Rows.Count > 0 Then
                rowItemD = TBLItem.Rows(0)
                primaryStockIn = ObjGRN.GetPrimaryStockIn(rowItemD.Primary_GRN)
            End If
            Dim TblProdQty As New DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
            Dim ListProduct As New List(Of Double)

            If Me.ObjItemInOut.ApakahDataBenar_CustomStockInOut(primaryStockIn, Me.baru, _
            ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In, TblProdQty, Me.TblItemCustomInOut, _
            ListProduct, ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    

    Private Sub FrmTransaksiAOP_Quick2_AddItembyCustom_Advance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ColProductCode.Frozen = True

            Me.TxtLocation.CharacterCasing = CharacterCasing.Upper
            Me.txtPrimaryLocation.Visible = False
            Me.colIsBatch.Visible = False
            Me.ColIsKain.Visible = False
            Me.ColIsLocation.Visible = False
            Me.colIsExpDate.Visible = False
            Me.ColPrimaryUnit.Visible = False
            Me.ColPrimaryUnitMulti.Visible = False
            Me.ColRatioDef.Visible = False
            Me.ColRatioMulti.Visible = False
            Me.ColDec.Visible = False
            Me.ColDecMulti.Visible = False
            Me.ColPrimaryPO.Visible = False
            Me.ColRemarksItem.Visible = False

            Me.ColTotalQtyMulti.MaxInputLength = 21
            Me.ColTotalQty.MaxInputLength = 21
            'Me.Col2Qty.MaxInputLength = 11
            Me.Col2SecondaryQty.MaxInputLength = 11

            Me.Col2SecondaryQty.HeaderText = "@"
            Me.Col2SecondaryUnit.HeaderText = "@ Unit"
            Me.ColPriceByMultiUnit.Visible = False
            Me.ColATCost.Visible = False
            Me.ColShowDetail.ReadOnly = False
            Me.ColShowDetail.Visible = False
            Me.ColDiscAmount.Visible = True
            Me.ColDiscAmount.ReadOnly = False
            Me.dgvItem.Columns(Me.ColDiscAmount.Index).DisplayIndex = 18



            AktifkanSelectedDgvItem = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll


            If Me.PrimaryLocation <> -1 Then
                Me.txtPrimaryLocation.Text = PrimaryLocation
                Me.TxtLocation.Text = ObjLoc.GetStorageAreaCode(Me.txtPrimaryLocation.Text)

                Me.TxtLocation.Enabled = False
                Me.btnFindLocation.Enabled = False
            End If
            Me.TxtCurrencyCode1.Text = Me.CurrencyCode
            Me.TxtPrimaryCurrency1.Text = Me.PrimaryCurrency
            Me.KosongkanItem()
            Me.txtPrimaryLocation.Text = "-1"
            Me.TxtLocation.Text = ""
            Me.TxtLocation.Focus()
            Me.IsiDgvItem()
            AktifkanSelectedDgvItem = True
            Me.dgvItem.ClearSelection()
            If Me.dgvItem.RowCount > 0 Then
                Me.dgvItem.Rows(0).Selected = True
            End If
            Me.CboUnitMulti.Items.Clear()
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.ReadOnly = False
            Me.SetCalculateBy()
            Me.TxtLocation.Focus()
            Me.txtAmount.Visible = False
            Me.txtAmount.TabStop = False
            Me.LblAmount.Visible = False
            Me.FormatDGV2()

         
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function IsiDgvItem() As Boolean

        Dim RowCustom As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim rowDgv As DataGridViewRow
        Dim ObjPO As New ClsTransaksiPO
        Dim PONo As String = ""
        Dim adaLokasi As Boolean = False
        Dim PrimaryRoll As Double = 22
        Dim unitRoll As String = ""
        Dim listPrimaryProduct As New List(Of Double)
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim decDef As Integer = 0
        Dim decMulti As Integer = 0
        Dim decRoll As Integer = 0
        Dim AngkaDecimalCurrency As Double = 0
        Try

            Me.dgvItem.Rows.Clear()
            Me.TblItemCustom.DefaultView.RowFilter = ""
            Me.TblItemCustom.DefaultView.Sort = "[Product Code] ASC"

            unitRoll = Me.objUnit.GetUnitCode(PrimaryRoll)

            For i As Integer = 0 To Me.TblItemCustom.DefaultView.Count - 1
                RowCustom = TblItemCustom.DefaultView(i).Row
                AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(RowCustom.Primary_Currency)
                With RowCustom
                    If listPrimaryProduct.Contains(RowCustom.Primary_Product) = False Then
                        listPrimaryProduct.Add(RowCustom.Primary_Product)
                        Me.dgvItem.Rows.Add()
                        rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        rowDgv.Cells(Me.ColIsKain.Index).Value = .Is_Kain
                        decDef = Me.ObjProd.GetAngkaDecimal(.Primary_Product)
                        decMulti = Me.ObjProd.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        decRoll = Me.ObjProd.GetAngkaDecimal(.Primary_Product, PrimaryRoll)
                        rowDgv.Cells(Me.ColDec.Index).Value = decDef
                        If .Is_Kain = True Then
                            rowDgv.Cells(Me.ColDecMulti.Index).Value = decRoll
                        Else
                            rowDgv.Cells(Me.ColDecMulti.Index).Value = decMulti
                        End If

                        rowDgv.Cells(Me.ColNo.Index).Value = (rowDgv.Index + 1).ToString
                        rowDgv.Cells(Me.ColProductCode.Index).Value = .Product_Code
                        rowDgv.Cells(Me.ColDescription.Index).Value = .Description
                        rowDgv.Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                        rowDgv.Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                        rowDgv.Cells(Me.ColPONO.Index).Value = .PO_No
                        rowDgv.Cells(Me.ColGRNNO.Index).Value = .GRN_No
                        rowDgv.Cells(Me.ColStockOutNo.Index).Value = .Stock_In_Purchase
                        rowDgv.Cells(Me.ColSuppDONo.Index).Value = .Supplier_Do_No
                        rowDgv.Cells(Me.ColQtyMaster.Index).Value = .Qty_Master
                        rowDgv.Cells(Me.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDgv.Cells(Me.ColMasterUnits.Index).Value = ""
                        rowDgv.Cells(Me.ColQty.Index).Value = .Qty
                        rowDgv.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, decDef)
                        rowDgv.Cells(Me.ColUnit.Index).Value = .Units
                        rowDgv.Cells(Me.ColSecQty.Index).Value = -1
                        rowDgv.Cells(Me.ColSecUnit.Index).Value = ""
                        rowDgv.Cells(Me.ColCurrency.Index).Value = .Currency_Code
                        rowDgv.Cells(Me.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, AngkaDecimalCurrency)
                        rowDgv.Cells(Me.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDgv.Cells(Me.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, AngkaDecimalCurrency)
                        rowDgv.Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, AngkaDecimalCurrency)
                        rowDgv.Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, AngkaDecimalCurrency)
                        rowDgv.Cells(Me.ColATCost.Index).Value = ObjCD.DigitGrouping(.At_Price, 0)
                        rowDgv.Cells(Me.ColShowDetail.Index).Value = "Show"
                        rowDgv.Cells(Me.ColRemarksItem.Index).Value = .Remarks
                        rowDgv.Cells(Me.ColPrimaryProd.Index).Value = .Primary_Product
                        rowDgv.Cells(Me.ColPrimaryPO.Index).Value = .Primary_PO
                        rowDgv.Cells(Me.ColPrimaryGRN.Index).Value = .Primary_GRN
                        rowDgv.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        rowDgv.Cells(Me.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDgv.Cells(Me.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDgv.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        If .Storage_Area_Code = "" Then
                            .Storage_Area_Code = "-"
                        End If
                        rowDgv.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location

                        rowDgv.Cells(Me.ColReferedToPO.Index).Value = .Refered_To_PO
                        If rowDgv.Cells(Me.ColReferedToPO.Index).Value = 0 Then
                            rowDgv.Cells(Me.ColPONO.Index).Value = ""
                        End If
                        rowDgv.Cells(Me.ColAtAmount.Index).Value = .At_Amount
                        rowDgv.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        If .Primary_Units_Multi = .Primary_Units Then
                            rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, Me.ObjProd.GetAngkaDecimal(.Primary_Product))
                        Else
                            If .Is_Kain = True Then
                                rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, decRoll)

                            Else
                                rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, decMulti)
                            End If

                        End If
                        rowDgv.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rowDgv.Cells(Me.ColUnitsCodeMulti.Index).Value = .Unit_Code_Multi
                        rowDgv.Cells(Me.ColRatioMulti.Index).Value = .Ratio_Reference
                        rowDgv.Cells(Me.ColRatioDef.Index).Value = .Ratio_Default
                        rowDgv.Cells(Me.ColPrimaryUnitMulti.Index).Value = .Primary_Units_Multi
                        rowDgv.Cells(Me.ColItemAdjustment.Index).Value = ObjCD.DigitGrouping(.Adjustment, .Decimal_Currency)
                        rowDgv.Cells(Me.ColItemAdjustmentReason.Index).Value = .Reason
                        rowDgv.Cells(Me.ColItemAdjustmentUser.Index).Value = .User_Code
                        If .Adjustment_Timestamp = #1/1/1950# Then
                            rowDgv.Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"
                        Else
                            rowDgv.Cells(Me.ColItemAdjustmentTimestamp.Index).Value = Format(.Adjustment_Timestamp, "dd MMM yyyy HH:mm")
                        End If



                        If .Primary_Location <> -1 Then
                            Me.txtPrimaryLocation.Text = .Primary_Location
                            Me.TxtLocation.Text = .Storage_Area_Code
                            'Me.txtPrimaryWarehouseFrom.Text = .primary_warehouse_from
                            'Me.txtWareHouseCodeFrom.Text = .warehouse_from
                        End If

                        Me.ObjProd.GetProductProperties(.Primary_Product, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
                        If LocationCont = False Then
                            rowDgv.Cells(Me.ColIsLocation.Index).Value = "0"
                        Else
                            rowDgv.Cells(Me.ColIsLocation.Index).Value = "1"
                        End If
                        If BatchNoCont = False Then
                            rowDgv.Cells(Me.colIsBatch.Index).Value = "0"
                        Else
                            rowDgv.Cells(Me.colIsBatch.Index).Value = "1"
                        End If
                        rowDgv.Cells(Me.colIsExpDate.Index).Value = ExpiryDateCont
                        If .Primary_Units <> .Primary_Units_Multi Then
                            If .Is_Kain = True Then
                                rowDgv.Cells(Me.ColRatioDesc.Index).Value = rowDgv.Cells(ColTotalQtyMulti.Index).Value _
                                & " " & .Unit_Code_Multi & " = " & rowDgv.Cells(ColTotalQty.Index).Value & " " & .Unit_Code

                            Else
                                rowDgv.Cells(ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, rowDgv.Cells(Me.ColDecMulti.Index).Value) _
                                  & " " & .Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, rowDgv.Cells(Me.ColDec.Index).Value) & " " & .Unit_Code
                                rowDgv.Cells(ColRatioMulti.Index).Value = .Ratio_Reference
                                rowDgv.Cells(ColRatioDef.Index).Value = .Ratio_Default
                            End If
                        Else
                            rowDgv.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDgv.Cells(Me.ColDec.Index).Value = decDef
                        If .Is_Kain = True Then
                            rowDgv.Cells(Me.ColDecMulti.Index).Value = decRoll
                        Else
                            rowDgv.Cells(Me.ColDecMulti.Index).Value = decMulti
                        End If
                        rowDgv.Cells(Me.ColPriceByMultiUnit.Index).Value = .Price_by_Multi_Unit
                        If .Price_by_Multi_Unit = "1" Then
                            rowDgv.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            rowDgv.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
                        Me.Hitung_Amount_Multi(rowDgv)
                        'Me.Hitung_Total_Multi()
                    Else
                        rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        If .Is_Kain = True Then
                            rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Roll + CDec(rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value), Me.ObjProd.GetAngkaDecimal(.Primary_Product, PrimaryRoll))
                        Else
                            rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty_Multi + CDec(rowDgv.Cells(Me.ColTotalQtyMulti.Index).Value), Me.ObjProd.GetAngkaDecimal(.Primary_Product, PrimaryRoll))
                        End If
                        rowDgv.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty + CDec(rowDgv.Cells(Me.ColTotalQty.Index).Value), Me.ObjProd.GetAngkaDecimal(.Primary_Product))
                        Me.Hitung_Amount_Multi(rowDgv)
                        'Me.Hitung_Total_Multi()
                    End If
                End With
            Next

            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.ObjCD.SetNumber(Me.dgvItem)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
  
    Private Sub TxtLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.GotFocus
        Me.TxtLocation.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtLocation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = Obj.GetPrimary(Me.TxtLocation.Text)
                    Me.TxtCodeNo.Focus()
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.TxtLocation.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = obj.GetPrimary(Me.TxtLocation.Text)
                    Me.TxtCodeNo.Focus()
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocation.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.LostFocus
        Me.TxtLocation.BackColor = Me.txtPrimaryLocation.BackColor
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea
        i.BackColorAll = Me.BackColorAll

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocation.Text <> "" Then
                i.txtCari.Text = Me.TxtLocation.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                Me.TxtCodeNo.Focus()
            Else
                Me.TxtLocation.Clear()
                Me.TxtLocation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim objPO As New ClsTransaksiPO

        Try
            If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
            e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index Or e.ColumnIndex = Me.ColShowDetail.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Then
                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColTotalQty.Index Then

                    If Me.dgvItem(Me.colIsBatch.Index, e.RowIndex).Value <> 0 Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    If Me.dgvItem.Item(Me.ColIsKain.Index, e.RowIndex).Value = True Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    If Me.dgvItemDetail.Visible = True Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    If String.IsNullOrEmpty(Me.TxtLocation.Text) = True And Me.dgvItem(Me.ColIsLocation.Index, e.RowIndex).Value <> 0 Then
                        MsgBox("Please enter location first.", MsgBoxStyle.Information, "Information")
                        Me.TxtLocation.Focus()
                        e.Cancel = True
                        Exit Sub
                    End If
                End If


                Me.Buffer = Me.dgvItem(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
   
    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim row As DataGridViewRow
        Dim AngkaDecimal, IndexTujuan As Integer
        Try
            'If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick2_AddByCustomProduct") = True Then
            '    Exit Sub
            'End If


            If e.ColumnIndex = Me.ColShowDetail.Index Then
                If e.RowIndex <> -1 Then
                    Me.ShowDetailCalculation(Me.dgvItem.Item(Me.ColProductCode.Index, _
                        e.RowIndex).Value, e.RowIndex)
                End If
                'ElseIf e.ColumnIndex = Me.ColRatioDesc.Index Then
                '    If e.RowIndex <> -1 Then
                '        If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                '                Me.TxtRemarkItem.Enabled = False Then
                '            Exit Sub
                '        End If

                '        Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                '        Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                '        Dim BufferRatioDesc As String = Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value

                '        frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                '        frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                '        frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                '        frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                '        frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                '        frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                '        If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                '            Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                '            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                '            row = Me.dgvItem.Rows(e.RowIndex)
                '            Me.SetRatioAndQtyDefault(row, True)
                '            Me.Hitung_Amount_Multi(row)
                '            Me.Hitung_AtCost_Multi()

                '            If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                '                MsgBox("Ratio value restored to previous value, because ratio you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                '                Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                '                Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                '                Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                '                Me.SetQtyDefaultOnDgv(e.RowIndex, Me.dgvItem.Item(Me.ColIsKain.Index, e.RowIndex).Value)
                '            End If

                '        End If
                '    End If
            ElseIf e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index And Me.TxtRemarkItem.Enabled = True Then

                If Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
                        'hitung total
                        Me.Hitung_Total_Multi()
                    End If

                    FrmRef = Nothing
                End If
                UpdatePriceByProduct(e.RowIndex)

                Dim rd_Item As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
                Dim rd_Detail As New DataGridViewRow
                AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryLocation.Text, -1, "", rd_Item, rd_Detail, False, False, rd_Item.Cells(Me.ColIsLocation.Index).Value)

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ShowDetailCalculation(ByVal CodeNo As String, ByVal rowIndex As Integer)
        Dim DetailCalculation As String = ""
        Dim row As DataGridViewRow
        Dim TotalOtherCost As Decimal = 0
        Dim TotalFOBPrice As Decimal = 0
        Dim SubTotal, NettPrice As Decimal
        Dim Currency As String = Me.dgvItem.Item(Me.ColCurrency.Index, rowIndex).Value.ToString
        Dim i As New FrmTransaksiAOPDetailCalc
        Dim AcuanRow As Integer

        Try
            'lakukan perhitungan FOB price
            DetailCalculation = "FOB Price(1) = Sub Total Before DP" + vbCrLf
            TotalFOBPrice = (Me.GetSubTotal_TanpaReturn_Multi * CDbl(Me.TaxExchageRate))
            DetailCalculation = DetailCalculation + "                     = IDR " + Me.ObjCD.DigitGrouping(TotalFOBPrice) + vbCrLf + vbCrLf

            'hitung other cost
            DetailCalculation = DetailCalculation + "Other Cost(2) = "
            If Me.TxtRemarkItem.Enabled = False Then
                AcuanRow = Me.dgvCost.Rows.Count
            Else
                AcuanRow = Me.dgvCost.Rows.Count - 1
            End If
            If (Me.TxtRemarkItem.Enabled = True And Me.dgvCost.Rows.Count > 1) Or _
            (Me.TxtRemarkItem.Enabled = False And Me.dgvCost.RowCount > 0) Then
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + row.Cells(Me.ColDescCost.Index).Value.ToString
                        Else
                            DetailCalculation = DetailCalculation + " + " + row.Cells(Me.ColDescCost.Index).Value.ToString
                        End If
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = "
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + "IDR (" + row.Cells(Me.ColCostNominal.Index).Value.ToString + _
                            " * " + row.Cells(Me.ColExcRateCost.Index).Value.ToString + ")"
                        Else
                            DetailCalculation = DetailCalculation + " + IDR (" + row.Cells(Me.ColCostNominal.Index).Value.ToString + _
                            " * " + row.Cells(Me.ColExcRateCost.Index).Value.ToString + ")"
                        End If
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = "
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + "IDR " + row.Cells(Me.ColCostAmount.Index).Value.ToString
                        Else
                            DetailCalculation = DetailCalculation + " + IDR " + row.Cells(Me.ColCostAmount.Index).Value.ToString
                        End If
                    End If
                Next
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        TotalOtherCost += row.Cells(Me.ColCostAmount.Index).Value
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalOtherCost))
                DetailCalculation = DetailCalculation + vbCrLf + vbCrLf
            Else
                DetailCalculation = DetailCalculation + " IDR 0" + vbCrLf + vbCrLf
            End If

            'hitung total cost
            DetailCalculation = DetailCalculation + " Total Cost(3) = FOB Price(1) + Other Cost(2)" + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice)) _
            + " + IDR " + Me.ObjCD.DigitGrouping(CStr(TotalOtherCost)) + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice + TotalOtherCost)) + vbCrLf + vbCrLf

            'lakukan perhitungan unit price
            DetailCalculation = DetailCalculation + "Unit price = Nett Price / SubTotal * Total Cost(3) " + vbCrLf
            For Each row In Me.dgvItem.Rows
                SubTotal += row.Cells(Me.ColAmount.Index).Value
            Next
            If Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryProd.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1) <> 0 Then
                NettPrice = Me.dgvItem.Item(Me.ColNettPrice.Index, rowIndex).Value
            Else
                NettPrice = 0
            End If

            DetailCalculation = DetailCalculation + "                = " + Currency + " (" + Me.dgvItem.Item(Me.ColAmount.Index, rowIndex).Value.ToString + " / " + _
            Currency + Me.ObjCD.DigitGrouping(SubTotal.ToString) + " * IDR " + Me.ObjCD.DigitGrouping((TotalFOBPrice + TotalOtherCost).ToString) + ")/ " + _
            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryProd.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1)) + " " + Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, rowIndex).Value.ToString + vbCrLf
            DetailCalculation = DetailCalculation + "                = IDR " + Me.dgvItem.Item(Me.ColATCost.Index, rowIndex).Value.ToString
            If Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryProd.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1) = 0 Then
                DetailCalculation = DetailCalculation + vbCrLf + "(All returned, AT Cost =0)"
            End If
            i.TxtDetailCalculation.Text = DetailCalculation
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim UbahByQtyDef As Boolean = False
        Dim RowIdxDgvAtas As Integer = -1
        Dim AngkaDecimal, IndexTujuan As Integer
        Try
            'yang masuk ke sini pasti buka product kain
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If

                RowIdxDgvAtas = e.RowIndex
                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
                    'untuk nilai qty dan secondary qty, pastikan nilainya positif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for qty in multi data.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) = 0 Then
                        MsgBox("Unit price must be > 0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Dim JmlQtyDef As Decimal = 0
                    JmlQtyDef = Me.ObjCD.DigitGrouping((CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) / CDec(Me.dgvItem.Item(Me.ColRatioMulti.Index, e.RowIndex).Value) * CDec(Me.dgvItem.Item(Me.ColRatioDef.Index, e.RowIndex).Value)), Me.dgvItem.Item(Me.ColDec.Index, e.RowIndex).Value)
                    'cek apakah edit sah
                    If Me.ApakahEditSah("", #1/1/1950#, JmlQtyDef, -1, EnumEditCol.Edit_QtyMulti, RowIdxDgvAtas, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'tutup
                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem.Item(Me.ColDecMulti.Index, e.RowIndex).Value)
                    If e.ColumnIndex = Me.ColUnitPrice.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                    End If

                    Me.HitungItem(e.RowIndex, AngkaDecimal)
                ElseIf e.ColumnIndex = Me.ColTotalQty.Index Then
                    'untuk nilai qty dan secondary qty, pastikan nilainya positif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for qty in default data.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'cek apakah edit sah
                    If Me.ApakahEditSah("", #1/1/1950#, CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), -1, EnumEditCol.Edit_QtyDef, RowIdxDgvAtas, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'tutup

                    Dim DecDumm As Integer = Me.dgvItem.Item(Me.ColDec.Index, e.RowIndex).Value
                    If Me.dgvItem(Me.ColIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value = True Then
                        DecDumm = Me.dgvItem.Item(Me.ColDecMulti.Index, e.RowIndex).Value
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DecDumm)
                    UbahByQtyDef = True
                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                    Me.HitungItem(e.RowIndex, AngkaDecimal)
                ElseIf e.ColumnIndex = Me.ColDiscount.Index Then
                    'pastikan nilai discount non negatif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                    'pastikan nilai diskon maksimum = 100
                    If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "1"
                    Me.HitungItem(e.RowIndex, 2)
                ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                    'pastikan nilai discount non negatif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                    'pastikan nilai diskon < Unit Price
                    If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0"
                    Me.HitungItem(e.RowIndex, AngkaDecimal)
                End If
                UpdatePriceByProduct(e.RowIndex)
            Else
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If

            Dim rd_Item As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            Dim rd_Detail As New DataGridViewRow
            AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
            Me.HitungItem(Me.dgvItem.SelectedCells(0).RowIndex, AngkaDecimal)
            Me.Hitung_Total_Multi()
            If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColTotalQty.Index Then
                Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryLocation.Text, -1, "", rd_Item, rd_Detail, False, UbahByQtyDef, rd_Item.Cells(Me.ColIsLocation.Index).Value)
                Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)
            End If
           
            

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub HitungItem(ByVal vRowIndex As Integer, ByVal rounding As Integer)
        Dim rowDgv As DataGridViewRow
        Dim DefaultPrecisionRow As Integer = 0
        Dim DiscPerct As Integer = 0

        Try
            rowDgv = Me.dgvItem.Rows(vRowIndex)
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDgv.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = rowDgv.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                rowDgv.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) * CDec(rowDgv.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                rowDgv.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            rowDgv.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) - CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, vRowIndex).Value = 1 Then
                Me.dgvItem.Item(Me.ColAmount.Index, vRowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIndex).Value) * _
                CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIndex).Value), rounding))
            Else
                'Me.dgvItem.Item(Me.ColAmount.Index, vRowIndex).Value = _
                'Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIndex).Value) * _
                'CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIndex).Value), rounding))

                Me.dgvItem.Item(Me.ColAmount.Index, vRowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIndex).Value) * _
                                CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIndex).Value), rounding))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub UpdateDgvItemByDgvDetail(ByVal rd As DataGridViewRow, _
   ByVal vTblItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable)
        Dim QtyMulti As Decimal = 0
        Dim QtyDef As Decimal = 0
        Dim PrimaryPO As Double = -1 'Me.dgvItem(Me.colPrimaryPO.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim PrimaryProd As Double = -1 'Me.dgvItem(Me.colPrimaryProd.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim DecMulti As Integer = 0 'Me.dgvItem(Me.colDecMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim DecDef As Integer = 0 'Me.dgvItem(Me.colDec.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim PrimaryUnit As Integer = -1 'Me.dgvItem(Me.colPrimaryUnit.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim PrimaryUnitMulti As Integer = -1 'Me.dgvItem(Me.ColPrimaryUnitMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim IsKain As Boolean = False 'Me.dgvItem(Me.ColPrimaryUnitMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim UnitDef As String = "" 'Me.dgvItem(Me.ColPrimaryUnitMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim UnitMulti As String = "" 'Me.dgvItem(Me.ColPrimaryUnitMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
        Dim CountRows As Integer = 0
        Dim RatioDef As Decimal = 0
        Dim RatioMulti As Decimal = 0
        Dim AngkaDecimal As Integer = 0
        Try
            PrimaryPO = Me.dgvItem(Me.ColPrimaryPO.Index, rd.Index).Value
            PrimaryProd = Me.dgvItem(Me.ColPrimaryProd.Index, rd.Index).Value
            DecMulti = Me.dgvItem(Me.ColDecMulti.Index, rd.Index).Value
            DecDef = Me.dgvItem(Me.ColDec.Index, rd.Index).Value
            PrimaryUnit = Me.dgvItem(Me.ColPrimaryUnit.Index, rd.Index).Value
            PrimaryUnitMulti = Me.dgvItem(Me.ColPrimaryUnitMulti.Index, rd.Index).Value
            IsKain = Me.dgvItem(Me.ColIsKain.Index, rd.Index).Value
            UnitDef = Me.dgvItem(Me.ColUnit.Index, rd.Index).Value
            UnitMulti = Me.dgvItem(Me.ColUnitsCodeMulti.Index, rd.Index).Value

            Me.GetQtyTotalPerPOnProduct(PrimaryProd, QtyMulti, QtyDef, RatioMulti, RatioDef, CountRows, vTblItemCustom)

            rd.Cells(ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(QtyMulti, DecMulti)
            rd.Cells(ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(QtyDef, DecDef)
            If PrimaryUnit <> PrimaryUnitMulti Then
                If IsKain = True Then
                    If QtyMulti <> 0 And QtyDef <> 0 Then
                        rd.Cells(ColRatioDesc.Index).Value = rd.Cells(Me.ColTotalQtyMulti.Index).Value _
                        & " " & UnitMulti & " = " & rd.Cells(Me.ColTotalQty.Index).Value & " " & UnitDef
                    Else
                        rd.Cells(Me.ColRatioDesc.Index).Value = "-"
                    End If
                End If
            Else
                rd.Cells(Me.ColRatioDesc.Index).Value = "-"
            End If

            'Me.HitungItem(rd.Index, AngkaDecimal)
            'Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub GetQtyTotalPerPOnProduct(ByVal vPrimaryProd As Double, ByRef vQtyMulti As Decimal, _
        ByRef vQtyDef As Decimal, ByRef vRatioMulti As Decimal, ByRef vRatioDef As Decimal, ByRef vCountRows As Integer, _
        ByVal vTblItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable)
        Dim rowItem As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row

        Try
            vQtyMulti = 0
            vQtyDef = 0
            vCountRows = 0
            vRatioMulti = 0
            vRatioDef = 0

            vTblItemCustom.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProd.ToString
            For Each dv As DataRowView In vTblItemCustom.DefaultView
                rowItem = dv.Row
                vCountRows += 1
                vQtyMulti += rowItem.Total_Qty_Multi
                vQtyDef += rowItem.Total_Qty
                vRatioMulti = rowItem.Ratio_Reference
                vRatioDef = rowItem.Ratio_Default
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function ApakahEditSah(ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal jmlDGV As Decimal, ByVal vQtyAtDumm As String, ByVal IdxColEdit As EnumEditCol, _
    ByVal RowIdxDgvAtas As Integer, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim QtyAtDumm As Decimal = -1
        Dim PrimaryLocation As Double = -1
        Dim PrimaryProduct As Double = -1
        Dim PrimaryPO As Double = -1
        Dim PrimaryPdO As Double = -1
        Dim RowAtas As DataGridViewRow = Me.dgvItem.Rows(RowIdxDgvAtas)
        Dim SecondaryQty As Decimal = -1

        Try
            PrimaryProduct = RowAtas.Cells(Me.colPrimaryProd.Index).Value
            PrimaryPO = RowAtas.Cells(Me.colPrimaryPO.Index).Value
            If RowAtas.Cells(Me.colIsLocation.Index).Value <> 0 Then
                PrimaryLocation = Me.txtPrimaryLocation.Text
            End If

            If Me.ObjProd.ApakahGoods(PrimaryProduct) = False Then
                Return True
            End If

            If IsNumeric(vQtyAtDumm) = True Then
                If Me.ObjCD.CekBilanganPositif(vQtyAtDumm) = True Then
                    QtyAtDumm = vQtyAtDumm
                End If
            Else
                QtyAtDumm = -1
            End If

            JmlGudang = Me.ObjStockInOutItem.GetJml(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, QtyAtDumm)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objAopItem.GetJml(Me.PrimaryTransaksi, PrimaryProduct, PrimaryPO, QtyAtDumm)
            Me.ObjProd.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlItem_Default(PrimaryPO, PrimaryProduct, BatchNo, vQtyAtDumm, PrimaryLocation, RowAtas.Cells(Me.colIsLocation.Index).Value)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGVTotal + jmlDGV
            JmlDGVLain = JmlDGVTotal - jmlDGV

            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make your stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    Dim MustQtyInput As String = ""
                    If IdxColEdit = EnumEditCol.Edit_QtyMulti Then
                        Dim RatioMulti As Decimal = CDec(RowAtas.Cells(Me.colRatioMulti.Index).Value)
                        Dim RatioDef As Decimal = CDec(RowAtas.Cells(Me.colRatioDef.Index).Value)
                        Dim PrimaryUnitMulti As Double = CDec(RowAtas.Cells(Me.ColPrimaryUnitMulti.Index).Value)
                        MustQtyInput = Me.ObjCD.DigitGrouping(((jmlDGV - JmlSetelahHapus) * RatioMulti) / RatioDef, Me.ObjProd.GetAngkaDecimal(PrimaryProduct, PrimaryUnitMulti)) + " " + RowAtas.Cells(Me.ColUnitsCodeMulti.Index).Value + "."
                    ElseIf IdxColEdit = EnumEditCol.Edit_QtyDef Then
                        MustQtyInput = Me.ObjCD.DigitGrouping(jmlDGV - JmlSetelahHapus, Me.ObjProd.GetAngkaDecimal(PrimaryProduct)) + " " + RowAtas.Cells(Me.ColUnit.Index).Value + "."
                    ElseIf IdxColEdit = EnumEditCol.Edit_QtyRoll Then
                        MustQtyInput = Me.ObjCD.DigitGrouping((jmlDGV - JmlSetelahHapus) / vQtyAtDumm, 0) + " Roll."
                    End If
                    MsgBox("Minimum value you must can enter is " + MustQtyInput, MsgBoxStyle.Information, "Information")
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetJmlItem_Default(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, ByVal vBatchNo As String, ByVal vQtyAt As String, _
    ByVal vPrimaryLoc As Double, ByVal vIsLocation As Boolean) As Decimal
        Dim rowItem As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim Filter As String = ""
        Dim QtyAt As Decimal = -1
        Dim TotalQtyItem As Decimal = 0
        Dim PrimaryLocation As Double = -1

        Try
            If IsNumeric(vQtyAt) = True Then
                If vQtyAt > 0 Then
                    QtyAt = CDec(vQtyAt)
                End If
            End If
            If vIsLocation = True Then
                PrimaryLocation = vPrimaryLoc
            End If

            Filter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [BatchNumber] = '" & vBatchNo.ToString & _
            "' AND [Qty At] = " & QtyAt.ToString & " AND [Primary Location] = " & PrimaryLocation
            Me.TblItemCustom.DefaultView.RowFilter = Filter
            For i As Integer = 0 To Me.TblItemCustom.DefaultView.Count - 1
                rowItem = Me.TblItemCustom.DefaultView(i).Row
                With rowItem
                    If .Primary_PO <> vPrimaryPO Then
                        TotalQtyItem += .Total_Qty
                    End If
                End With
            Next

            Return TotalQtyItem
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub UpdateTblInOutItem(ByVal vTblItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
       ByVal vPrimaryTrans As Double, ByVal vPrimaryLocation As Double, ByVal vQtyAt As Decimal, ByVal vBatch As String, ByVal rd_item As DataGridViewRow, _
       ByVal rd_detail As DataGridViewRow, ByVal vIsDgvDetail As Boolean, ByVal vUbahDiQtyDef_DgvAtas As Boolean, ByVal Is_LocControl As Integer, Optional ByVal Hapus As Boolean = False)
        Dim rowItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim Filter As String = ""
        Dim QtyIn_New As Decimal = 0
        Dim IsSimpan As Boolean = False

        Try
            If Hapus = False Then
                If vIsDgvDetail = False Then
                    Filter = "[Primary Product] = " & _
                    rd_item.Cells(Me.ColPrimaryProd.Index).Value.ToString
                    QtyIn_New = rd_item.Cells(Me.ColTotalQtyMulti.Index).Value
                    If vUbahDiQtyDef_DgvAtas = True Then
                        QtyIn_New = rd_item.Cells(Me.ColTotalQty.Index).Value
                    End If
                Else
                    Filter = "[Primary Product] = " & _
                    rd_item.Cells(Me.ColPrimaryProd.Index).Value.ToString & " AND [BatchNumber] = '" & vBatch.ToString & "'"

                    If rd_item.Cells(Me.ColIsKain.Index).Value = True Then
                        Filter &= " AND [Qty At] = " & CDec(vQtyAt).ToString
                    End If
                    QtyIn_New = rd_detail.Cells(Me.Col2QtyIn.Index).Value
                End If
            Else
                If vIsDgvDetail = False Then
                    Filter = "[Primary Product] = " & _
                    rd_item.Cells(Me.ColPrimaryProd.Index).Value.ToString
                Else
                    Filter = "[Primary Product] = " & _
                    rd_item.Cells(Me.ColPrimaryProd.Index).Value.ToString & " AND [BatchNumber] = '" & vBatch.ToString & "'"
                    If rd_item.Cells(Me.ColIsKain.Index).Value = True Then
                        Filter &= " AND [Qty At] = " & CDec(vQtyAt).ToString
                    End If

                End If
            End If


            If Is_LocControl <> 0 Then
                Filter &= " AND [Primary Location] = " & vPrimaryLocation.ToString
            Else
                vPrimaryLocation = -1
            End If

            If QtyIn_New <> 0 Then
                IsSimpan = True
            End If

            vTblItemCustom.DefaultView.RowFilter = Filter
            If IsSimpan = True Then
                'jika simpan
                If vTblItemCustom.DefaultView.Count > 0 Then
                    rowItemCustom = vTblItemCustom.DefaultView(0).Row
                    With rowItemCustom
                        If .Is_Kain = True Then
                            If vIsDgvDetail = True Then
                                .Qty_At = rd_detail.Cells(Me.Col2SecondaryQty.Index).Value
                                .Qty_Roll = rd_detail.Cells(Me.Col2QtyIn.Index).Value
                                .Total_Qty_Multi = .Qty_Roll
                                .Total_Qty = Me.ObjCD.DigitGrouping(.Qty_At * .Qty_Roll, .Decimal_Product)
                                .Ratio_Default = rd_item.Cells(Me.ColRatioDef.Index).Value
                                .Ratio_Reference = rd_item.Cells(Me.ColRatioMulti.Index).Value
                            End If
                        Else
                            .Ratio_Default = rd_item.Cells(Me.ColRatioDef.Index).Value
                            .Ratio_Reference = rd_item.Cells(Me.ColRatioMulti.Index).Value
                            If vIsDgvDetail = True Then
                                .Total_Qty = rd_detail.Cells(Me.Col2QtyIn.Index).Value
                                .Total_Qty_Multi = Me.ObjCD.DigitGrouping((.Total_Qty * .Ratio_Reference) / .Ratio_Default, .Decimal_Multi)
                            Else
                                If vUbahDiQtyDef_DgvAtas = True Then
                                    .Total_Qty = rd_item.Cells(Me.ColTotalQty.Index).Value
                                    .Total_Qty_Multi = Me.ObjCD.DigitGrouping((.Total_Qty * .Ratio_Reference) / .Ratio_Default, .Decimal_Multi)
                                Else
                                    .Total_Qty_Multi = rd_item.Cells(Me.ColTotalQtyMulti.Index).Value
                                    .Total_Qty = Me.ObjCD.DigitGrouping((.Total_Qty_Multi * .Ratio_Default) / .Ratio_Reference, .Decimal_Product)
                                End If
                            End If
                            .Qty_At = -1
                            .Qty_Roll = 0
                        End If
                        If vIsDgvDetail = True Then
                            .BatchNumber = rd_detail.Cells(Me.Col2BatchNo.Index).Value
                            If IsDate(rd_detail.Cells(Me.Col2ExpiryDate.Index).Value) = False Then
                                .Expiry_Date = #1/1/1950#
                            Else
                                .Expiry_Date = Format(CDate(rd_detail.Cells(Me.Col2ExpiryDate.Index).Value), "dd MMM yyyy")
                            End If
                        Else
                            .BatchNumber = ""
                            .Expiry_Date = #1/1/1950#
                        End If
                        '.Unit_Price = rd_item.Cells(Me.ColUnitPrice.Index).Value
                        '.Discount = rd_item.Cells(Me.ColDiscount.Index).Value
                        '.Disc_Amount = rd_item.Cells(Me.ColDiscAmount.Index).Value
                        '.Nett_Price = rd_item.Cells(Me.ColNettPrice.Index).Value
                        '.Amount = rd_item.Cells(Me.ColAmount.Index).Value
                        '.At_Amount = rd_item.Cells(Me.ColAtAmount.Index).Value
                        '.At_Price = rd_item.Cells(Me.ColATCost.Index).Value

                        '.Price_by_Multi_Unit = rd_item.Cells(Me.ColPriceByMultiUnit.Index).Value

                    End With


                Else
                    rowItemCustom = vTblItemCustom.Newtbltrans_custom_aop_item1Row
                    With rowItemCustom
                        .Primary = rd_item.Index + 1
                        '.Primary_Transaksi = Me.PrimaryTransaksiAOP
                        .Product_Code = rd_item.Cells(Me.ColProductCode.Index).Value
                        .Supplier_Do_No = ""
                        .Description = rd_item.Cells(Me.ColDescription.Index).Value
                        .PO_No = rd_item.Cells(Me.ColPONO.Index).Value
                        .GRN_No = rd_item.Cells(Me.ColGRNNO.Index).Value
                        .Stock_In_Purchase = rd_item.Cells(Me.ColStockOutNo.Index).Value
                        .Supplier_Do_No = rd_item.Cells(Me.ColSuppDONo.Index).Value
                        .Qty_Master = rd_item.Cells(Me.ColQtyMaster.Index).Value
                        .Qty_Per_Master = rd_item.Cells(Me.ColQtyPerMaster.Index).Value
                        .Qty = rd_item.Cells(Me.ColQty.Index).Value

                        .Total_Qty = rd_item.Cells(Me.ColTotalQty.Index).Value

                        .Unit_Code = rd_item.Cells(Me.ColUnit.Index).Value
                        .Secondary_Qty = -1

                        .Currency_Code = rd_item.Cells(Me.ColCurrency.Index).Value
                        .Unit_Price = rd_item.Cells(Me.ColUnitPrice.Index).Value
                        .Discount = rd_item.Cells(Me.ColDiscount.Index).Value
                        .Disc_Amount = rd_item.Cells(Me.ColDiscAmount.Index).Value
                        .Nett_Price = rd_item.Cells(Me.ColNettPrice.Index).Value
                        .Amount = rd_item.Cells(Me.ColAmount.Index).Value
                        .At_Price = rd_item.Cells(Me.ColATCost.Index).Value
                        .Remarks = rd_item.Cells(Me.ColRemarksItem.Index).Value
                        .Primary_Product = rd_item.Cells(Me.ColPrimaryProd.Index).Value
                        .Primary_PO = rd_item.Cells(Me.ColPrimaryPO.Index).Value
                        .Primary_GRN = rd_item.Cells(Me.ColPrimaryGRN.Index).Value
                        .Primary_Master_Unit = -1
                        .Primary_Units = rd_item.Cells(Me.ColPrimaryUnit.Index).Value
                        .Primary_Secondary_Units = rd_item.Cells(Me.ColPrimarySecUnit.Index).Value
                        .Primary_Currency = rd_item.Cells(Me.ColPrimaryCurrency.Index).Value
                        If .Storage_Area_Code = "" Then
                            .Storage_Area_Code = "-"
                        End If

                        .Storage_Area_Code = Me.TxtLocation.Text
                        .Primary_Location = Me.txtPrimaryLocation.Text
                        .Refered_To_PO = rd_item.Cells(Me.ColReferedToPO.Index).Value
                        .At_Amount = rd_item.Cells(Me.ColAtAmount.Index).Value
                        .Disc_From_Perct = rd_item.Cells(Me.ColDiscFromPerct.Index).Value
                        .Total_Qty_Multi = rd_item.Cells(Me.ColTotalQtyMulti.Index).Value
                        .Unit_Code_Multi = rd_item.Cells(Me.ColUnitsCodeMulti.Index).Value
                        .Ratio_Reference = rd_item.Cells(Me.ColRatioMulti.Index).Value
                        .Ratio_Default = rd_item.Cells(Me.ColRatioDef.Index).Value
                        .Primary_Units_Multi = rd_item.Cells(Me.ColPrimaryUnitMulti.Index).Value
                        .Adjustment = rd_item.Cells(Me.ColItemAdjustment.Index).Value
                        .Reason = rd_item.Cells(Me.ColItemAdjustmentReason.Index).Value
                        .User_Code = rd_item.Cells(Me.ColItemAdjustmentUser.Index).Value
                        .Adjustment_Timestamp = #1/1/1950#
                        .Price_by_Multi_Unit = rd_item.Cells(Me.ColPriceByMultiUnit.Index).Value
                        .Unit_Code_Multi = rd_item.Cells(Me.ColUnitsCodeMulti.Index).Value
                        .Is_Kain = rd_item.Cells(Me.ColIsKain.Index).Value
                        .Decimal_Currency = Me.ObjCurrency.GetDefaultPrecission(rowItemCustom.Primary_Currency)
                        'item detail
                        If vIsDgvDetail = True Then
                            .BatchNumber = rd_detail.Cells(Me.Col2BatchNo.Index).Value
                            If IsDate(rd_detail.Cells(Me.Col2ExpiryDate.Index).Value) = False Then
                                .Expiry_Date = #1/1/1950#
                            Else
                                .Expiry_Date = Format(CDate(rd_detail.Cells(Me.Col2ExpiryDate.Index).Value), "dd MMM yyyy")
                            End If
                        Else
                            .BatchNumber = ""
                            .Expiry_Date = #1/1/1950#
                        End If

                        If .Is_Kain = True Then
                            .Qty_At = rd_detail.Cells(Me.Col2SecondaryQty.Index).Value
                            .Qty_Roll = rd_detail.Cells(Me.Col2QtyIn.Index).Value
                            .Total_Qty_Multi = .Qty_Roll
                            .Total_Qty = .Qty_At * .Qty_Roll
                            .Ratio_Default = rd_item.Cells(Me.ColRatioDef.Index).Value
                            .Ratio_Reference = rd_item.Cells(Me.ColRatioMulti.Index).Value
                        Else
                            .Ratio_Default = rd_item.Cells(Me.ColRatioDef.Index).Value
                            .Ratio_Reference = rd_item.Cells(Me.ColRatioMulti.Index).Value
                            If vIsDgvDetail = True Then
                                .Total_Qty = rd_detail.Cells(Me.Col2QtyIn.Index).Value
                                .Total_Qty_Multi = (.Total_Qty * .Ratio_Reference) / .Ratio_Default
                            Else
                                If vUbahDiQtyDef_DgvAtas = True Then
                                    .Total_Qty = rd_item.Cells(Me.ColTotalQty.Index).Value
                                    .Total_Qty_Multi = (.Total_Qty * .Ratio_Reference) / .Ratio_Default
                                Else
                                    .Total_Qty_Multi = rd_item.Cells(Me.ColTotalQtyMulti.Index).Value
                                    .Total_Qty = (.Total_Qty_Multi * .Ratio_Default) / .Ratio_Reference
                                End If
                            End If
                            .Qty_At = -1
                            .Qty_Roll = 0
                        End If

                        .Secondary_Qty = -1
                        .Remarks = ""
                        .Decimal_Multi = rd_item.Cells(Me.ColDecMulti.Index).Value
                        .Decimal_Product = rd_item.Cells(Me.ColDec.Index).Value
                        .Primary_Units_Multi = rd_item.Cells(Me.ColPrimaryUnitMulti.Index).Value
                    End With
                    vTblItemCustom.Rows.Add(rowItemCustom)
                End If
            Else
                'jika hapus (edit qty ke 0)
                If vTblItemCustom.DefaultView.Count > 0 Then
                    For x As Integer = 0 To vTblItemCustom.DefaultView.Count - 1
                        rowItemCustom = vTblItemCustom.DefaultView(0).Row
                        vTblItemCustom.Rows.Remove(rowItemCustom)
                    Next
                End If
            End If

            vTblItemCustom.DefaultView.RowFilter = "[Primary Location] <> -1"
            If vTblItemCustom.DefaultView.Count > 0 Then
                Me.TxtLocation.Enabled = False
                Me.btnFindLocation.Enabled = False
            Else
                Me.TxtLocation.Enabled = True
                Me.btnFindLocation.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTblItemCustom.DefaultView.RowFilter = Nothing
        End Try
    End Sub

    'samakan harga per product karena terkait dengan form advance
    Private Sub UpdatePriceByProduct(ByVal vRowIdxAcuan As Integer)
        Dim rd As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim PrimaryProduct As Double = -1

        Try
            PrimaryProduct = Me.dgvItem.Item(Me.ColPrimaryProd.Index, vRowIdxAcuan).Value
            Me.TblItemCustom.DefaultView.RowFilter = "[Primary Product] = " & PrimaryProduct
            For i As Integer = 0 To Me.TblItemCustom.DefaultView.Count - 1
                rd = Me.TblItemCustom.DefaultView(i).Row
                With rd
                    If .Primary_Product = PrimaryProduct Then
                        .price_by_multi_unit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, vRowIdxAcuan).Value
                        .Disc_From_Perct = Me.dgvItem.Item(Me.ColDiscFromPerct.Index, vRowIdxAcuan).Value
                        .Unit_Price = Me.dgvItem.Item(Me.ColUnitPrice.Index, vRowIdxAcuan).Value
                        .Discount = Me.dgvItem.Item(Me.ColDiscount.Index, vRowIdxAcuan).Value
                        .Disc_Amount = Me.dgvItem.Item(Me.ColDiscAmount.Index, vRowIdxAcuan).Value
                        .Nett_Price = Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIdxAcuan).Value
                        .Amount = Me.dgvItem.Item(Me.ColAmount.Index, vRowIdxAcuan).Value


                        Hitung_AmountInDB(rd, -1)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub Hitung_AmountInDB(ByRef row As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row, ByVal vSecQty As String)
        Dim DefaultPrecisionRow As Integer
        Dim DiscPerct As Integer = 0

        Try
            With row
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(.Primary_Currency)
                'hitung nett price
                DiscPerct = .Disc_From_Perct
                If DiscPerct = 1 Then
                    row.Disc_Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(((CDec(.Unit_Price) * CDec(.Discount))) / 100, DefaultPrecisionRow))
                Else
                    row.Discount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(((CDec(.Disc_Amount) * 100) / CDec(.Unit_Price)), 2))
                End If
                .Nett_Price = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(.Unit_Price) - CDec(.Disc_Amount), DefaultPrecisionRow))

                If IsNumeric(vSecQty) = False Then
                    If .price_by_multi_unit = 1 Then
                        .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total_Qty_Multi * .Nett_Price, DefaultPrecisionRow))
                    Else
                        .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total_Qty * .Nett_Price, DefaultPrecisionRow))
                    End If
                Else
                    If vSecQty <> -1 Then
                        If .price_by_multi_unit = 1 Then
                            .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(.Total_Qty_Multi) * .Nett_Price * CDbl(vSecQty), DefaultPrecisionRow))
                        Else
                            .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(.Total_Qty) * .Nett_Price * CDbl(vSecQty), DefaultPrecisionRow))
                        End If
                    Else
                        If .price_by_multi_unit = 1 Then
                            .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(.Total_Qty_Multi) * .Nett_Price, DefaultPrecisionRow))
                        Else
                            .Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(.Total_Qty) * .Nett_Price, DefaultPrecisionRow))
                        End If
                    End If
                End If
                'End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Dim IdxSelectedLast As Integer = -1

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Dim rowIdx As Integer = -1

        Try
            If AktifkanSelectedDgvItem = True Then
                If Me.dgvItem.SelectedCells.Count <= 0 Then
                    Exit Sub
                End If

                If IdxSelectedLast <> -1 Then
                    If IdxSelectedLast = Me.dgvItem.SelectedCells(0).RowIndex Then
                        Exit Sub
                    End If
                End If

                IdxSelectedLast = Me.dgvItem.SelectedCells(0).RowIndex

                rowIdx = Me.dgvItem.SelectedCells(0).RowIndex
                Me.SetColumnDgvDetail(Me.dgvItem(Me.ColIsKain.Index, rowIdx).Value, Me.dgvItem(Me.colIsBatch.Index, rowIdx).Value, Me.dgvItem(Me.colIsExpDate.Index, rowIdx).Value)
                Me.ShowDgvDetail(Me.dgvItem(Me.ColPrimaryPO.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value, Me.dgvItem(Me.ColPrimaryProd.Index, rowIdx).Value, Me.txtPrimaryLocation.Text, Me.TblItemCustom, Me.dgvItem(Me.ColIsLocation.Index, rowIdx).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ShowDgvDetail(ByVal vPrimaryTrans As Double, ByVal vPrimaryProd As Double, ByVal vPrimaryLoc As Double, _
  ByVal vTblItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, ByVal Is_LocControl As Integer)
        Dim rowItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim rd As DataGridViewRow
        Dim Filter As String = ""

        Try
            Me.dgvItemDetail.Rows.Clear()
            Me.dgvItemDetail.AllowUserToAddRows = False
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Filter = "[Primary Product] = " & vPrimaryProd.ToString
            If Is_LocControl <> 0 Then
                Filter &= " AND [Primary Location] = " & vPrimaryLoc.ToString
            End If

            vTblItemCustom.DefaultView.RowFilter = Filter
            vTblItemCustom.DefaultView.Sort = "[Expiry Date], [BatchNumber], [Qty At]"
            For Each dv As DataRowView In vTblItemCustom.DefaultView
                rowItemCustom = dv.Row
                With rowItemCustom
                    rd = Me.dgvItemDetail.Rows(Me.dgvItemDetail.Rows.Add)
                    rd.Cells(Me.Col2SecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_At, .Decimal_Product)
                    rd.Cells(Me.Col2SecondaryUnit.Index).Value = .Unit_Code
                    If .Is_Kain = True Then
                        rd.Cells(Me.Col2QtyIn.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Roll, .Decimal_Multi)
                        rd.Cells(Me.Col2Unit.Index).Value = .Unit_Code_Multi
                        rd.Cells(Me.Col2Total.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product)
                    Else
                        rd.Cells(Me.Col2QtyIn.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product)
                        rd.Cells(Me.Col2Unit.Index).Value = .Unit_Code
                        rd.Cells(Me.Col2Total.Index).Value = Me.ObjCD.DigitGrouping(0, .Decimal_Product)
                    End If
                    rd.Cells(Me.Col2TotalUnit.Index).Value = .Unit_Code
                    rd.Cells(Me.Col2BatchNo.Index).Value = .BatchNumber
                    rd.Cells(Me.Col2ExpiryDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                End With
            Next
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemDetail.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetColumnDgvDetail(ByVal vIsKain As Boolean, ByVal vBatchControl As Integer, ByVal vExpDateControl As Integer)
        'Me.dgvItem.Rows.Add("1", "000001/PO/01/2019", "KT-001", "Kain Trading", "10", "ROLL", "0", "ROLL", "1 ROLL = 50 Meter", "0", "Meter")
        'Me.dgvItem.Rows.Add("2", "000002/PO/01/2019", "BR-002", "Bearing B002-No Control", "20", "CTN", "0", "CTN", "1 CTN = 10 Pcs", "0", "Pcs")
        'Me.dgvItem.Rows.Add("3", "000003/PO/01/2019", "PT-001", "Cat Merah 003-Batch Control", "1", "Drum", "0", "Drum", "1 Drum = 1 Drum", "0", "Drum")
        'Me.dgvItem.Rows.Add("4", "000003/PO/01/2019", "DR-001", "Obat Pemutih 004-Batch & Expiry Control", "10", "Kaleng", "0", "Kaleng", "1 Kaleng = 0.5 Kg", "0", "Kg")

        Try
            If vIsKain = True Then
                Me.SplitContainer1.Panel2Collapsed = False
                Me.Col2SecondaryQty.Visible = True
                Me.Col2SecondaryUnit.Visible = True
                Me.Col2BatchNo.Visible = False
                Me.Col2ExpiryDate.Visible = False
                Me.Col2Total.Visible = True
                Me.Col2TotalUnit.Visible = True
                Me.Label2.Text = "Secondary Unit Data"
            Else
                If vBatchControl = 0 And vExpDateControl = 0 Then
                    Me.SplitContainer1.Panel2Collapsed = True
                    Me.Col2SecondaryQty.Visible = False
                    Me.Col2SecondaryUnit.Visible = False
                    Me.Col2BatchNo.Visible = False
                    Me.Col2ExpiryDate.Visible = False
                    Me.Col2Total.Visible = False
                    Me.Col2TotalUnit.Visible = False
                ElseIf vBatchControl <> 0 And vExpDateControl = 0 Then
                    Me.SplitContainer1.Panel2Collapsed = False
                    Me.Col2SecondaryQty.Visible = False
                    Me.Col2SecondaryUnit.Visible = False
                    Me.Col2BatchNo.Visible = True
                    Me.Col2ExpiryDate.Visible = False
                    Me.Col2Total.Visible = False
                    Me.Col2TotalUnit.Visible = False
                    Me.Label2.Text = "Batch Data"
                ElseIf vBatchControl <> 0 And vExpDateControl <> 0 Then
                    Me.SplitContainer1.Panel2Collapsed = False
                    Me.Col2SecondaryQty.Visible = False
                    Me.Col2SecondaryUnit.Visible = False
                    Me.Col2BatchNo.Visible = True
                    Me.Col2ExpiryDate.Visible = True
                    Me.Col2Total.Visible = False
                    Me.Col2TotalUnit.Visible = False
                    Me.Label2.Text = "Batch and Expiry Data"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItemDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemDetail.CellBeginEdit
        Dim objPO As New ClsTransaksiPO

        Try
            If e.ColumnIndex <> Me.Col2Unit.Index And e.ColumnIndex <> Me.Col2SecondaryUnit.Index And e.ColumnIndex <> Me.Col2Total.Index And e.ColumnIndex <> Me.Col2TotalUnit.Index Then
                If Me.dgvItem.Rows.Count = 0 Then
                    e.Cancel = True
                    Exit Sub
                End If
                If String.IsNullOrEmpty(Me.TxtLocation.Text) = True And Me.dgvItem(Me.ColIsLocation.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value <> 0 Then
                    MsgBox("Please enter location first.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocation.Focus()
                    e.Cancel = True
                    Exit Sub
                End If

                If e.ColumnIndex = Me.Col2QtyIn.Index Then
                    If Me.dgvItem.Item(Me.ColIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value = False Then
                        If Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value = "" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    Else
                        If Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value = 0 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If

                Me.BufferDetail = Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemDetail.CellEndEdit
        Dim index As Integer = -1
        Dim decDef As Integer = 0
        Dim decMulti As Integer = 0
        Dim ParamQtyAt As Decimal = 0
        Dim ParamBatch As String = ""
        Dim ParamExpDate As Date = #1/1/1950#
        Dim IsKain As Boolean = False
        Dim JmlQtyUpdate As Decimal = 0
        Dim RowIdxDgvAtas As Integer = -1

        Try
            If Me.dgvItem.SelectedCells.Count <= 0 Then
                MsgBox("Please select data above first.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                Exit Sub
            End If

            If IsNothing(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If IsNumeric(Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value) = True And IsNumeric(Me.BufferDetail) = True Then
                    If CDec(Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value) = CDec(Me.BufferDetail) Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                Else
                    If Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                End If

                RowIdxDgvAtas = Me.dgvItem.SelectedCells(0).RowIndex
                decDef = Me.dgvItem(Me.ColDec.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                decMulti = Me.dgvItem(Me.ColDecMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                ParamBatch = Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value
                IsKain = Me.dgvItem(Me.ColIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) = True Then
                    ParamQtyAt = Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value
                End If
                If IsDate(Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value) = True Then
                    ParamExpDate = CDate(Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value)
                Else
                    ParamExpDate = #1/1/1950#
                End If

                If IsKain = True Then
                    If Me.Col2SecondaryQty.Visible = True Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) = True And _
                        Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value) = True Then
                            JmlQtyUpdate = Me.ObjCD.DigitGrouping(CDec(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) * CDec(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value), decDef)
                        End If
                    End If
                Else
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value) = True Then
                        JmlQtyUpdate = Me.ObjCD.DigitGrouping(CDec(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value), decDef)
                    End If
                End If

                If e.ColumnIndex = Me.Col2SecondaryQty.Index Then
                    'untuk nilai qty dan secondary qty, pastikan nilainya positif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for qty in multi data.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If

                    index = Me.GetIndex(True, Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value, False, "", e.RowIndex)
                    If index >= 0 Then
                        Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If

                    ParamQtyAt = Me.BufferDetail

                    'cek apakah edit sah
                    If Me.ApakahEditSah(ParamBatch, ParamExpDate, 0, ParamQtyAt, EnumEditCol.Edit_QtyRoll, RowIdxDgvAtas, False) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    'tutup

                    Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value, decDef)
                ElseIf e.ColumnIndex = Me.Col2QtyIn.Index Then
                    'untuk nilai qty dan secondary qty, pastikan nilainya positif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for qty in default data.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If

                    Dim DecDumm As Integer = decDef
                    Dim EditColBy As EnumEditCol = EnumEditCol.Edit_QtyDef
                    If Me.dgvItem(Me.ColIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value = True Then
                        DecDumm = decMulti
                        EditColBy = EnumEditCol.Edit_QtyRoll
                    End If

                    'cek apakah edit sah
                    If Me.ApakahEditSah(ParamBatch, ParamExpDate, JmlQtyUpdate, ParamQtyAt, EditColBy, RowIdxDgvAtas, True) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    'tutup

                    Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value, DecDumm)

                ElseIf e.ColumnIndex = Me.Col2BatchNo.Index Then
                    Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("*", "").Replace("'", "").Replace("%", "")
                    Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    index = Me.GetIndex(False, -1, True, Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                    If index >= 0 Then
                        Me.dgvItemDetail(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    ParamBatch = Me.BufferDetail

                    If IsDate(Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value) = True Then
                        Dim rowBawah As DataGridViewRow = Me.dgvItemDetail.Rows(e.RowIndex)
                        If Me.CekBatasanExpDate(rowBawah, Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value, Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value, True) = False Then
                            Exit Sub
                        End If
                    End If

                    'cek apakah edit sah
                    If Me.ApakahEditSah(ParamBatch, ParamExpDate, 0, ParamQtyAt, EnumEditCol.Edit_QtyDef, RowIdxDgvAtas, False) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    'tutup
                ElseIf e.ColumnIndex = Me.Col2ExpiryDate.Index Then
                    If IsDate(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = CDate(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Format(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value, "dd MMM yyyy")

                    If IsDate(Me.BufferDetail) = True Then
                        ParamExpDate = CDate(Me.BufferDetail)
                    End If

                    Dim rowBawah As DataGridViewRow = Me.dgvItemDetail.Rows(e.RowIndex)
                    If Me.CekBatasanExpDate(rowBawah, Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value, Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value, False) = False Then
                        Exit Sub
                    End If

                    'cek apakah edit sah
                    If Me.ApakahEditSah(ParamExpDate, Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value, 0, ParamQtyAt, EnumEditCol.Edit_QtyDef, RowIdxDgvAtas, False) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
                        Exit Sub
                    End If
                    'tutup

                    Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                End If

                If IsKain = True Then
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) = False Or _
                    Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value) = False Then
                        Exit Sub
                    End If

                    Dim TotalQty As Decimal = CDec(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, e.RowIndex).Value) * Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value
                    Me.dgvItemDetail.Item(Me.Col2Total.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(TotalQty, decDef)
                Else
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItemDetail.Item(Me.Col2QtyIn.Index, e.RowIndex).Value) = False Then
                        Exit Sub
                    End If
                    If Me.Col2BatchNo.Visible = True Then
                        If IsNothing(Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value) = True Or String.IsNullOrEmpty(Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, e.RowIndex).Value) = True Then
                            Exit Sub
                        End If
                    End If
                    If Me.Col2ExpiryDate.Visible = True Then
                        If IsDate(Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, e.RowIndex).Value) = False Then
                            Exit Sub
                        End If
                    End If
                End If

                Dim rd_Item As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                Dim rd_Detail As DataGridViewRow = Me.dgvItemDetail.Rows(e.RowIndex)
                Dim angkaDecimal As Int16 = 0
                angkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)

                Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryLocation.Text, ParamQtyAt, ParamBatch, rd_Item, rd_Detail, True, False, rd_Item.Cells(Me.ColIsLocation.Index).Value)
                Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)
                Me.HitungItem(Me.dgvItem.SelectedCells(0).RowIndex, angkaDecimal)
                Me.UpdatePriceByProduct(Me.dgvItem.SelectedCells(0).RowIndex)
                'Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryLocation.Text, ParamQtyAt, ParamBatch, rd_Item, rd_Detail, True, False, rd_Item.Cells(Me.ColIsLocation.Index).Value)
                Me.Hitung_Total_Multi()
            Else
                Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = Me.BufferDetail
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function CekBatasanExpDate(ByVal rowBawah As DataGridViewRow, ByVal vBatch As String, ByVal vExpDate As Date, _
   ByVal vIsChangeByBatch As Boolean) As Boolean
        Dim PrimaryProd As Double = -1
        Dim PrimaryPO As Double = -1
        Dim IsKain As Boolean = False
        Dim KenaBatasan As Boolean = False

        Try
            PrimaryProd = CDbl(Me.dgvItem(Me.colPrimaryProd.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
            PrimaryPO = CDbl(Me.dgvItem(Me.colPrimaryPO.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
            IsKain = CBool(Me.dgvItem(Me.colIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
            If vIsChangeByBatch = False Then
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, rowBawah.Index).Value = Format(vExpDate, "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If vExpDate < Me.TglTransaksi.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, rowBawah.Index).Value = Me.BufferDetail
                        Return False
                    End If
                End If
            End If

            'pastikan tanggal expiry date tidak konflik dengan data yang sudah ada
            Dim ExpiryDateLama As Date = Me.ObjProd.GetExpiryDate(PrimaryProd, vBatch)
            If ExpiryDateLama <> #1/1/1950# Then
                If ExpiryDateLama.Date <> vExpDate Then
                    KenaBatasan = True
                    GoTo ToBatasan
                End If
            End If
            'bandingkan dengan yang sudah ada didalam dgv
            'PENYESUAIAN CODING DENGAN DEFAULTVIEW
            Dim Filter As String = ""
            Dim BatchNo As String = ""
            If String.IsNullOrEmpty(vBatch) = False Then
                BatchNo = vBatch
            End If
            Filter = "[Primary Product] = " & _
            PrimaryProd.ToString & " AND [BatchNumber] = '" & BatchNo.ToString & "'"
            If IsKain = True Then
                Filter &= " AND [Qty At] = " & CDec(Me.dgvItemDetail.Item(Me.Col2SecondaryQty.Index, rowBawah.Index).Value).ToString
            End If
            Me.TblItemCustom.DefaultView.RowFilter = Filter
            If Me.TblItemCustom.DefaultView.Count > 0 Then
                Dim rowItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
                rowItemCustom = Me.TblItemCustom.DefaultView(0).Row
                If IsDate(rowItemCustom.Expiry_Date) = True Then
                    ExpiryDateLama = rowItemCustom.Expiry_Date
                Else
                    ExpiryDateLama = #1/1/1950#
                End If
                KenaBatasan = True
                GoTo ToBatasan
            End If

            'bandingkan dengan yang sudah ada didalam transaksi lain. Detail lihat point no 8.c pada 
            'fungsi apakah data item benar
            '1. stock in (semua stock in)
            Dim JenisTransaksi As ClsTransaksiStockInOut.EnumTipeStockInOut = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            For loopInOut As Integer = 0 To 4 Step 1
                Select Case loopInOut
                    Case 0
                        '   a. stock in purchase
                        JenisTransaksi = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                        Exit Select
                    Case 1
                        '   b. stock in sales return
                        JenisTransaksi = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                        Exit Select
                    Case 2
                        '   c. stock in adjust
                        JenisTransaksi = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                        Exit Select
                    Case 3
                        '   d. stock in back
                        JenisTransaksi = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                        Exit Select
                    Case 4
                        '   e. stock in proses
                        JenisTransaksi = ClsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                        Exit Select
                End Select
                ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(PrimaryProd, vBatch, JenisTransaksi)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> vExpDate Then
                        KenaBatasan = True
                        GoTo ToBatasan
                    End If
                End If
            Next
            '2. transaksi stock bundling untuk item hasil bundling
            ExpiryDateLama = Me.ObjReference.GetExpiryDateAfter_ItemBundAft(PrimaryProd, vBatch)
            If ExpiryDateLama <> #1/1/1950# Then
                If ExpiryDateLama.Date <> vExpDate Then
                    KenaBatasan = True
                    GoTo ToBatasan
                End If
            End If

ToBatasan:
            If KenaBatasan = True Then
                If vIsChangeByBatch = False Then
                    MsgBox("Expiry date for batch no: " + CStr(vBatch) + " is " + Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, rowBawah.Index).Value = Me.BufferDetail
                    Return False
                Else
                    Dim BufferExpiryDate As String = "-"
                    If MsgBox("Expiry date for batch no: " + CStr(vBatch) + " is " + Format(CDate(vExpDate), "dd MMM yyyy") + _
                    ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        'update data expiry date juga
                        Me.dgvItemDetail.Item(Me.Col2ExpiryDate.Index, rowBawah.Index).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                        Return True
                    Else
                        Me.dgvItemDetail.Item(Me.Col2BatchNo.Index, rowBawah.Index).Value = Me.BufferDetail
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetIndex(ByVal vCekQtyAt As Boolean, ByVal vQtyAt As Decimal, ByVal vCekBatch As Boolean, ByVal vBatch As String, ByVal vRowIdx As Integer) As Integer
        Dim retval As Integer = -1
        Dim rd As DataGridViewRow

        Try
            For i As Integer = 0 To Me.dgvItemDetail.Rows.Count - 1
                rd = Me.dgvItemDetail.Rows(i)
                If vCekQtyAt = True Then
                    If vRowIdx <> i Then
                        If CDec(rd.Cells(Me.Col2SecondaryQty.Index).Value) = vQtyAt Then
                            MsgBox("Please enter other value for qty @.", MsgBoxStyle.Information, "Information")
                            retval = i
                        End If
                    End If
                End If
                If vCekBatch = True Then
                    If vRowIdx <> i Then
                        If rd.Cells(Me.Col2BatchNo.Index).Value = vBatch Then
                            MsgBox("Please enter other value for batch.", MsgBoxStyle.Information, "Information")
                            retval = i
                        End If
                    End If
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemDetail_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemDetail.RowsAdded
        Dim rd As DataGridViewRow = Me.dgvItemDetail.Rows(Me.dgvItemDetail.RowCount - 1)

        Try
            If Me.dgvItem.SelectedCells.Count > 0 Then
                rd.Cells(Me.Col2SecondaryUnit.Index).Value = Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                rd.Cells(Me.Col2TotalUnit.Index).Value = Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                If Me.dgvItem.Item(Me.ColIsKain.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value = True Then
                    rd.Cells(Me.Col2Unit.Index).Value = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                Else
                    rd.Cells(Me.Col2Unit.Index).Value = Me.dgvItem.Item(Me.ColUnit.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                End If
            End If

            rd.Cells(Me.Col2ExpiryDate.Index).Value = ""
            If Me.Col2ExpiryDate.Visible = True Then
                rd.Cells(Me.Col2ExpiryDate.Index).Value = Format(Now.Date, "dd MMM yyyy")
            End If
            rd.Cells(Me.Col2SecondaryQty.Index).Value = 0
            rd.Cells(Me.Col2QtyIn.Index).Value = 0
            rd.Cells(Me.Col2Total.Index).Value = 0
            rd.Cells(Me.Col2BatchNo.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemDetail_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemDetail.UserDeletingRow
        'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
        Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
        JmlDGV, JmlSetelahHapus, QtyAtDumm As Decimal
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim rowItem As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
        Dim rowDetail As DataGridViewRow = e.Row
        Dim rowDummy As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim ListProd As New List(Of Double)
        Dim angkaDecimal As Double = 0
       
        Try
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If ListProd.Contains(rd.Cells(Me.ColPrimaryProd.Index).Value) = False Then
                    If rd.Cells(Me.ColIsLocation.Index).Value = 0 Then
                        Me.TblItemCustom.DefaultView.RowFilter = "[Primary Product] = " & CDbl(rd.Cells(Me.ColPrimaryProd.Index).Value).ToString
                        For Each dv As DataRowView In Me.TblItemCustom.DefaultView
                            rowDummy = dv.Row
                            rowDummy.Primary_Location = -1
                            rowDummy.Storage_Area_Code = ""
                        Next
                    End If

                    ListProd.Add(rd.Cells(Me.ColPrimaryProd.Index).Value)
                End If
            Next
            If Me.ObjProd.ApakahGoods(rowItem.Cells(Me.ColPrimaryProd.Index).Value) = True Then
                PrimaryLocation = Me.txtPrimaryLocation.Text
                PrimaryPdO = -1
                PrimaryProduct = rowItem.Cells(Me.ColPrimaryProd.Index).Value
                BatchNo = rowDetail.Cells(Me.Col2BatchNo.Index).Value
                If IsDate(rowDetail.Cells(Me.Col2ExpiryDate.Index).Value) = True Then
                    ExpiryDate = CDate(rowDetail.Cells(Me.Col2ExpiryDate.Index).Value)
                Else
                    ExpiryDate = #1/1/1950#
                End If
                SecondaryQty = -1
                If IsNumeric(rowDetail.Cells(Me.Col2SecondaryQty.Index).Value) = True Then
                    QtyAtDumm = rowDetail.Cells(Me.Col2SecondaryQty.Index).Value
                Else
                    QtyAtDumm = -1
                End If

                'cek apakah edit sah
                If Me.ApakahEditSah(BatchNo, ExpiryDate, 0, QtyAtDumm, EnumEditCol.Edit_QtyRoll, rowItem.Index, False) = False Then
                    e.Cancel = True
                    Exit Sub
                End If
                'tutup

                'JmlGudang = Me.ObjStockInOutItem.GetJml(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, QtyAtDumm)
                'JmlDGV = Me.GetJmlItem_Default(rowItem.Cells(Me.colPrimaryPO.Index).Value, PrimaryProduct, BatchNo, rowItem.Cells(Me.Col2SecondaryQty.Index).Value, Me.txtPrimaryLocation.Text, rowItem.Cells(Me.colIsLocation.Index).Value)
                'JmlObjLama = Me.ObjStockInOutItem.GetJml2(Me.PrimaryTransaksi, PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate, QtyAtDumm)
                'JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                'If JmlSetelahHapus < 0 Then
                '    MsgBox("You cannot delete this data, because this will make your stock quantity minus.", MsgBoxStyle.Information, "Information")
                '    e.Cancel = True
                '    Exit Sub
                'End If

                Dim rd_Item As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                Dim rd_Detail As DataGridViewRow = Me.dgvItemDetail.Rows(e.Row.Index)
                Dim ParamQtyAt As Decimal = rd_Detail.Cells(Me.Col2SecondaryQty.Index).Value
                Dim ParamBatch As String = rd_Detail.Cells(Me.Col2BatchNo.Index).Value
                rd_Detail.Cells(Me.Col2QtyIn.Index).Value = 0
                Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryLocation.Text, ParamQtyAt, ParamBatch, rd_Item, rd_Detail, True, False, rd_Item.Cells(Me.ColIsLocation.Index).Value)
                Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)

            End If
            angkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
            Me.HitungItem(Me.dgvItem.SelectedCells(0).RowIndex, angkaDecimal)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function CekBatasanMinusHapus(ByVal PrimaryLocation As Double, _
       ByVal PrimaryPO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
       ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal QtyAt As Decimal, _
       Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Double
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
            JmlGudang = Me.ObjStockInOutItem.GetJml(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, QtyAt)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objAopItem.GetJml(Me.PrimaryTransaksiAOP, PrimaryProduct, PrimaryPO, QtyAt)
            Me.ObjProd.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = 0
            JmlSetelahEdit = JmlGudang - JmlObjLama + JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
                    Me.ObjProd.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'sampe sini 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItemDetail.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim RowdgvData As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
        Dim RowdgvItem As DataGridViewRow = e.Row
        Dim RowDummy As DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim jmlDGV As Decimal = 0
        Dim primaryLocation As Double = -1
        Dim Filter As String = ""
        Try
            Dim ListProd As New List(Of Double)
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If ListProd.Contains(rd.Cells(Me.ColPrimaryProd.Index).Value) = False Then
                    If rd.Cells(Me.ColIsLocation.Index).Value = 0 Then
                        Me.TblItemCustom.DefaultView.RowFilter = "[Primary Product] = " & CDbl(rd.Cells(Me.ColPrimaryProd.Index).Value).ToString
                        For Each dv As DataRowView In Me.TblItemCustom.DefaultView
                            RowDummy = dv.Row
                            RowDummy.Primary_Location = -1
                            RowDummy.Storage_Area_Code = ""
                        Next
                    End If

                    ListProd.Add(rd.Cells(Me.ColPrimaryProd.Index).Value)
                End If
            Next
            Dim rd_Item As DataGridViewRow = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
            Dim rd_Detail As DataGridViewRow
            If Me.Baru = False Then
                Filter = "[Primary Product]=" & RowdgvData.Cells(Me.ColPrimaryProd.Index).Value
                Me.TblItemCustom.DefaultView.RowFilter = Nothing
                Me.TblItemCustom.DefaultView.RowFilter = Filter
                If TblItemCustom.DefaultView.Count > 0 Then
                    For i As Integer = 0 To Me.TblItemCustom.DefaultView.Count - 1
                        RowDummy = Me.TblItemCustom.DefaultView(i).Row
                        primaryLocation = RowDummy.Primary_Location
                        jmlDGV = RowDummy.Total_Qty
                        If Me.ObjProd.ApakahGoods(RowdgvData.Cells(Me.ColPrimaryProd.Index).Value) = True Then
                            If Me.CekBatasanMinusHapus(primaryLocation, RowDummy.Primary_PO, RowDummy.Primary_Product, RowDummy.BatchNumber, RowDummy.Expiry_Date, _
                                                       -1, jmlDGV, RowDummy.Qty_At, False) = True Then
                                Me.UpdateTblInOutItem(Me.TblItemCustom, RowDummy.Primary_PO, primaryLocation, RowDummy.Qty_At, "", rd_Item, rd_Detail, True, False, rd_Item.Cells(Me.ColIsLocation.Index).Value, True)
                                Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)
                            Else
                                e.Cancel = True
                                Exit For
                            End If
                        Else
                            Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, primaryLocation, -1, "", rd_Item, rd_Detail, True, False, rd_Item.Cells(Me.ColIsLocation.Index).Value, True)
                            Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)
                        End If

                    Next
                End If
            Else
                If Me.dgvItem.Item(Me.ColIsLocation.Index, e.Row.Index).Value = True Then
                    primaryLocation = Me.txtPrimaryLocation.Text
                End If
                Me.UpdateTblInOutItem(Me.TblItemCustom, rd_Item.Cells(Me.ColPrimaryPO.Index).Value, primaryLocation, -1, "", rd_Item, rd_Detail, False, False, rd_Item.Cells(Me.ColIsLocation.Index).Value, True)
                Me.UpdateDgvItemByDgvDetail(rd_Item, Me.TblItemCustom)
            End If
            IdxSelectedLast = -1
            'Me.dgvItemDetail.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        If Me.TxtRemarkItem.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = BackColorAll
            Else
                TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProd.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGboItem()
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGboItem()
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtUnitPrice.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If

            setLocation_Product()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

        End Try
    End Sub
    Private Sub PrepareMultiUnit(ByVal vPrimaryProduct As Double)
        Dim dtaMultiUnit As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnitMulti.DataSource = Nothing
            If Me.ObjPO.IsMultiUnitProduct(vPrimaryProduct) = False Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProd.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.TxtPrimaryUnitsDefault.Text = Me.TxtPrimaryUnitsMulti.Text
                Me.LblUnitsMulti.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text)
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = True
                Me.CboUnitMulti.Visible = False
                Me.CboUnitMulti.TabStop = False
                Me.LblUnitPrice.Text = "Unit Price/" & LblUnitsMulti.Text
                'Me.SetQtyDefault()

                If Me.TxtPrimaryUnitsDefault.Text = "" Or Me.TxtPrimaryUnitsDefault.Text = "0" Then
                    Me.TxtPrimaryUnitsDefault.Text = "-1"
                End If
                If Me.TxtPrimaryUnitsMulti.Text = "" Or Me.TxtPrimaryUnitsMulti.Text = "0" Then
                    Me.TxtPrimaryUnitsMulti.Text = "-1"
                End If
                If Me.LblUnitsMulti.Text = "" Then
                    Me.LblUnitsMulti.Text = "Units"
                End If
                If Me.LblUnitsDefault.Text = "" Then
                    Me.LblUnitsDefault.Text = "Units"
                End If
            Else
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                dtaMultiUnit.Connection.ConnectionString = Core_Module.GetConstring
                dtaMultiUnit.FillBy_PrimaryProduct(Me.tblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.ObjPO.IsiCboUnit(vPrimaryProduct, Me.CboUnitMulti, Me.tblCboUnit, Me.tblMultiUnit, Me.TxtPrimaryUnitsMulti)
                Me.CboUnitMulti.SelectedIndex = -1
                Me.CboUnitMulti.SelectedIndex = 0
            End If

            setLocation_Product()
            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub Hitung_Amount2()

        Try
            Dim Decimal_Def As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), Decimal_Def)

            If Me.TxtDiscFromPerct.Text = "" Then
                Me.TxtDiscFromPerct.Text = "1"
            End If
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def)
            Else
                'hitung disc %
                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                Else
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount1.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
                End If
            End If

            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount1.Text), Decimal_Def)
            'Obs
            'Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.TxtUnitPrice.Text) * _
            '(100 - CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def))
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text) * _
                CDec(Me.TxtQtyMulti.Text), Decimal_Def)
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text) * _
                0, Decimal_Def)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetCboPriceByMultiUnit()
        Dim UnitCode1 As String = ""
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal

        Try
            Me.cboPriceByMultiUnit.Items.Clear()

            If Me.CboUnitMulti.Items.Count = 0 Then
                'tanpa multi unit
                Me.LblUnitPrice.Text = "Unit Price"
                'Me.LblUnitPrice.Text = "Unit Price / " + Me.LblUnitsDefault.Text
                Me.CboPriceByMultiUnit.Enabled = False
                'Me.CboPriceByMultiUnit.Visible = False
                Me.CboPriceByMultiUnit.TabStop = False
                'Me.Label35.Visible = False
                'Me.Label7.Visible = False
                Me.CboPriceByMultiUnit.Items.Add(Me.LblUnitsDefault.Text)
                Me.txtPriceByMultiUnit.Text = "1"
                Me.CboPriceByMultiUnit.SelectedIndex = 0



                Me.ObjPO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.PrimaryCurrency, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc, Me.txtPriceByMultiUnit.Text)

                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                Me.TxtDiscFromPerct.Text = DiscFromPerc
            Else
                If Me.CboUnitMulti.Items.Count = 1 Then
                    Me.LblUnitPrice.Text = "Unit Price"
                    'Me.LblUnitPrice.Text = "Unit Price / " + Me.LblUnitsDefault.Text
                    Me.CboPriceByMultiUnit.Enabled = False
                    'Me.CboPriceByMultiUnit.Visible = False
                    Me.CboPriceByMultiUnit.TabStop = False
                    'Me.Label35.Visible = False
                    'Me.Label7.Visible = False
                    Me.CboPriceByMultiUnit.Items.Add(Me.LblUnitsDefault.Text)
                    Me.txtPriceByMultiUnit.Text = "1"
                    Me.CboPriceByMultiUnit.SelectedIndex = 0


                Else
                    Me.LblUnitPrice.Text = "Unit Price"
                    UnitCode1 = Me.ObjProd.GetUnitCode(Me.TxtPrimaryCodeNo.Text)


                    Me.CboPriceByMultiUnit.Enabled = True
                    'Me.CboPriceByMultiUnit.Visible = False
                    Me.CboPriceByMultiUnit.TabStop = True
                    'Me.Label35.Visible = True
                    'Me.Label7.Visible = True
                    Me.CboPriceByMultiUnit.Items.Add(UnitCode1)
                    Me.txtPriceByMultiUnit.Text = "1"
                    If Me.CboUnitMulti.Text <> UnitCode1 Then
                        Me.CboPriceByMultiUnit.Items.Add(Me.CboUnitMulti.Text)
                        Me.CboPriceByMultiUnit.SelectedIndex = Me.ObjProd.GetPriceByMultiUnit(Me.TxtPrimaryCodeNo.Text)
                    Else
                        Me.CboPriceByMultiUnit.SelectedIndex = 0
                    End If

                    'Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
                End If
            End If
            If Me.TxtPrimaryCodeNo.Text = "" Or Me.TxtPrimaryCodeNo.Text = "-1" Then
                Me.CboPriceByMultiUnit.Items.Clear()
            End If
            Me.Hitung_Amount2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetGboItem()
        Try

            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.LblUnitsMulti.Text = Me.ObjProd.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.LblUnitsMulti.Text = "Units"
            End If
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub KosongkanItem()
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.RowCount = 0 Then
                If IsNumeric(Me.PrimaryCurrency) = True Then
                    Me.TxtPrimaryCurrency1.Text = Me.PrimaryCurrency
                Else
                    Me.TxtPrimaryCurrency1.Text = "1"
                End If
            Else
                row = Me.dgvItem.Rows(0)
                Me.TxtPrimaryCurrency1.Text = row.Cells(Me.ColPrimaryCurrency.Index).Value
            End If
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

            Me.TxtUnitPrice.Text = "0"
            Me.LblUnitPrice.Text = "Unit Price"
            Me.TxtDiscount.Text = "0"
            Me.txtNettPrice.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtAmount.Text = "0"
            Me.TxtRemarkItem.Text = ""
            Me.TxtCodeNo.Text = ""
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtDesc.Text = ""

            Me.TxtDiscFromPerct.Text = "1"
            Me.TxtDiscAmount1.Text = 0
        
            Me.Label7.Text = ""
            '.CboUnitMulti.Items.Clear()
            Me.SetGboItem()


            Me.CboPriceByMultiUnit.Enabled = False
            Me.CboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.CboPriceByMultiUnit.Items.Clear()
            'Me.CboUnitMulti.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProd.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGboItem()

                    If My.Computer.Keyboard.ShiftKeyDown = False Then
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtQtyMulti.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    Else
                        Me.btnInput.Focus()
                    End If
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGboItem()
                        If My.Computer.Keyboard.ShiftKeyDown = False Then
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtQtyMulti.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        Else
                          Me.btnInput.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.KosongkanItem()

                If My.Computer.Keyboard.ShiftKeyDown = False Then
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtQtyMulti.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                Else
                     Me.btnInput.Focus()
                End If
            End If
            setLocation_Product()
            SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)
        End Try
    End Sub
    Private Sub setLocation_Product()

        Try
            If Me.CboUnitMulti.Visible = True Then
                'Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.CboUnitMulti.Width - 5
                'Me.CboUnitMulti.Left = Me.TxtDiscount.Right - Me.CboUnitMulti.Width

            Else
                'Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.LblUnitsMulti.Width - 5
                'Me.LblUnitsMulti.Left = Me.TxtDiscount.Right - Me.LblUnitsMulti.Width

            End If
            Me.Label7.Text = Me.LblUnitsMulti.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrencyCode1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.GotFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = BackColorAll
        Else
            TxtCurrencyCode1.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtCurrencyCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrencyCode1.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
                Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
    Private Sub SetCalculateBy()
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                'reset data item exchange rate
                Me.dgvExchageRates.Rows.Clear()
                If Me.CalculateByIndex = 0 Then
                    'hitung berdasarkan sales invoice
                    If Me.dgvItem.RowCount = 0 Then
                        'kalo ga ada anak, maka exc
                        Me.dgvExchageRates.Rows.Add("0", Me.NoAOP, Me.CurrencyCode, _
                        1, "", Me.PrimaryTransaksiAOP, Me.PrimaryCurrency)
                    Else
                        Dim rowDGV As DataGridViewRow
                        rowDGV = Me.dgvItem.Rows(0)
                        Me.dgvExchageRates.Rows.Add("0", Me.NoAOP, rowDGV.Cells(Me.ColCurrency.Index).Value, _
                        1, "", Me.PrimaryTransaksiAOP, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    End If
                ElseIf Me.CalculateByIndex = 1 Then
                    'hitung berdasarkan delivery order
                    'hitung berdasarkan sales order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryGRN.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColGRNNO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimaryGRN.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                ElseIf Me.CalculateByIndex = 2 Then
                    'hitung berdasarkan sales order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryPO.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(5).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimaryPO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                End If
                'tampilan exchange rate
                'Me.SetDGVExchangeRate()
                'hitung ulang
                Me.Hitung_Total_Multi()
                'set number
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrencyCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.TxtCurrencyCode1.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)

                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        End If
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

                        If My.Computer.Keyboard.ShiftKeyDown = True Then
                            If Me.TxtLocation.Visible = True Then
                                Me.TxtLocation.Focus()
                            End If
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency1.Text = "-1"

                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    If Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    End If
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrencyCode1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.LostFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.Buffer = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = BackColorAll
        Else
            TxtUnitPrice.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e, 9999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtDiscFromPerct.Text = "0" Then
                    If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                        Me.TxtUnitPrice.Text = Me.Buffer
                        Exit Sub
                    End If
                End If
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                Me.Hitung_Amount2()

                If My.Computer.Keyboard.ShiftKeyDown = False Then
                    If Me.TxtDiscount.Enabled = True Then
                        Me.TxtDiscount.Focus()
                    End If
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            If Me.TxtDiscount.Enabled = True Then
                If CDec(Me.TxtDiscount.Text) = 0 Then
                    Me.TxtDiscount.Text = 0
                End If
                Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = BackColorAll
        Else
            TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscFromPerct.Text = "1"
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)

                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Enter
        Try
            Me.buffer2 = Me.TxtDiscAmount1.Text
            Me.TxtDiscFromPerct.Text = "0"
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.GotFocus
        If Me.TxtDiscAmount1.ReadOnly = False And Me.TxtDiscAmount1.Enabled = True Then
            TxtDiscAmount1.BackColor = BackColorAll
        Else
            TxtDiscAmount1.BackColor = Me.txtPrimaryLocation.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount1, sender, e, Me.TxtUnitPrice.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount1, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount1.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtDiscAmount1.Text = Me.buffer2
                    Exit Sub
                End If
                If CDec(Me.TxtDiscAmount1.Text) <> CDec(Me.buffer2) Then
                    'If IsNumeric(Me.TxtUnitPrice.Text) AndAlso CDec(Me.TxtUnitPrice.Text) <> 0 Then
                    '    Me.TxtDiscount.Text = CDec(CDec(Me.TxtDiscAmount.Text) / CDec(Me.TxtUnitPrice.Text)) * 100
                    'Else
                    '    Me.TxtDiscount.Text = 0
                    'End If
                    Me.TxtDiscFromPerct.Text = "0"
                    Me.Hitung_Amount2()
                Else
                    Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount1.Text, 2)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.LostFocus
        If Me.TxtDiscAmount1.ReadOnly = False And Me.TxtDiscAmount1.Enabled = True Then
            TxtDiscAmount1.BackColor = Me.TxtPrimaryCurrency1.BackColor
        End If
    End Sub

    Private Sub btnInput_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Dim PrimaryLocationDummy As Double = -1
        Dim obj As New Core_Module.clsMasterStorageArea
        Dim QtyAt_Dummy As String = ""
        Dim ApakahKainTrading As Boolean = False
        Dim DecDef As Integer = 0
        Dim DecMulti As Integer = 0
        Dim DecRoll As Integer = 0
        Dim primaryRoll As Decimal = 22
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
           
            If Me.ObjProdCust.ApakahKainTrading(Me.TxtPrimaryCodeNo.Text) = True Then
                ApakahKainTrading = True
                Dim UnitDef As String = Me.ObjProd.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.CboPriceByMultiUnit.Text <> UnitDef Then
                    MsgBox("Please specify price by " & UnitDef, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.ApakahDataItemBenar = True Then
                If Me.txtPrimaryLocation.Text.Trim <> -1 Then
                    PrimaryLocationDummy = Me.txtPrimaryLocation.Text
                End If

                indexSama = Me.GetIndex
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)


                    AngkaDecimal = Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    DecDef = Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    DecMulti = Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                    DecRoll = Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, primaryRoll)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)

                    With row
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColProductCode.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                        .Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                        .Cells(Me.ColPONO.Index).Value = ""
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColGRNNO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtDiscAmount1.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text
                        Me.ObjProd.GetProductProperties(Me.TxtPrimaryCodeNo.Text, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
                        If LocationCont = False Then
                            .Cells(Me.ColIsLocation.Index).Value = "0"
                        Else
                            .Cells(Me.ColIsLocation.Index).Value = "1"
                        End If
                        If BatchNoCont = False Then
                            .Cells(Me.colIsBatch.Index).Value = "0"
                        Else
                            .Cells(Me.colIsBatch.Index).Value = "1"
                        End If
                        .Cells(Me.colIsExpDate.Index).Value = ExpiryDateCont
                        .Cells(Me.ColIsKain.Index).Value = ApakahKainTrading
                        ''@rosadinnn
                        ''@MUaT
                        .Cells(Me.ColTotalQty.Index).Value = "0"
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnitsDefault.Text

                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = Me.TxtCurrencyCode1.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtAmount.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryGRN.Index).Value = -1
                        .Cells(Me.ColPrimaryProd.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryPO.Index).Value = -1
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.TxtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency1.Text
                        .Cells(Me.ColShowDetail.Index).Value = "Show"
                        .Cells(Me.ColReferedToPO.Index).Value = 0
                        .Cells(Me.ColItemAdjustment.Index).Value = 0
                        .Cells(Me.ColItemAdjustmentReason.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentUser.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "#1-1-1950#"

                        .Cells(Me.ColTotalQtyMulti.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.LblUnitsMulti.Text
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.TxtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioMulti.Index).Value = Me.TxtRatioReference.Text
                        .Cells(Me.ColRatioDef.Index).Value = Me.TxtRatioDefault.Text
                        If Me.TxtPrimaryUnitsMulti.Text <> Me.TxtPrimaryUnitsDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                    ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)) _
                                    & " " & Me.LblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                    Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & " " & Me.LblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsCodeMulti.Index).Value
                        Else
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                        If ApakahKainTrading = True Then
                            .Cells(Me.ColDecMulti.Index).Value = DecRoll
                        Else
                            .Cells(Me.ColDecMulti.Index).Value = DecMulti
                        End If
                        .Cells(Me.ColDec.Index).Value = DecDef
                    End With
                    Me.dgvItem.ClearSelection()
                    IdxSelectedLast = -1
                    Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Selected = True
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty, 

                    ''@rosadinnn
                    ''@MUaT
                    'row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyMulti.Text)
                    If Me.TxtPrimaryUnitsMulti.Text <> row.Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        MsgBox("There is already  " & row.Cells(Me.ColTotalQtyMulti.Index).Value & " " & row.Cells(Me.ColUnit.Index).Value & " of " & Me.TxtCodeNo.Text & " exist." & vbNewLine & "You cannot input this item with different units.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text), ObjProd.GetAngkaDecimal(row.Cells(Me.ColPrimaryProd.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))
                    'row.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyDefault.Text), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))

                    Me.SetQtyDefaultOnDgv(indexSama, ApakahKainTrading)

                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)
                    Me.Hitung_Amount_Multi(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text

                Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                Me.FormatDGV2()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                        Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                Me.Hitung_Total_Multi()

               

                Me.TxtQtyMulti.Text = "0"
                Me.txtAmount.Text = "0"
                Me.TxtRemarkItem.Text = ""
                Me.KosongkanItem()
            Else
                'Me.KosongkanItem()
                'Me.TxtCodeNo.Focus()
            End If

            If PrimaryLocationDummy <> -1 Then
                Me.txtPrimaryLocation.Text = PrimaryLocationDummy
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
            End If
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub
    Private Sub Hitung_Total_Multi()
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Decimal = 0

        Try
            'hitung ulang unit price
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub Hitung_AtCost_Multi()
        Dim TotalOngkosLain As Decimal = 0
        Dim TotalCost As Decimal = 0
        Dim SubTotal As Decimal
        Dim row As DataGridViewRow
        Dim Acuan As Integer
        Dim DefPrec As Integer = 0

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            If Me.TxtRemarkItem.Enabled = False Then
                Acuan = Me.dgvCost.Rows.Count
            Else
                Acuan = Me.dgvCost.Rows.Count - 1
            End If

            'hitung total ongkos lain - lain
            For Each row In Me.dgvCost.Rows
                If row.Index <> Acuan Then
                    TotalOngkosLain += row.Cells(Me.ColCostAmount.Index).Value
                End If
            Next

            'For Each row In Me.dgvItem.Rows
            'SubTotal += row.Cells(Me.ColAmount.Index).Value
            'Next
            SubTotal = Me.GetSubTotal_TanpaReturn_Multi

            'hitung total harga FOB
            'jika rupiah maka langsung perhitungkan rupiah, jika tidak maka gunakan kurs konversi ke
            'rupiah
            TotalCost = Math.Round((Me.GetSubTotal_TanpaReturn_Multi * CDec(Me.TaxExchageRate)) + _
            TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero) ' --> duty stamp dan tax amount tidak diperhitungkan untuk At Cost
            'CDbl(Me.TxtDutyStamp.Text) + Me.txtTaxAmount.Text + 

            'hitung unit price untuk masing - masing item
            For Each row In Me.dgvItem.Rows
                If SubTotal = 0 Or TotalCost = 0 Then
                    row.Cells(Me.ColATCost.Index).Value = 0
                Else
                    If row.Cells(Me.ColTotalQtyMulti.Index).Value - Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryProd.Index).Value, _
                    row.Cells(Me.ColPrimaryGRN.Index).Value, -1) <> 0 Then
                        If row.Index <> Me.dgvItem.RowCount - 1 Then
                            row.Cells(Me.ColAtAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColAmount.Index).Value * TotalCost) / _
                            SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                            'row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
                            'SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        Else
                            row.Cells(Me.ColAtAmount.Index).Value = TotalCost - Me.GetTotalAtAmount_KecualiBarisAkhir
                            'row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
                            'SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        End If
                        'disini agak berbeda dg form quick acop, karena tidak return revisi aop
                        row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDec(row.Cells(Me.ColAtAmount.Index).Value) / _
                        CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), _
                        DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                    Else
                        row.Cells(Me.ColATCost.Index).Value = 0
                        row.Cells(Me.ColAtAmount.Index).Value = 0
                    End If
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.ColAtAmount.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function GetTotalAtAmount_KecualiBarisAkhir() As Decimal
        Dim d As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    d += CDec(rowDGV.Cells(Me.ColAtAmount.Index).Value)
                End If
            Next

            Return d
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetJmlDGVPR(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryGRN As Double, _
   ByVal vSecQty As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvPR.Rows
                If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColSecQty_PR.Index).Value = vSecQty Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetSubTotal_TanpaReturn_Multi() As Double
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Double = 0

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If Me.CalculateByIndex = 0 Then
                    'berdasarkan sales invoice
                    exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
                ElseIf Me.CalculateByIndex = 1 Then
                    'berdasarkan delivery order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
                ElseIf Me.CalculateByIndex = 2 Then
                    'berdasarkan saler order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
                End If

                'hitung purchase return jika ada
                Dim DefaultPrecisionRow As Double = 1
                Dim SubTotal1 As Double = 0
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

                SubTotal1 = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                     CDec(row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)) * exchangRate, DefaultPrecisionRow)

                SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
            Next

            Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetExchangeRate(ByVal vPrimaryTransaksi As Double) As Double
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvExchageRates.Rows
                If row.Cells(5).Value = vPrimaryTransaksi Then
                    Return row.Cells(3).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub FormatDGV2()
        Dim row As DataGridViewRow
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim TidakAdaPO As Boolean = True
        Dim TidakAdaMultiUnit As Boolean = True
        Dim AdaPriceByDefaultUnit As Boolean = False
        Dim TidakAdaKainTrading As Boolean = True

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                If CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) <> CDec(row.Cells(Me.ColTotalQty.Index).Value) Then
                    TidakAdaMultiUnit = False
                Else
                    If row.Cells(Me.ColIsKain.Index).Value = True Then
                        TidakAdaMultiUnit = False
                    End If
                End If
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    TidakAdaPO = False
                End If


                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    AdaPriceByDefaultUnit = True
                End If

            Next
            Me.ColPriceByMultiUnit_Str.Visible = AdaPriceByDefaultUnit

            '2. tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False

       
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
         
            ''po
            'If TidakAdaPO = False Then
            '    Me.ColPONO.Visible = True
            '    If Me.TxtRemarkItem.Enabled = True Then
            '        Me.txtSupplierCode.Enabled = False
            '        Me.txtSupplierCode.TabStop = False
            '        Me.btnFindSupplier.Enabled = False
            '        Me.btnFindSupplier.TabStop = False
            '    End If
            'Else
            '    Me.ColPONO.Visible = False
            '    If Me.TxtRemark.Enabled = True Then
            '        Me.txtSupplierCode.Enabled = True
            '        Me.txtSupplierCode.TabStop = True
            '        Me.btnFindSupplier.Enabled = True
            '        Me.btnFindSupplier.TabStop = True
            '    End If
            'End If
            'multiunit
            If TidakAdaMultiUnit = True Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If
           
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub Hitung_Amount_Multi(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value - _
            row.Cells(Me.ColDiscAmount.Index).Value, DefaultPrecisionRow))

            'hitung amount
            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItemAdjustment.Index).Value), DefaultPrecisionRow)
            Else
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value * _
                row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItemAdjustment.Index).Value), DefaultPrecisionRow)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer, ByVal vApakahKain As Boolean)
        Try
            If vApakahKain = False Then
                Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIdx).Value) * _
                                                         CDec(Me.dgvItem.Item(Me.ColRatioDef.Index, vRowIdx).Value / _
                                                         CDec(Me.dgvItem.Item(Me.ColRatioMulti.Index, vRowIdx).Value)), _
                                                         Me.ObjProd.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryProd.Index, vRowIdx).Value))
                'Else
                '    Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIdx).Value) * _
                '                                             CDec(Me.dgvItem.Item(Me.ColQtyAt.Index, vRowIdx).Value), _
                '                                             Me.ObjProd.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryProd.Index, vRowIdx).Value))
                '    Me.dgvItem.Item(Me.ColRatioDesc.Index, vRowIdx).Value = Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIdx).Value & " " & Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, vRowIdx).Value & " = " & _
                '                                             Me.dgvItem.Item(Me.ColQty.Index, vRowIdx).Value & " " & Me.dgvItem.Item(Me.ColUnit.Index, vRowIdx).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function GetIndex() As Integer
        Dim BatchNo As String = ""
        Dim PrimaryLocation As Double = -1
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryProd.Index).Value = Me.TxtPrimaryCodeNo.Text Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function ApakahDataItemBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
        Dim AngkaDecimal As Integer = 0

        Try
            'data salah jika:
            '1. tidak diisikan, kode product, currency
            'cek data yang harus diisi
            If Me.ObjProd.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text) = True Then
                If Me.TxtLocation.Text = "" Then
                    MsgBox("Please enter location first.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocation.Focus()
                    Return False
                End If
            End If
            If Me.TxtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            KoleksiNamaData.Add("Currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode1)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. jika sudah ada item, maka kode currencynya harus sama
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                If Me.TxtPrimaryCurrency1.Text <> rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                    MsgBox("Currency must be identical for each item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. berkaitan dengan data angka:
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            '   a. unit price, qty harus positif
            KoleksiNamaData.Add("Unit Price")
            KoleksiTextBox.Add(Me.TxtUnitPrice)
            'KoleksiNamaData.Add("Qty")
            'KoleksiTextBox.Add(Me.TxtQtyMulti)
            If ModCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            '   b. discount paling besar 100% --> otomatis + non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("discount")
            KoleksiTextBox.Add(Me.TxtDiscount)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '4. Remarks maksimal 150 karakter
            If Me.TxtRemarkItem.Text.Length > 150 Then
                MsgBox("Maximum remarks data length is 150 character", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ObjProdCust.ApakahKainTrading(Me.TxtPrimaryCodeNo.Text) = True Then
                If Me.CboUnitMulti.Text.ToString.StartsWith("ROLL") = False Then
                    MsgBox("Please input unit with ROLL", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                For Each rd As DataGridViewRow In Me.dgvItem.Rows
                    With rd
                        If rd.Cells(Me.ColIsKain.Index).Value = True Then
                            If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                MsgBox("Please specify price by " & rd.Cells(Me.ColUnit.Index).Value & " for row: " & rd.Index + 1, MsgBoxStyle.Information, "Information")
                                rd.Selected = True
                                Return False
                            End If
                        End If
                    End With
                Next
            End If

            'hanya bisa input ke dgv by product (+batch +qty at)
            Dim PesanText As String = ""
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If rd.Cells(Me.ColPrimaryProd.Index).Value = Me.TxtPrimaryCodeNo.Text Then
                    PesanText = "You can't input, because product: " & Me.TxtCodeNo.Text
                    rd.Selected = True
                    Exit For
                End If
            Next
            If PesanText <> "" Then
                PesanText &= " has been entered."
                MsgBox(PesanText, MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim hasil As MsgBoxResult
        Dim ObjCustom As New CustomSMESS.ClsTransaksiAOP_Quick2_AddByCustomProduct
        Try
            ObjCustom.IsiTblInOutByTblAOP(Me.TblItemCustomInOut, Me.TblItemCustom, Me.PrimaryTransaksiAOP)

            If ApakahDataBenar() = True Then

                hasil = MsgBox("Add item?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.FromOthBtn = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            hasil = MsgBox("Cancel add item and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Me.FromOthBtn = True
                Me.Close()
            Else
                cancelClose = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        If Me.TxtRemarkItem.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        End If
    End Sub

    Private Sub CboUnitMulti_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnitMulti.SelectedIndexChanged
        'Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow
        Try
            If Me.CboUnitMulti.SelectedIndex = -1 Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProd.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.TxtPrimaryUnitsDefault.Text = Me.TxtPrimaryUnitsMulti.Text
                Me.LblUnitsMulti.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text)
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = True
                Me.CboUnitMulti.Visible = False
                Me.CboUnitMulti.TabStop = False
                Me.LblUnitPrice.Text = "Unit Price/" & LblUnitsMulti.Text

                If Me.TxtPrimaryUnitsDefault.Text = "" Or Me.TxtPrimaryUnitsDefault.Text = "0" Then
                    Me.TxtPrimaryUnitsDefault.Text = "-1"
                End If
                If Me.TxtPrimaryUnitsMulti.Text = "" Or Me.TxtPrimaryUnitsMulti.Text = "0" Then
                    Me.TxtPrimaryUnitsMulti.Text = "-1"
                End If
                If Me.LblUnitsMulti.Text = "" Then
                    Me.LblUnitsMulti.Text = "Units"
                End If
                If Me.LblUnitsDefault.Text = "" Then
                    Me.LblUnitsDefault.Text = "Units"
                End If
            Else
                row = Me.tblCboUnit.DefaultView(Me.CboUnitMulti.SelectedIndex).Row
                Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, row.Primary), ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, row.Primary))
                Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, row.Primary), ObjProd.GetAngkaDecimal(row.Primary))
                Me.TxtPrimaryUnitsMulti.Text = row.Primary
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                Me.TxtPrimaryUnitsDefault.Text = ObjProd.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.LblUnitsMulti.Text = Me.CboUnitMulti.Text
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnitMulti.Text
            End If
            Me.SetQtyDefault()
            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub SetQtyDefault()
        Try
            'Me.TxtQtyDefaul.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPriceByMultiUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPriceByMultiUnit.SelectedIndexChanged
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow

        Try
            If Me.CboPriceByMultiUnit.SelectedIndex <> -1 Then
                If Me.TxtRemarkItem.Enabled = True Then
                    If Me.CboPriceByMultiUnit.Text <> Me.ObjProd.GetUnitCode(Me.TxtPrimaryCodeNo.Text) Then
                        Me.txtPriceByMultiUnit.Text = "1"
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"
                    End If

                    If Me.GetHistoricalPrice = True Then
                        Me.ObjPO.GetHistoricalPrice(Me.objUnit.GetPrimary(Me.CboPriceByMultiUnit.Text), Me.TxtPrimaryCodeNo.Text, _
                        Me.TxtPrimaryCurrency1.Text, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc, Me.txtPriceByMultiUnit.Text)

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                        Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                        Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                        Me.TxtDiscFromPerct.Text = DiscFromPerc
                    End If

                    If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        If CDec(Me.TxtUnitPrice.Text) > 0 Then
                            Me.Hitung_Amount2()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCodeNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.DialogResult = Windows.Forms.DialogResult.Cancel Then
                Me.cancelClose = True

            End If
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetGboItem()
                If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                    Me.TxtUnitPrice.Focus()
                Else
                    Me.TxtCurrencyCode1.Focus()
                End If

            Else
                Me.TxtCodeNo.Clear()
                Me.TxtDiscFromPerct.Text = "1"
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
            setLocation_Product()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

        End Try
    End Sub
End Class