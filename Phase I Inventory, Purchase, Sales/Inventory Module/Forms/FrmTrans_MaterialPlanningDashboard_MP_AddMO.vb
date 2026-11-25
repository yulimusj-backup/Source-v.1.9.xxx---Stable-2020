Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_MP_AddMO
    Dim LangsungClose As Boolean = False
    Dim ObjCD As New Core_Module.ClsCekData
    Public rowCodeNo As DseTransaksiMaterialPlanning3.DataTable1Row
    Dim rowItemMO As DseTransaksiMaterialPlanning3.DataTable2Row
    Dim rowItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
    Dim rowItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
    Dim rowItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
    Dim rowItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
    Dim rowItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
    Dim rowItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow
    Dim Buffer As String = ""

    Public TblCodeNo As New DseTransaksiMaterialPlanning3.DataTable1DataTable
    Public TblItemMO As New DseTransaksiMaterialPlanning3.DataTable2DataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable
    Public vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable
    Public vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable
    Public vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable
    Public vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable
    Public vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable

    Enum EnumPrioritasMO As Integer
        DO_
        SO
        RM_Req
        PrOB
    End Enum

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'cek apakah data benar -> dilakukan ketika cell end edit
            If MsgBox("Start creating MO?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_AddMO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Cancel MO creation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboNoDeliveryDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboNoDeliveryDate.CheckedChanged
        Try
            If Me.rboNoDeliveryDate.Checked = True Then
                Me.dgvItemMO.Rows.Clear()
                Me.dgvCodeNo.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboSetDeliveryDate2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSetDeliveryDate2.CheckedChanged
        Try
            If Me.rboSetDeliveryDate2.Checked = True Then
                Me.IsikanDGVCodeNo()
                Me.dgvCodeNo.ClearSelection()
                If Me.dgvCodeNo.Rows.Count > 0 Then
                    Me.dgvCodeNo.Rows(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVCodeNo()
        Dim rowCodeNo As DseTransaksiMaterialPlanning3.DataTable1Row
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvCodeNo.ClearSelection()
            Me.dgvCodeNo.Rows.Clear()
            Me.dgvCodeNo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            For Each rowCodeNo In Me.TblCodeNo.Rows
                Me.dgvCodeNo.Rows.Add()
                rowDGV = Me.dgvCodeNo.Rows(Me.dgvCodeNo.Rows.Count - 1)
                rowDGV.Cells(Me.ColHeader_CodeNo.Index).Value = rowCodeNo.Code_No
                rowDGV.Cells(Me.ColHeader_Desc.Index).Value = rowCodeNo.Desc
                rowDGV.Cells(Me.ColHeader_QtyOrder.Index).Value = Me.ObjCD.DigitGrouping(rowCodeNo.Qty_Order, rowCodeNo.DefPrec)
                rowDGV.Cells(Me.ColHeader_QtyMO.Index).Value = Me.ObjCD.DigitGrouping(rowCodeNo.Qty_MO, rowCodeNo.DefPrec)
                rowDGV.Cells(Me.ColHeader_Unit.Index).Value = rowCodeNo.Unit_Code
                rowDGV.Cells(Me.ColHeader_DefPrec.Index).Value = rowCodeNo.DefPrec
                rowDGV.Cells(Me.ColHeader_PrimaryCodeNo.Index).Value = rowCodeNo.Primary_Product
            Next

            Me.dgvCodeNo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_AddMO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.KonversiDataTable()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KonversiDataTable()
        '1. pindahkan data ke tabel item mo
        '2. isikan data mo

        Dim RowCodeNo As DseTransaksiMaterialPlanning3.DataTable1Row
        Dim RowItemMO As DseTransaksiMaterialPlanning3.DataTable2Row
        Dim DV As Data.DataRowView
        Dim JmlMODummy As Decimal
        Dim Idx As Integer = 0

        Try
            Me.vDT_ItemSOBDelivery.DefaultView.Sort = "delivery date"
            Me.TblItemMO.DefaultView.Sort = "[Req Date], [Tipe Doc Int]"

            '1. pindahkan data ke tabel item mo
            For Each RowCodeNo In Me.TblCodeNo.Rows
                Me.vDT_ItemDO.DefaultView.RowFilter = "[primary product] = " + RowCodeNo.Primary_Product.ToString
                For Each DV In Me.vDT_ItemDO.DefaultView
                    Me.rowItemDO = DV.Row
                    RowItemMO = Me.TblItemMO.NewDataTable2Row
                    RowItemMO.DefPrec = RowCodeNo.DefPrec
                    RowItemMO.Doc_No = Me.rowItemDO.DO_Number
                    RowItemMO.Primary_Product = RowCodeNo.Primary_Product
                    RowItemMO.Primary_Unit = RowCodeNo.Primary_Unit
                    RowItemMO.Qty_MO = 0
                    RowItemMO.Qty_Order_Balance = Me.rowItemDO.qty
                    RowItemMO.Req_Date = Me.rowItemDO.DO_Date
                    RowItemMO.Tipe_Doc_Int = CInt(FrmTrans_MaterialPlanningDashboard_MP_AddMO.EnumPrioritasMO.DO_)
                    RowItemMO.Tipe_Doc_Str = "Pending DO"
                    RowItemMO.Unit_Code = RowCodeNo.Unit_Code
                    RowItemMO.Is_WIP = 0

                    Me.TblItemMO.Rows.Add(RowItemMO)
                Next

                Me.vDT_ItemSOB.DefaultView.RowFilter = "[primary product] = " + RowCodeNo.Primary_Product.ToString
                For Each DV In Me.vDT_ItemSOB.DefaultView
                    Me.rowItemSOB = DV.Row
                    RowItemMO = Me.TblItemMO.NewDataTable2Row
                    RowItemMO.DefPrec = RowCodeNo.DefPrec
                    RowItemMO.Doc_No = Me.rowItemSOB.No_Transaksi
                    RowItemMO.Primary_Product = RowCodeNo.Primary_Product
                    RowItemMO.Primary_Unit = RowCodeNo.Primary_Unit
                    RowItemMO.Qty_MO = 0
                    RowItemMO.Qty_Order_Balance = Me.rowItemSOB.qty
                    If rowItemSOB.delivery_date_cont = 1 Then
                        RowItemMO.Req_Date = Me.GetRequiredDate(Me.rowItemSOB.identifier_so, Me.rowItemSOB.qty)
                    Else
                        RowItemMO.Req_Date = #1/1/2200#
                    End If
                    RowItemMO.Tipe_Doc_Int = CInt(FrmTrans_MaterialPlanningDashboard_MP_AddMO.EnumPrioritasMO.SO)
                    RowItemMO.Tipe_Doc_Str = "Sales Order"
                    RowItemMO.Unit_Code = RowCodeNo.Unit_Code
                    RowItemMO.Is_WIP = 0

                    Me.TblItemMO.Rows.Add(RowItemMO)
                Next

                Me.vDT_ItemPrOB.DefaultView.RowFilter = "[primary product] = " + RowCodeNo.Primary_Product.ToString
                For Each DV In Me.vDT_ItemPrOB.DefaultView
                    Me.rowItemPrOB = DV.Row
                    RowItemMO = Me.TblItemMO.NewDataTable2Row
                    RowItemMO.DefPrec = RowCodeNo.DefPrec
                    RowItemMO.Doc_No = Me.rowItemPrOB.No_Transaksi
                    RowItemMO.Primary_Product = RowCodeNo.Primary_Product
                    RowItemMO.Primary_Unit = RowCodeNo.Primary_Unit
                    RowItemMO.Qty_MO = 0
                    RowItemMO.Qty_Order_Balance = Me.rowItemPrOB.qty
                    If rowItemPrOB.delivery_date_cont = 1 Then
                        RowItemMO.Req_Date = Me.GetRequiredDate_PrOB(Me.rowItemPrOB.identifier_pr_item, Me.rowItemPrOB.qty)
                    Else
                        RowItemMO.Req_Date = #1/1/2200#
                    End If
                    RowItemMO.Tipe_Doc_Int = CInt(FrmTrans_MaterialPlanningDashboard_MP_AddMO.EnumPrioritasMO.PrOB)
                    RowItemMO.Tipe_Doc_Str = "Purchase Request"
                    RowItemMO.Unit_Code = RowCodeNo.Unit_Code
                    RowItemMO.Is_WIP = 0

                    Me.TblItemMO.Rows.Add(RowItemMO)
                Next

                Me.vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + RowCodeNo.Primary_Product.ToString
                For Each DV In Me.vDT_ItemRMReq.DefaultView
                    Me.rowItemRMReq = DV.Row
                    RowItemMO = Me.TblItemMO.NewDataTable2Row
                    RowItemMO.DefPrec = RowCodeNo.DefPrec
                    RowItemMO.Doc_No = Me.rowItemRMReq.jo_no_rec
                    RowItemMO.Primary_Product = RowCodeNo.Primary_Product
                    RowItemMO.Primary_Unit = RowCodeNo.Primary_Unit
                    RowItemMO.Qty_MO = 0
                    RowItemMO.Qty_Order_Balance = Me.rowItemRMReq.qty
                    RowItemMO.Req_Date = Me.rowItemRMReq.required_date
                    RowItemMO.Tipe_Doc_Int = CInt(FrmTrans_MaterialPlanningDashboard_MP_AddMO.EnumPrioritasMO.RM_Req)
                    RowItemMO.Tipe_Doc_Str = "Raw Material Request"
                    RowItemMO.Unit_Code = RowCodeNo.Unit_Code
                    RowItemMO.Is_WIP = Me.rowItemRMReq.is_wip

                    Me.TblItemMO.Rows.Add(RowItemMO)
                Next

                '2. isikan data mo
                Me.TblItemMO.DefaultView.RowFilter = "[Primary Product] = " + RowCodeNo.Primary_Product.ToString
                JmlMODummy = 0
                Idx = 0

                While JmlMODummy <= RowCodeNo.Qty_Order And Idx <= Me.TblItemMO.DefaultView.Count - 1
                    Me.rowItemMO = Me.TblItemMO.DefaultView(Idx).Row

                    If RowCodeNo.Qty_Order - JmlMODummy > Me.rowItemMO.Qty_Order_Balance Then
                        Me.rowItemMO.Qty_MO = Me.rowItemMO.Qty_Order_Balance
                    Else
                        Me.rowItemMO.Qty_MO = RowCodeNo.Qty_Order - JmlMODummy
                    End If

                    JmlMODummy += Me.rowItemMO.Qty_MO
                    Idx += 1
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetRequiredDate_PrOB(ByVal vIdentifierPrOB As String, ByVal vQtyOrderBalance As Decimal) As Date
        Dim vJmlDummy As Decimal = vQtyOrderBalance

        Try
            Me.vDT_ItemPrOBDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifierPrOB + "'"

            If Me.vDT_ItemPrOBDelivery.DefaultView.Count = 0 Then
                Return #1/1/2200#
            Else
                For ij As Double = Me.vDT_ItemPrOBDelivery.DefaultView.Count - 1 To 0 Step -1
                    Me.rowItemPrOBDelivery = Me.vDT_ItemPrOBDelivery.DefaultView(ij).Row
                    If vJmlDummy <= Me.rowItemPrOBDelivery.delivery_qty Then
                        Return Me.rowItemPrOBDelivery.delivery_date
                    Else
                        vJmlDummy -= Me.rowItemPrOBDelivery.delivery_qty
                    End If
                Next

                Return #1/1/2200#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetRequiredDate(ByVal vIdentifierSO As String, ByVal vQtyOrderBalance As Decimal) As Date
        Dim vJmlDummy As Decimal = vQtyOrderBalance

        Try
            Me.vDT_ItemSOBDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifierSO + "'"

            If Me.vDT_ItemSOBDelivery.DefaultView.Count = 0 Then
                Return #1/1/2200#
            Else
                For ij As Double = Me.vDT_ItemSOBDelivery.DefaultView.Count - 1 To 0 Step -1
                    Me.rowItemSOBDelivery = Me.vDT_ItemSOBDelivery.DefaultView(ij).Row
                    If vJmlDummy <= Me.rowItemSOBDelivery.delivery_qty Then
                        Return Me.rowItemSOBDelivery.delivery_date
                    Else
                        vJmlDummy -= Me.rowItemSOBDelivery.delivery_qty
                    End If
                Next

                Return #1/1/2200#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvCodeNo_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCodeNo.SelectionChanged
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCodeNo As Double = -1
        Dim DV As Data.DataRowView 

        Try
            Me.dgvItemMO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemMO.Rows.Clear()

            If Me.dgvCodeNo.SelectedCells.Count > 0 Then
                PrimaryCodeNo = Me.dgvCodeNo.Item(Me.ColHeader_PrimaryCodeNo.Index, Me.dgvCodeNo.SelectedCells(0).RowIndex).Value 

                Me.TblItemMO.DefaultView.RowFilter = "[Primary Product] = " + PrimaryCodeNo.ToString
                Me.TblItemMO.DefaultView.Sort = "[Req Date] ASC, [Tipe Doc Int] DESC"

                For Each DV In Me.TblItemMO.DefaultView
                    Me.rowItemMO = DV.Row
                    Me.dgvItemMO.Rows.Add()
                    rowDGV = Me.dgvItemMO.Rows(Me.dgvItemMO.Rows.Count - 1)
                    With Me.rowItemMO
                        rowDGV.Cells(Me.ColItem_DocNo.Index).Value = .Doc_No
                        rowDGV.Cells(Me.ColItem_DocType.Index).Value = .Tipe_Doc_Str
                        rowDGV.Cells(Me.ColItem_DocTypInt.Index).Value = .Tipe_Doc_Int
                        rowDGV.Cells(Me.ColItem_QtyMo.Index).Value = Me.ObjCD.DigitGrouping(.Qty_MO, .DefPrec)
                        rowDGV.Cells(Me.ColItem_QtyOrderBalance.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Order_Balance, .DefPrec)
                        If .Req_Date = #1/1/2200# Then
                            rowDGV.Cells(Me.ColItem_RequiredDate.Index).Value = "-"
                        Else
                            rowDGV.Cells(Me.ColItem_RequiredDate.Index).Value = Format(.Req_Date, "dd MMM yyyy")
                        End If
                        rowDGV.Cells(Me.ColItem_Unit.Index).Value = .Unit_Code
                        rowDGV.Cells(Me.ColItem_DefPrec.Index).Value = .DefPrec
                        rowDGV.Cells(Me.ColItem_PrimaryCodeNo.Index).Value = .Primary_Product
                        If .Is_WIP = 1 Then
                            rowDGV.Cells(Me.ColItem_IsWIP.Index).Value = Boolean.TrueString
                        Else
                            rowDGV.Cells(Me.ColItem_IsWIP.Index).Value = Boolean.FalseString
                        End If
                    End With
                Next
            End If

            Me.dgvItemMO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemMO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemMO.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColItem_QtyMo.Index Then
                Me.Buffer = Me.dgvItemMO.Item(Me.ColItem_QtyMo.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemMO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemMO.CellEndEdit
        Dim rowDGV As DataGridViewRow
        Dim IsWIP As Integer = 0

        Try
            If e.ColumnIndex = Me.ColItem_QtyMo.Index Then
                If IsNothing(Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    If CDec(Me.Buffer) = CDec(Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value) Then
                        Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        If Me.ObjCD.CekNonNegatif(Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        Else
                            Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value, _
                            Me.dgvItemMO.Item(Me.ColItem_DefPrec.Index, e.RowIndex).Value)

                            'cek batasan vs qty order pe dokument
                            If CDec(Me.dgvItemMO.Item(Me.ColItem_QtyMo.Index, e.RowIndex).Value) > _
                            CDec(Me.dgvItemMO.Item(Me.ColItem_QtyOrderBalance.Index, e.RowIndex).Value) Then
                                MsgBox("Please specify MO Qty <= Order Balance Qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If

                            'cek batasan vs qty order per code no
                            For Each Me.rowCodeNo In Me.TblCodeNo.Rows
                                If Me.rowCodeNo.Primary_Product = Me.dgvItemMO(Me.ColItem_PrimaryCodeNo.Index, e.RowIndex).Value Then
                                    If Me.rowCodeNo.Qty_Order < Me.rowCodeNo.Qty_MO - CDec(Me.Buffer) + Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        MsgBox("Total Order Qty for " + Me.rowCodeNo.Code_No + " is " + Me.ObjCD.DigitGrouping(Me.rowCodeNo.Qty_Order, Me.rowCodeNo.DefPrec) + _
                                        " " + Me.rowCodeNo.Unit_Code + ", please make sure this limitation is followed.", MsgBoxStyle.Information, "Information")
                                        Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Exit Sub
                                    End If

                                    'update qty baru di tbl item mo
                                    For Each Me.rowItemMO In Me.TblItemMO.Rows
                                        If Me.rowItemMO.Primary_Product = Me.rowCodeNo.Primary_Product Then
                                            If Me.rowItemMO.Doc_No = Me.dgvItemMO.Item(Me.ColItem_DocNo.Index, e.RowIndex).Value And _
                                            Me.rowItemMO.Tipe_Doc_Int = Me.dgvItemMO.Item(Me.ColItem_DocTypInt.Index, e.RowIndex).Value Then
                                                If Me.dgvItemMO.Item(Me.ColItem_IsWIP.Index, e.RowIndex).Value = Boolean.TrueString Then
                                                    IsWIP = 1
                                                Else
                                                    IsWIP = 0
                                                End If

                                                If Me.rowItemMO.Is_WIP = IsWIP Then
                                                    Me.rowItemMO.Qty_MO = CDec(Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value)
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next

                                    'update qty mo baru di tbl code no
                                    Me.rowCodeNo.Qty_MO = Me.rowCodeNo.Qty_MO - CDec(Me.Buffer) + Me.dgvItemMO.Item(e.ColumnIndex, e.RowIndex).Value

                                    'update qty mo baru di dgv code no
                                    For Each rowDGV In Me.dgvCodeNo.Rows
                                        If rowDGV.Cells(Me.ColHeader_PrimaryCodeNo.Index).Value = Me.rowCodeNo.Primary_Product Then
                                            rowDGV.Cells(Me.ColHeader_QtyMO.Index).Value = Me.ObjCD.DigitGrouping(Me.rowCodeNo.Qty_MO, Me.rowCodeNo.DefPrec)
                                            Exit For
                                        End If
                                    Next

                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 
End Class