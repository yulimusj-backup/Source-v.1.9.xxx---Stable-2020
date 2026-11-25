Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockOutSales_QtyMinus
    Public FrmRef As FrmTransaksiStockOutSales
    Public TBLMasterProductQty As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
    Public ListPrimaryProduct As New Queue(Of Double) 
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public DGVItem As DataGridView
 
    Dim ObjCD As New Core_Module.ClsCekData 

    Private Sub IsiTBLProductDetail(ByVal vListPrimary As Queue(Of Double), ByVal vTBL As dseTransaksiStockInOut_Sales.DataTable4DataTable)
        Dim DTA As New dseTransaksiStockInOut_SalesTableAdapters.DataTable4TableAdapter
        Dim TBLDummy As New dseTransaksiStockInOut_Sales.DataTable4DataTable
        Dim row, rowDummy As dseTransaksiStockInOut_Sales.DataTable4Row

        Try
            vTBL.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each PrimaryProduct As Double In ListPrimaryProduct
                TBLDummy.Rows.Clear()
                DTA.Fill(TBLDummy, PrimaryProduct)
                For Each rowDummy In TBLDummy.Rows
                    row = vTBL.NewDataTable4Row
                    For jk As Integer = 0 To vTBL.Columns.Count - 1
                        row.Item(jk) = rowDummy.Item(jk)
                    Next
                    vTBL.Rows.Add(row)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutSales_QtyMinus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim RowMasterProductQty As Core_Module.dseMasterProduct.tblmaster_productqtyRow 
        Dim TBLDetail As New dseTransaksiStockInOut_Sales.DataTable4DataTable
        Dim RowDetail As dseTransaksiStockInOut_Sales.DataTable4Row
        Dim DTAStorage As New dseTransaksiStockInOut_SalesTableAdapters.DataTable5TableAdapter
        Dim TBLStorage As New dseTransaksiStockInOut_Sales.DataTable5DataTable
        Dim rowStorage As dseTransaksiStockInOut_Sales.DataTable5Row
        Dim AdaExpiry As Boolean = False
        Dim AdaBatch As Boolean = False
        Dim AdaLocation As Boolean = False

        Try
            Me.dgvMinus.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvMinus.Rows.Clear()

            'isikan table detail 
            Me.IsiTBLProductDetail(ListPrimaryProduct, TBLDetail)
            DTAStorage.Fill(TBLStorage)

            For Each RowMasterProductQty In Me.TBLMasterProductQty.Rows
                If RowMasterProductQty.Qty < 0 Then
                    If ListPrimaryProduct.Contains(RowMasterProductQty.Primary_Product) = True Then
                        If RowMasterProductQty.Expiry_Date <> #1/1/1950# Then
                            AdaExpiry = True
                        End If
                        If RowMasterProductQty.Primary_Storage_Area <> -1 Then
                            AdaLocation = True
                        End If
                        If RowMasterProductQty.Lot_No <> "" Then
                            AdaBatch = True
                        End If
                        TBLDetail.DefaultView.RowFilter = "Primary = " + RowMasterProductQty.Primary_Product.ToString
                        RowDetail = TBLDetail.DefaultView(0).Row

                        Me.dgvMinus.Rows.Add()
                        rowDGV = Me.dgvMinus.Rows(Me.dgvMinus.Rows.Count - 1)
                        rowDGV.Cells(Me.ColBatch.Index).Value = RowMasterProductQty.Lot_No
                        rowDGV.Cells(Me.ColCodeNo.Index).Value = RowDetail.Product_Code
                        rowDGV.Cells(Me.ColDesc.Index).Value = RowDetail.Description
                        If RowMasterProductQty.Expiry_Date = #1/1/1950# Then
                            rowDGV.Cells(Me.ColExpiry.Index).Value = Format("-")
                        Else
                            rowDGV.Cells(Me.ColExpiry.Index).Value = Format(RowMasterProductQty.Expiry_Date, "dd MMM yyyy")
                        End If
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvMinus.Rows.Count
                        rowDGV.Cells(Me.ColShortage.Index).Value = Me.ObjCD.DigitGrouping(RowMasterProductQty.Qty, RowDetail._Decimal)
                        rowDGV.Cells(Me.ColUnit.Index).Value = RowDetail.Unit_Code
                        If RowMasterProductQty.Primary_Storage_Area = -1 Then
                            rowDGV.Cells(Me.ColLocation.Index).Value = ""
                            rowDGV.Cells(Me.ColWarehouse.Index).Value = ""
                        Else
                            TBLStorage.DefaultView.RowFilter = "Primary = " + RowMasterProductQty.Primary_Storage_Area.ToString
                            rowStorage = TBLStorage.DefaultView(0).Row
                            rowDGV.Cells(Me.ColLocation.Index).Value = rowStorage.Storage_Area_Code
                            rowDGV.Cells(Me.ColWarehouse.Index).Value = rowStorage.Warehouse_Code
                        End If

                        For Each rowDGV1 In Me.DGVItem.rows
                            If rowDGV1.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = RowMasterProductQty.Primary_Product And _
                            rowDGV1.Cells(Me.FrmRef.ColPrimaryLocation.Index).Value = RowMasterProductQty.Primary_Storage_Area And _
                            rowDGV1.Cells(Me.FrmRef.ColBatch.Index).Value = RowMasterProductQty.Lot_No Then
                                rowDGV.Cells(Me.ColQtyStockOut.Index).Value = rowDGV1.Cells(Me.FrmRef.ColQty.Index).Value
                                rowDGV.Cells(Me.ColQtyOnHand.Index).Value = Me.ObjCD.DigitGrouping(Math.Abs(CDec(rowDGV.Cells(Me.ColShortage.Index).Value)) + _
                                CDec(rowDGV.Cells(Me.ColQtyStockOut.Index).Value), RowDetail._Decimal)
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
            If AdaBatch = True Then
                Me.ColBatch.Visible = True
            Else
                Me.ColBatch.Visible = False
            End If
            If AdaExpiry = True Then
                Me.ColExpiry.Visible = True
            Else
                Me.ColExpiry.Visible = False
            End If
            If AdaLocation = True Then
                Me.ColLocation.Visible = True
            Else
                Me.ColLocation.Visible = False
            End If
            Me.dgvMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class