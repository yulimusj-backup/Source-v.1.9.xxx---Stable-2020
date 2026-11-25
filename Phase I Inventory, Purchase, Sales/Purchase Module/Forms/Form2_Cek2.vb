Public Class Form2_Cek2
    Dim CodeNo, Desc As String
    Dim QtySC, QtyQty As Decimal
    Dim PrimaryProduct As Double
    Dim DefPrec As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
    Dim ObjProd As New Core_Module.clsMasterProduct

    Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
    Dim DTASC As New Inventory_Module.DseCekTableAdapters.tbltransaksi_stock_cardTableAdapter
    Dim DTAQty As New Inventory_Module.DseCekTableAdapters.tblmaster_productqtyTableAdapter
    Dim DTAProduct As New Inventory_Module.DseCekTableAdapters.tblmaster_productTableAdapter
    Dim TBLProduct As New Inventory_Module.DseCek.tblmaster_productDataTable
    Dim rowProduct As Inventory_Module.DseCek.tblmaster_productRow
    Dim DTATransfer As New Inventory_Module.DseCekTableAdapters.tbltrans_location_transferTableAdapter
    Dim TBLTransfer As New Inventory_Module.DseCek.tbltrans_location_transferDataTable
    Dim rowTransfer As Inventory_Module.DseCek.tbltrans_location_transferRow
    Dim DTABundlingBefore As New Inventory_Module.DseCekTableAdapters.tbltrans_stock_bundling_item_beforeTableAdapter
    Dim TBLBundlingBefore As New Inventory_Module.DseCek.tbltrans_stock_bundling_item_beforeDataTable
    Dim rowBundlingBefore As Inventory_Module.DseCek.tbltrans_stock_bundling_item_beforeRow
    Dim DTABundlingAfter As New Inventory_Module.DseCekTableAdapters.tbltrans_stock_bundling_item_afterTableAdapter
    Dim TBLBundlingAfter As New Inventory_Module.DseCek.tbltrans_stock_bundling_item_afterDataTable
    Dim rowBundlingAfter As Inventory_Module.DseCek.tbltrans_stock_bundling_item_afterRow
    Dim DTAItemInOut As New Inventory_Module.DseCekTableAdapters.tbltrans_stock_inout_itemTableAdapter
    Dim TBLItemInOut As New Inventory_Module.DseCek.tbltrans_stock_inout_itemDataTable
    Dim rowItemInOut As Inventory_Module.DseCek.tbltrans_stock_inout_itemRow
    Dim DTAST As New Inventory_Module.dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_itempenyesuaian1TableAdapter
    Dim TBLST As New Inventory_Module.dseTransaksiStockTaking.tbltrans_stock_taking_itempenyesuaian1DataTable
    Dim rowST As Inventory_Module.dseTransaksiStockTaking.tbltrans_stock_taking_itempenyesuaian1Row

    Private Sub CekData()
        Dim DTAProduct As New DseCekTableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New DseCek.tblmaster_productDataTable
        Dim rowProduct As DseCek.tblmaster_productRow
        Dim DTASC As New DseCekTableAdapters.DataTable3TableAdapter
        Dim TBLSC As New DseCek.DataTable3DataTable
        Dim rowSc As DseCek.DataTable3Row
        Dim DTAQty As New DseCekTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLQty As New DseCek.tblmaster_productqtyDataTable
        Dim rowQty As DseCek.tblmaster_productqtyRow
        Dim QtySC, QtyQty As Decimal
        Dim ListPrimaryProduct As New List(Of Double)
        Dim DTA1 As New Core_Module.DseMasterProduct1TableAdapters.tblmaster_productqty1TableAdapter
        Dim TBL1 As New Core_Module.DseMasterProduct1.tblmaster_productqty1DataTable
        Dim row1 As Core_Module.DseMasterProduct1.tblmaster_productqty1Row

        Try
            DTA1.Fill(TBL1)
            For Each row1 In TBL1.Rows
                ListPrimaryProduct.Add(row1.Primary_Product)
            Next
            Me.ProgressBar1.Value = 0

            DTAProduct.Fill(TBLProduct)
            DTASC.Fill(TBLSC)
            DTAQty.Fill(TBLQty)
            Me.ProgressBar1.Maximum = TBLProduct.Rows.Count

            For Each rowProduct In TBLProduct.Rows
                Me.ProgressBar1.Value += 1
                Me.LblProgress.Text = "Check stock data incosistency: " + rowProduct.Product_Code

                QtySC = 0
                QtyQty = 0

                TBLSC.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                If TBLSC.DefaultView.Count = 0 Then
                    QtySC = 0
                Else
                    For Each DV As DataRowView In TBLSC.DefaultView
                        rowSc = DV.Row
                        QtySC += rowSc.Qty
                    Next
                End If
                TBLQty.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                If TBLQty.DefaultView.Count = 0 Then
                    QtyQty = 0
                Else
                    For Each DV As DataRowView In TBLQty.DefaultView
                        rowQty = DV.Row
                        QtyQty += rowQty.Qty
                    Next
                End If
                If QtySC <> QtyQty Or QtySC < 0 Or QtyQty < 0 Or ListPrimaryProduct.Contains(rowProduct.Primary) Then
                    Me.CodeNo = rowProduct.Product_Code
                    Me.Desc = rowProduct.Description
                    Me.QtyQty = QtyQty
                    Me.QtySC = QtySC
                    Me.DefPrec = rowProduct._Decimal
                    Me.PrimaryProduct = rowProduct.Primary
                    ListPrimaryProduct.Add(rowProduct.Primary)
                    Me.BackgroundWorker1.ReportProgress(1)
                End If
            Next
            Me.LblProgress.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.Rows.Clear()
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.CekData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Dim rowDGV As DataGridViewRow

        Try
            Me.DataGridView1.Rows.Add()
            rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
            rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.CodeNo
            rowDGV.Cells(Me.ColDesc.Index).Value = Me.Desc
            rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Me.QtyQty, Me.DefPrec)
            rowDGV.Cells(Me.ColQtySC.Index).Value = Me.ObjCD.DigitGrouping(Me.QtySC, Me.DefPrec)
            rowDGV.Cells(Me.ColPrimaryProduct.Index).Value = Me.PrimaryProduct
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.ProgressBar1.Value = 0
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Data checked")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If MsgBox("Start repair?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker2.RunWorkerAsync()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim InOut As Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Decimal
        Dim DTASC As New Inventory_Module.dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim DescSC As String = ""
        Dim TotalLine As Decimal = 0
        Dim StartTime As Date
        Dim ElapseTime As Decimal = 0
        Dim RemainingTime As Decimal = 0
        Dim Rate As Decimal = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow

        Try
            ObjCD.UpdateThreading()

            Me.UpdateCodeNo()

            For Each rowDGV In Me.DataGridView1.Rows
                StartTime = Now
                Me.LblProgress.Text = "Preparing data " + rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString
                Me.DataGridView1.ClearSelection()
                rowDGV.Selected = True

                Me.DTAProduct.Fill(Me.TBLProduct, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)
                For Each Me.rowProduct In Me.TBLProduct.Rows
                    'hapuskan semua stock card -> hanya untuk category tertentu saja
                    Me.DTASC.DeleteQuery_By_Primary_Product(Me.rowProduct.Primary)
                    'hapuskan semua tbl product qty -> hanya untuk category tertentu saja
                    Me.DTAQty.DeleteQuery_byPrimaryProduct(Me.rowProduct.Primary)
                Next

                'load item stock in out
                Me.TBLItemInOut.Rows.Clear()
                Me.DTAItemInOut.Fill(Me.TBLItemInOut, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)

                'load item location transfer
                Me.TBLTransfer.Rows.Clear()
                Me.DTATransfer.Fill(Me.TBLTransfer, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)

                'load item bundling before
                Me.TBLBundlingBefore.Rows.Clear()
                Me.DTABundlingBefore.Fill(Me.TBLBundlingBefore, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)

                'load item bundling after
                Me.TBLBundlingAfter.Rows.Clear()
                Me.DTABundlingAfter.Fill(Me.TBLBundlingAfter, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)

                'load item stock taking adjustment
                Me.DTAST.FillBy_PrimaryProduct(TBLST, rowDGV.Cells(Me.ColPrimaryProduct.Index).Value)

                'untuk setiap baris item stock in out,
                TotalLine = Me.TBLItemInOut.Rows.Count + Me.TBLTransfer.Rows.Count + Me.TBLBundlingBefore.Rows.Count + Me.TBLBundlingAfter.Rows.Count + Me.TBLST.Rows.Count

                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = TotalLine

                For Each rowST In TBLST.Rows
                    Me.ProgressBar1.Value += 1
                    ElapseTime = (Now - StartTime).TotalSeconds
                    Rate = Me.ProgressBar1.Value / ElapseTime
                    RemainingTime = (Me.ProgressBar1.Maximum - Me.ProgressBar1.Value) / Rate
                    RemainingTime = RemainingTime / 60
                    Me.LblProgress.Text = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + " - Stock Taking - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Remaning time: " + ObjCD.Convert_ToString(RemainingTime)

                    If rowST.Qty <> 0 Then
                        JenisSC = Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking
                        If rowST.Qty > 0 Then
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Stock Taking"

                            Me.Koreksi4SimpanIn(rowST.Start_Date, ObjProd, rowST.Primary_Product, DTA1, rowST.Qty, _
                            -1, rowST.Primary_Location, rowST.Batch_No, -1, rowST.Expiry_Date)
                        Else
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Stock Taking"

                            Me.Koreksi4SimpanOut(rowST.Start_Date, ObjProd, rowST.Primary_Product, DTA1, rowST.Qty, _
                            -1, rowST.Primary_Location, rowST.Batch_No, rowST.Expiry_Date)
                        End If

                        If InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                            Qty_In = rowST.Qty
                            Qty_Out = 0
                        Else
                            Qty_In = 0
                            Qty_Out = rowST.Qty
                        End If

                        ObjSC.Simpan(rowST.Primary_Stock_Taking, JenisSC, rowST.No_Transaksi, rowST.Primary_Product, _
                        Qty_In, Qty_Out, DescSC, InOut, rowST.Start_Date, DTASC, rowST.Batch_No, _
                        rowST.Expiry_Date, rowST.Primary_Location, "")
                    End If

                Next

                For Each rowItemInOut In Me.TBLItemInOut.Rows
                    Me.ProgressBar1.Value += 1
                    ElapseTime = (Now - StartTime).TotalSeconds
                    Rate = Me.ProgressBar1.Value / ElapseTime
                    RemainingTime = (Me.ProgressBar1.Maximum - Me.ProgressBar1.Value) / Rate
                    RemainingTime = RemainingTime / 60
                    Me.LblProgress.Text = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + " - Stock In Out - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Remaning time: " + ObjCD.Convert_ToString(RemainingTime)

                    'cek jenis transaksi stock card
                    JenisSC = Me.GetTipeStockCard(rowItemInOut.Tipe_Transaksi)
                    'cek in/out
                    Select Case CInt(rowItemInOut.Tipe_Transaksi)
                        Case CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust)
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Stock in adjust"
                            Exit Select
                        Case CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess)
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Stock in back process"
                            Exit Select
                        Case CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process)
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Stock in process"
                            Exit Select
                        Case CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Purchase from " + rowItemInOut.Supplier_Name
                            Exit Select
                        Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                            DescSC = "Sales return from " + rowItemInOut.Customer_Name
                            Exit Select
                        Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                            DescSC = "Stock out adjust"
                            Exit Select
                        Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                            DescSC = "Stock out process"
                            Exit Select
                        Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                            DescSC = "Purchase return to " + rowItemInOut.Supplier_Name
                            Exit Select
                        Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                            InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                            DescSC = "Sales to " + rowItemInOut.Customer_Name
                            Exit Select
                    End Select

                    'simpan product qty
                    If CInt(rowItemInOut.Tipe_Transaksi) = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                    CInt(rowItemInOut.Tipe_Transaksi) = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                    CInt(rowItemInOut.Tipe_Transaksi) = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                    CInt(rowItemInOut.Tipe_Transaksi) = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                    CInt(rowItemInOut.Tipe_Transaksi) = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        Me.Koreksi4SimpanIn(rowItemInOut.Tanggal_Transaksi, ObjProd, rowItemInOut.Primary_Product, DTA1, rowItemInOut.Qty, _
                        -1, rowItemInOut.Primary_Storage_Area, rowItemInOut.Batch_No, -1, rowItemInOut.Expiry_Date)
                    Else
                        Me.Koreksi4SimpanOut(rowItemInOut.Tanggal_Transaksi, ObjProd, rowItemInOut.Primary_Product, DTA1, rowItemInOut.Qty, _
                        -1, rowItemInOut.Primary_Storage_Area, rowItemInOut.Batch_No, rowItemInOut.Expiry_Date)
                    End If

                    'simpan data stock card
                    If InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                        Qty_In = rowItemInOut.Qty
                        Qty_Out = 0
                    Else
                        Qty_In = 0
                        Qty_Out = rowItemInOut.Qty
                    End If

                    'simpan sc
                    ObjSC.Simpan(rowItemInOut.Primary_Transaksi, JenisSC, rowItemInOut.No_Transaksi, rowItemInOut.Primary_Product, _
                    Qty_In, Qty_Out, DescSC, InOut, rowItemInOut.Tanggal_Transaksi, DTASC, rowItemInOut.Batch_No, _
                    rowItemInOut.Expiry_Date, rowItemInOut.Primary_Storage_Area, rowItemInOut.Remarks)
                Next

                'untuk setiap baris item transfer
                'simpan product qty
                For Each Me.rowTransfer In Me.TBLTransfer
                    Me.ProgressBar1.Value += 1
                    ElapseTime = (Now - StartTime).TotalSeconds
                    Rate = Me.ProgressBar1.Value / ElapseTime
                    RemainingTime = (Me.ProgressBar1.Maximum - Me.ProgressBar1.Value) / Rate
                    RemainingTime = RemainingTime / 60
                    Me.LblProgress.Text = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + " - Location Transfer - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Remaning time: " + ObjCD.Convert_ToString(RemainingTime)

                    Me.Koreksi4SimpanIn(Me.rowTransfer.Tanggal_Transaksi, ObjProd, Me.rowTransfer.Primary_Product, DTA1, Me.rowTransfer.Qty, _
                   -1, Me.rowTransfer.primary_destination_area, Me.rowTransfer.Batch_No, -1, Me.rowTransfer.Expiry_Date)
                    Me.Koreksi4SimpanOut(Me.rowTransfer.Tanggal_Transaksi, ObjProd, Me.rowTransfer.Primary_Product, DTA1, Me.rowTransfer.Qty, _
                   -1, Me.rowTransfer.Primary_Storage_Area, Me.rowTransfer.Batch_No, Me.rowTransfer.Expiry_Date)
                Next

                'untuk setiap baris item stock bundling before
                JenisSC = Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
                InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                For Each Me.rowBundlingBefore In Me.TBLBundlingBefore.Rows
                    Me.ProgressBar1.Value += 1
                    ElapseTime = (Now - StartTime).TotalSeconds
                    Rate = Me.ProgressBar1.Value / ElapseTime
                    RemainingTime = (Me.ProgressBar1.Maximum - Me.ProgressBar1.Value) / Rate
                    RemainingTime = RemainingTime / 60
                    Me.LblProgress.Text = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + " - Bundling Before - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Remaning time: " + ObjCD.Convert_ToString(RemainingTime)


                    'lakukan koreksi di qty detail
                    Me.Koreksi4SimpanOut(Me.rowBundlingBefore.Tanggal_Transaksi, ObjProd, Me.rowBundlingBefore.Primary_Product, DTA1, Me.rowBundlingBefore.Qty, _
                    -1, Me.rowBundlingBefore.Primary_Storage_Area, Me.rowBundlingBefore.Batch_No, Me.rowBundlingBefore.Expiry_Date)

                    'simpan data stock card
                    Qty_In = 0
                    Qty_Out = Me.rowBundlingBefore.Qty
                    With Me.rowBundlingBefore
                        ObjSC.Simpan(Me.rowBundlingBefore.Primary_Transaksi, JenisSC, Me.rowBundlingBefore.No_Transaksi, Me.rowBundlingBefore.Primary_Product, _
                        Qty_In, Qty_Out, "Stock bundling", InOut, .Tanggal_Transaksi, DTASC, .Batch_No, .Expiry_Date, .Primary_Storage_Area, .Remarks)
                    End With
                Next

                'untuk setiap baris item stock bundling after
                'simpan sc
                'simpan product qty
                JenisSC = Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
                InOut = Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                For Each Me.rowBundlingAfter In Me.TBLBundlingAfter.Rows
                    With Me.rowBundlingAfter
                        Me.ProgressBar1.Value += 1
                        ElapseTime = (Now - StartTime).TotalSeconds
                        Rate = Me.ProgressBar1.Value / ElapseTime
                        RemainingTime = (Me.ProgressBar1.Maximum - Me.ProgressBar1.Value) / Rate
                        RemainingTime = RemainingTime / 60
                        Me.LblProgress.Text = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + " - Bundling After - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - Remaning time: " + ObjCD.Convert_ToString(RemainingTime)

                        'lakukan koreksi di qty detail
                        Me.Koreksi4SimpanIn(.Tanggal_Transaksi, ObjProd, .Primary_Product, DTA1, .Qty, _
                        -1, .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)

                        'simpan data stock card
                        Qty_In = .Qty
                        Qty_Out = 0
                        ObjSC.Simpan(.Primary_Transaksi, JenisSC, .No_Transaksi, .Primary_Product, _
                        Qty_In, Qty_Out, "Stock bundling", InOut, .Tanggal_Transaksi, DTASC, .Batch_No, .Expiry_Date, .Primary_Storage_Area, .Remarks)
                    End With
                Next
            Next

            Me.UpdateQtyMin_TblQty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateQtyMin_TblQty()
        Dim DTA As New Core_Module.DseMasterProduct1TableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLDummy As New Core_Module.DseMasterProduct1.tblmaster_productqtyDataTable
        Dim TBL As New Core_Module.DseMasterProduct1.tblmaster_productqtyDataTable
        Dim row, rowDummy As Core_Module.DseMasterProduct1.tblmaster_productqtyRow
        Dim ListPrimaryExclude As New List(Of Double)
        Dim ListPrimaryProduct As New List(Of Double)
        Dim DTA1 As New Core_Module.DseMasterProduct1TableAdapters.DataTable1TableAdapter
        Dim TBL1 As New Core_Module.DseMasterProduct1.DataTable1DataTable
        Dim row1 As Core_Module.DseMasterProduct1.DataTable1Row
        Dim DTA2 As New Core_Module.DseMasterProduct1TableAdapters.tblmaster_productqty1TableAdapter
        Dim TBL2 As New Core_Module.DseMasterProduct1.tblmaster_productqty1DataTable
        Dim row2 As Core_Module.DseMasterProduct1.tblmaster_productqty1Row
        Dim mn As Integer = 0

        Try
            Me.LblProgress.Text = "Preparing data for minus data checking"

            'kumpulkan primary product yang totalnya minus -> ini tidak boleh di repair otomatis
            DTA1.Fill(TBL1)
            For Each row1 In TBL1.Rows
                ListPrimaryExclude.Add(row1.Primary_Product)
            Next

            'kumpulkan primary product yang ada qty minus pada tbl master product qty, kecuali yang ada di tabel di atas
            DTA2.Fill(TBL2)
            For Each row2 In TBL2.Rows
                If ListPrimaryExclude.Contains(row2.Primary_Product) = False Then
                    ListPrimaryProduct.Add(row2.Primary_Product)
                End If
            Next

            'untuk setiap primary product yang akan diperbaiki:
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = ListPrimaryProduct.Count
            For Each PrimaryProduct In ListPrimaryProduct
                Me.ProgressBar1.Value += 1
                Me.LblProgress.Text = "Minus data repair: " + Me.ObjProd.GetProductCode(PrimaryProduct)

                TBLDummy.Rows.Clear()
                TBL.Rows.Clear()
                mn = 0

                '1. kumpulkan tabel product qty
                DTA.Fill(TBLDummy, PrimaryProduct)

                '2. pindahkan ke tabel dummy untuk yang positif saja -> sort berdasarkan expiry, batch, qty
                While mn <= TBLDummy.Rows.Count - 1
                    rowDummy = TBLDummy.Rows(mn)
                    If rowDummy.Qty > 0 Then
                        row = TBL.Newtblmaster_productqtyRow
                        For ij As Integer = 0 To TBL.Columns.Count - 1
                            row.Item(ij) = rowDummy.Item(ij)
                        Next
                        row.Primary = -1
                        TBL.Rows.Add(row)
                        TBLDummy.Rows.Remove(rowDummy)
                    ElseIf rowDummy.Qty = 0 Then
                        'hapus dari database, remove row + jgn dipindah ke tabel baru
                        TBLDummy.Rows.Remove(rowDummy)
                    Else
                        'minus qty, cek baris selanjutnya
                        mn += 1
                    End If
                End While

                '3. kompensasi positif dengan negatif
                mn = 0
                While mn <= TBLDummy.Rows.Count - 1
                    rowDummy = TBLDummy.Rows(mn)
                    row = TBL.Rows(0)
                    If rowDummy.Qty + row.Qty < 0 Then
                        rowDummy.Qty += row.Qty
                        TBL.Rows.Remove(row)
                    Else
                        row.Qty += rowDummy.Qty
                        mn += 1
                        If row.Qty = 0 Then
                            TBL.Rows.Remove(row)
                        End If
                    End If
                End While

                '4. hapus item product qty
                DTA.Delete(PrimaryProduct)

                '5. simpan item product qty yang baru
                For Each row In TBL.Rows
                    With row
                        DTA.Insert(.Primary_Product, .Qty, -1, .Lot_No, .Expiry_Date, .Primary_Storage_Area)
                    End With
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Koreksi4SimpanIn(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                       vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanOut(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
   Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
   Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
   vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
   #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                        vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty * -1, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTipeStockCard(ByVal vTipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut) As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Try
            Select Case vTipeStockInOut
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_BackProcess
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Process
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Purchase
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_SalesReturn
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Process
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                Case Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                    Return Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.Button1.PerformClick()
            MsgBox("Data repaired", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_Cek2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.LblProgress.Text = ""
        'Me.DataGridView1.Rows.Clear()
        'Me.AddRow("K00581721100")
        'Me.AddRow("K0ME81721A00")
        'Me.AddRow("P-IND-01-040")
        'Me.AddRow("P-IND-01-416")
        'Me.AddRow("S-CHE-W0003")
    End Sub

    Private Sub AddRow(ByVal vCodeNo As String)
        Dim rowDGV As DataGridViewRow
        Dim PrimaryProduct As Double = -1

        Try
            Me.DataGridView1.Rows.Add()
            rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
            rowDGV.Cells(Me.ColCodeNo.Index).Value = vCodeNo
            PrimaryProduct = Me.ObjProd.GetPrimary(vCodeNo)
            rowDGV.Cells(Me.ColDesc.Index).Value = Me.ObjProd.GetDescription(PrimaryProduct)
            rowDGV.Cells(Me.ColPrimaryProduct.Index).Value = PrimaryProduct
            rowDGV.Cells(Me.ColQty.Index).Value = 0
            rowDGV.Cells(Me.ColQtySC.Index).Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateCodeNo()
        Dim DTA As New Inventory_Module.DseCekTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New Inventory_Module.DseCek.tblmaster_productDataTable
        Dim row As Inventory_Module.DseCek.tblmaster_productRow

        Try
            'fill data table
            DTA.FillBy(TBL)

            'for each row:
            Me.ProgressBar1.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                'buang '
                'buang space
                'update code no
                Me.ProgressBar1.Value += 1
                Me.LblProgress.Text = "Repairing Code No: " + row.Product_Code
                row.Product_Code = row.Product_Code.Replace("'", "")
                row.Product_Code = row.Product_Code.Trim
                DTA.UpdateQuery(row.Product_Code, row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class