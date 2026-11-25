Option Explicit On
Option Strict Off
Imports MySql.Data.MySqlClient
Imports MySql.Data

Public Class FrmTransaksiDO_Quick_PerBatch
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""
    Public PrimaryTransaction As Double = -1
    Public PrimaryCode As Double = -1
    Public IsBatchControlled As Integer = 0
    Public IsExpiryControlled As Integer = 0
    Public QtySOB As Decimal = 0
    Public MultiUnit As String = ""
    Public PrimaryUnit As Double = -1
    Public PrimaryLocation As Double = -1
    Public PrimaryWarehouse As Double = -1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""
    Dim mDeleteIndex As Integer = -1
    Dim LangsungTutup As Boolean = False
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim ObjWR As New Core_Module.clsMasterWarehouse
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim ObjStockInOutItem As New Inventory_Module.clsTransaksiStockInOutItem
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Public Tbl As New DseTransaksiQuick_DO.TblSimpan1DataTable
    Public TblStock As New DseTransaksiQuick_DO2.stock_detail_viewDataTable
    Public TblStockTemp As New DseTransaksiQuick_DO.TempStockDetailDataTable
    Public TblPendingDO As New DseTransaksiQuick_DO2.Pending_DO_ViewDataTable
    Dim rowStock As DseTransaksiQuick_DO2.stock_detail_viewRow
    Dim rowStockTemp As DseTransaksiQuick_DO.TempStockDetailRow
    Dim rowPendingDO As DseTransaksiQuick_DO2.Pending_DO_ViewRow
    Dim Row As DseTransaksiQuick_DO.TblSimpan1Row

    Private Sub FrmTransaksiQuick_StockInPurchase_PerBatch_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungTutup = False Then
                If MsgBox("Close dan discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuick_StockInPurchase_PerBatch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Dgv.AllowUserToAddRows = False
            Me.Dgv.AllowUserToDeleteRows = False

            If IsExpiryControlled = 0 Then
                Me.ColExpiryDate.Visible = False
            Else
                Me.ColExpiryDate.Visible = True
            End If
            Me.Text = "Option Batch No & Expiry Date"

            Me.IsiDGV()
            If Me.Dgv.Rows.Count = 0 Then
                MsgBox("There is no stock for this product.", MsgBoxStyle.Information, "Information")
                Me.LangsungTutup = True
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            If Me.ApakahDataBenar() = True Then
                If MsgBox("Add this data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.LangsungTutup = True
                    Me.UpdateTbl()
                    Me.Close()
                End If
            Else
                Me.LangsungTutup = True
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim JmlQty As Decimal = 0
        Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlDGVAll As Decimal
        Dim decimalPoint As Integer = 0
        Try
            If Me.Dgv.Rows.Count = 0 Then
                MsgBox("There is no stock for this product.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If CDec(Me.Dgv.Item(Me.ColQtyIn.Index, 0).Value) = 0 Then
                MsgBox("Please enter qty in minimum 1 item.", MsgBoxStyle.Information, "Information")
                Exit Function
            End If

            For Each row As DataGridViewRow In Me.Dgv.Rows
                JmlQty += CDec(row.Cells(Me.ColQtyIn.Index).Value)
            Next


            'If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
            '    JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)

            '    ''@rosadinnn
            '    ''@MUaT
            '    Dim ToleranceUp As Decimal = objPO.GetToleranceUp(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
            '    JmlPOB += ToleranceUp

            'Else
            '    JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
            'End If

            JmlPOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.PrimaryTransaction, Me.PrimaryCode)
            Dim ToleranceUp As Decimal = ObjSO.GetToleranceUp(Me.PrimaryTransaction, Me.PrimaryCode)
            JmlPOB += ToleranceUp

            JmlObjLama = 0
            Dim SecQty As Decimal = 0

            Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
            Me.ObjProduct.GetProductProperties(Me.PrimaryCode, LocationCont, BatchNoCont, _
            ExpiryDateCont, UseSecondary) ', decimalPoint)
            JmlDGVAll = JmlQty

            ''@rosadinnn
            ''@MUaT
            Dim MaxCanInput As Decimal = 0
            'If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
            '    JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGVAll - CDec(Me.TxtQtyperBatch.Text)
            '    MaxCanInput = JmlPOB + JmlObjLama - JmlDGVAll
            'Else
            '    JmlPOBBaru = JmlPOB - JmlDGVAll - CDec(Me.TxtQtyperBatch.Text)
            '    MaxCanInput = JmlPOB - JmlDGVAll
            'End If

            JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGVAll
            MaxCanInput = JmlPOB

            If JmlPOBBaru < 0 Then
                MsgBox("Maximum total qty you can enter is " + Me.ObjCD.DigitGrouping(CStr(MaxCanInput), Me.Dgv.Item(Me.ColDecimal.Index, 0).Value) + _
                        " " + CStr(Me.Dgv.Item(Me.ColUnit.Index, 0).Value) + ".", MsgBoxStyle.Information, "Information")

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Dgv_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles Dgv.CellBeginEdit
        Try
            'If e.ColumnIndex = Me.ColBatchNo.Index Then
            '    If IsBatchControlled = 0 Then
            '        e.Cancel = True
            '    Else
            '        Me.Buffer = Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value
            '    End If
            'ElseIf e.ColumnIndex = Me.ColExpiryDate.Index Then
            '    If IsExpiryControlled = 0 Then
            '        e.Cancel = True
            '    Else
            '        Me.Buffer = Me.Dgv.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
            '    End If
            'ElseIf e.ColumnIndex = Me.ColQtyIn.Index Then
            If e.ColumnIndex = Me.ColQtyIn.Index Then
                Me.Buffer = Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal eRow As Integer, ByVal vBatchNo As String) As Integer

        Try
            For Each row As DataGridViewRow In Me.Dgv.Rows
                If eRow <> row.Index Then
                    If vBatchNo = row.Cells(Me.ColBatchNo.Index).Value Then
                        Return 1
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Dgv_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Dgv.CellEndEdit
        Dim rowdgv As DataGridViewRow = Me.Dgv.Rows(e.RowIndex)
        Try
            If e.ColumnIndex = Me.ColQtyIn.Index Then
                If Me.IsBatchControlled = 1 Then
                    If String.IsNullOrEmpty(Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value) = True Then
                        MsgBox("Please enter batch no first.", MsgBoxStyle.Information, "Information")
                        Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value = Me.Buffer

                        If Me.Dgv.Rows.Count > 1 Then
                            'jika data baru, hapuskan baris ybs
                            Me.mDeleteIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If

                    If IsNumeric(Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for data.", MsgBoxStyle.Information, "Information")
                        Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'Cek Qty Available
                    If rowdgv.Cells(Me.ColQtyIn.Index).Value > rowdgv.Cells(Me.ColQtyAvail.Index).Value Then
                        MsgBox("Maximum qty can input : " & rowdgv.Cells(Me.ColQtyAvail.Index).Value & ".", MsgBoxStyle.Information, "Information")
                        Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'Cek Qty Total
                    If Me.cekQtyDgvVSSOB(rowdgv.Cells(Me.ColQtyIn.Index).Value, e.RowIndex) = False Then
                        Me.Dgv.Item(Me.ColQtyIn.Index, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        simpanStockTemp(rowdgv, rowdgv.Cells(Me.ColQtyIn.Index).Value)
                    End If

                    Me.UpdateQtyIn()
                End If

                'ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                '    Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value = Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value.ToString.Trim.ToUpper
                '    If Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value.ToString.Length > 250 Then
                '        Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value = Me.Buffer
                '        Exit Sub
                '    End If

                '    If Me.GetIndex(e.RowIndex, Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value) = 1 Then
                '        MsgBox("Duplication batch no, please change batch no", MsgBoxStyle.Information, "Information")
                '        Me.Dgv.Item(Me.ColBatchNo.Index, e.RowIndex).Value = Me.Buffer
                '        Exit Sub
                '    End If

                'ElseIf e.ColumnIndex = Me.ColExpiryDate.Index Then
                '    If IsDate(Me.Dgv.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                '        MsgBox("Please specify a valid date.", MsgBoxStyle.Information, "Information")
                '        Me.Dgv.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                '        If e.RowIndex = Me.Dgv.Rows.Count - 2 And e.RowIndex <> 0 Then
                '            'jika data baru, hapuskan baris ybs
                '            Me.mDeleteIndex = e.RowIndex
                '            Me.BackgroundWorker1.RunWorkerAsync()
                '        End If
                '        Exit Sub
                '    End If
                '    Me.Dgv.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.Dgv.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Dgv_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles Dgv.RowsAdded
        Dim row As DataGridViewRow
        Try
            row = Me.Dgv.Rows(Me.Dgv.Rows.Count - 1)
            With row
                .Cells(Me.ColNo.Index).Value = Me.Dgv.Rows.Count
                .Cells(Me.ColBatchNo.Index).Value = ""
                .Cells(Me.ColExpiryDate.Index).Value = Format(Now, "dd MMM yyyy")
                .Cells(Me.ColQtyIn.Index).Value = "0"
                .Cells(Me.ColUnit.Index).Value = Me.MultiUnit
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(10)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Dgv.Rows.RemoveAt(Me.mDeleteIndex)
            Me.mDeleteIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function UpdateTbl() As Boolean
        Dim RowSimpan As DseTransaksiQuick_DO.TblSimpan1Row = Me.Tbl.NewTblSimpan1Row
        Try
            Me.Tbl.DefaultView.RowFilter = "[Primary Product]=" & PrimaryCode & " AND [Primary Transaction]=" & Me.PrimaryTransaction
            For i As Integer = Me.Tbl.DefaultView.Count - 1 To 0 Step -1
                Me.Row = Me.Tbl.DefaultView(i).Row
                RowSimpan.ItemArray = CType(Me.Row.ItemArray.Clone(), Object()) 'untuk simpan sementara row
                Me.Tbl.Rows.Remove(Me.Row)
            Next

            For Each rowDGV As DataGridViewRow In Me.Dgv.Rows
                If CDec(rowDGV.Cells(Me.ColQtyIn.Index).Value) > 0 Then
                    Me.Tbl.DefaultView.RowFilter = "[Primary Product]=" & PrimaryCode & " AND [Batch No]='" & rowDGV.Cells(Me.ColBatchNo.Index).Value & "'" & _
                    " AND [Primary Transaction]=" & Me.PrimaryTransaction
                    If Me.Tbl.DefaultView.Count = 0 Then
                        Me.Row = Me.Tbl.NewTblSimpan1Row

                        Me.Row.Primary_Transaction = Me.PrimaryTransaction
                        Me.Row.Transaction_No = RowSimpan.Transaction_No
                        Me.Row.Primary_Customer = RowSimpan.Primary_Customer
                        Me.Row.Primary_Currency = RowSimpan.Primary_Currency
                        Me.Row.Primary_Product = Me.PrimaryCode
                        Me.Row.Primary_Unit = Me.PrimaryUnit
                        Me.Row.Primary_Multi_Unit = RowSimpan.Primary_Multi_Unit
                        Me.Row.Primary_Location = Me.PrimaryLocation
                        Me.Row.Primary_Warehouse = Me.PrimaryWarehouse
                        Me.Row.Customer_Code = RowSimpan.Customer_Code
                        Me.Row.Customer_Name = RowSimpan.Customer_Name
                        Me.Row.Unit_Code_Default = RowSimpan.Unit_Code_Default
                        Me.Row.Unit_Code_Multi = RowSimpan.Unit_Code_Multi
                        Me.Row.Product_Code = RowSimpan.Product_Code
                        Me.Row.Product_Desc = RowSimpan.Product_Desc
                        Me.Row.SOB_Multi = RowSimpan.SOB_Multi
                        Me.Row.Tolerance_Up = RowSimpan.Tolerance_Up
                        Me.Row.SOB = RowSimpan.SOB
                        Me.Row.Qty_In = rowDGV.Cells(Me.ColQtyIn.Index).Value * (RowSimpan.Ratio_Default / RowSimpan.Ratio_References)  'isi qty default
                        Me.Row.Qty_In_Multi = rowDGV.Cells(Me.ColQtyIn.Index).Value
                        Me.Row.Ratio_Default = RowSimpan.Ratio_Default
                        Me.Row.Ratio_References = RowSimpan.Ratio_References
                        Me.Row.Location_Code = RowSimpan.Location_Code
                        Me.Row.Location_Desc = RowSimpan.Location_Desc
                        Me.Row.Warehouse_Code = RowSimpan.Warehouse_Code
                        Me.Row.Warehosue_Desc = RowSimpan.Warehosue_Desc
                        Me.Row.Identifier_CustAddress = RowSimpan.Identifier_CustAddress
                        Me.Row.Tax = RowSimpan.Tax
                        Me.Row.FakturPajak = RowSimpan.FakturPajak
                        Me.Row.identifier_address_ship_to = RowSimpan.identifier_address_ship_to
                        Me.Row.Primary_Customer_ship_to = RowSimpan.Primary_Customer_ship_to
                        Me.Row.identifier_address_customer_ship_to = RowSimpan.identifier_address_customer_ship_to
                        Me.Row.Batch_No = rowDGV.Cells(Me.ColBatchNo.Index).Value
                        If Me.IsExpiryControlled = 1 Then
                            Me.Row.Expiry_Date = rowDGV.Cells(Me.ColExpiryDate.Index).Value
                        End If
                        Me.Row.Remarks = RowSimpan.Remarks
                        Me.Row.Transaction_Date = RowSimpan.Transaction_Date

                        Me.Tbl.AddTblSimpan1Row(Me.Row)
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiDGV()
        Try
            Me.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.Dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.Dgv.AllowUserToAddRows = False
            Me.Tbl.DefaultView.RowFilter = "[Primary Product]=" & PrimaryCode & " AND [Primary Transaction]=" & PrimaryTransaction
            For i As Integer = 0 To Me.Tbl.DefaultView.Count - 1
                Me.Row = Me.Tbl.DefaultView(i).Row
                isiDataBatchNo(Me.PrimaryCode, Me.PrimaryTransaction)
                'rowDGV = Me.Dgv.Rows(Me.Dgv.Rows.Add)
                'With rowDGV
                '    .Cells(Me.ColNo.Index).Value = ""
                '    .Cells(Me.ColBatchNo.Index).Value = Me.Row.Batch_No
                '    .Cells(Me.ColExpiryDate.Index).Value = Format(Me.Row.Expiry_Date, "dd MMM yyyy")
                '    .Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(Me.Row.Qty_In_Multi, Me.Row.Decimal_Multi)
                '    .Cells(Me.ColUnit.Index).Value = Me.Row.Unit_Code_Multi
                '    .Cells(Me.ColDecimal.Index).Value = Me.Row.Decimal_Multi
                'End With
            Next
            Me.ObjCD.SetNumber(Me.Dgv)
            'Me.Dgv.AllowUserToAddRows = True
            Me.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Dgv_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles Dgv.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.Dgv)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub isiDataBatchNo(ByVal vPrimaryCode As Double, ByVal vPrimarySo As Double)
        Dim rowDGV As DataGridViewRow
        'Dim DTA As New Inventory_Module.dseLaporanStockDetailTableAdapters.stock_detail_viewTableAdapter
        'Dim TBLStock As New Inventory_Module.dseLaporanStockDetail.stock_detail_viewDataTable
        Dim qtyAvail As Decimal = 0
        Dim qtyTemp As Decimal = 0
        Dim qtyPendingDO As Decimal = 0
        Try
            Me.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.Dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.Dgv.AllowUserToAddRows = False

            'DTA.FillBy_PrimaryProduct(TblStock, vPrimaryCode)
            Me.TblStock.DefaultView.RowFilter = "[Primary Product]= " & Me.PrimaryCode & " AND [Primary Unit]= " & Me.PrimaryUnit & _
            " AND [Primary Storage Area]= " & Me.PrimaryLocation & " AND [Primary Warehouse]= " & Me.PrimaryWarehouse
            If Me.TblStock.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.TblStock.DefaultView.Count - 1
                    rowStock = Me.TblStock.DefaultView(i).Row
                    rowDGV = Me.Dgv.Rows(Me.Dgv.Rows.Add)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColBatchNo.Index).Value = rowStock.Lot_No
                        .Cells(Me.ColExpiryDate.Index).Value = Format(rowStock.Expiry_Date, "dd MMM yyyy")
                        .Cells(Me.ColQtyAvail.Index).Value = Me.ObjCD.DigitGrouping(Me.rowStock.Qty, Me.Row.Decimal_Multi)
                        'cek temp
                        qtyTemp = Me.getTempQty(rowStock.Primary_Product, rowStock.Primary_Unit, rowStock.Primary_Storage_Area, rowStock.Primary_Warehouse, _
                        Me.PrimaryTransaction, rowStock.Lot_No, rowStock.Expiry_Date)
                        If qtyTemp > 0 Then
                            .Cells(Me.ColQtyAvail.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColQtyAvail.Index).Value) - qtyTemp, Me.Row.Decimal_Multi)
                        End If

                        'cek Pending DO
                        qtyPendingDO = Me.getPendingDOQty(rowStock.Primary_Product, rowStock.Primary_Unit, rowStock.Primary_Storage_Area, rowStock.Primary_Warehouse, _
                        rowStock.Lot_No, rowStock.Expiry_Date)
                        If qtyPendingDO > 0 Then
                            .Cells(Me.ColQtyAvail.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColQtyAvail.Index).Value) - qtyPendingDO, Me.Row.Decimal_Multi)
                        End If

                        'cek Qty Udah diisi apa belum
                        qtyTemp = Me.getTempQtyProduct(rowStock.Primary_Product, rowStock.Primary_Unit, rowStock.Primary_Storage_Area, rowStock.Primary_Warehouse, _
                        Me.PrimaryTransaction, rowStock.Lot_No, rowStock.Expiry_Date)
                        If qtyTemp > 0 Then
                            .Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(qtyTemp, Me.Row.Decimal_Multi)
                        End If

                        '.Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(Me.Row.Qty_In_Multi, Me.Row.Decimal_Multi)
                        .Cells(Me.ColUnit.Index).Value = Me.Row.Unit_Code_Multi
                        .Cells(Me.ColDecimal.Index).Value = Me.Row.Decimal_Multi
                    End With
                    If rowDGV.Cells(Me.ColQtyAvail.Index).Value <= 0 Then
                        Me.Dgv.Rows.Remove(rowDGV)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.Dgv.AllowUserToAddRows = True
        End Try

    End Sub

    Private Function getAvailableQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByRef vPrimaryLocation As Double, _
  ByRef vPrimaryWarehouse As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyAvailable As Decimal = 0
        Try

            Me.TblStock.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
            " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblStock.DefaultView.Count > 0 Then
                Me.rowStock = Me.TblStock.DefaultView(0).Row
                vQtyAvailable = Me.rowStock.Qty
            Else
                Return 0
            End If

            Return vQtyAvailable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TblStock.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getPendingDOQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByRef vPrimaryLocation As Double, _
 ByRef vPrimaryWarehouse As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyPendingDO As Decimal = 0
        Try

            Me.TblPendingDO.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & _
            " AND [Batch No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblPendingDO.DefaultView.Count > 0 Then
                For Each rowdv As DataRowView In Me.TblPendingDO.DefaultView
                    Me.RowPendingDO = rowdv.Row
                    vQtyPendingDO += Me.RowPendingDO.Total_Qty
                Next
            Else
                Return 0
            End If

            Return vQtyPendingDO
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TblPendingDO.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getTempQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryLocation As Double, _
    ByVal vPrimaryWarehouse As Double, ByVal vPrimarySO As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyTemp As Decimal = 0
        Try
            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & " AND [Primary SO]<> " & vPrimarySO & _
            " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then
                Me.rowStockTemp = Me.TblStockTemp.DefaultView(0).Row
                vQtyTemp = Me.rowStockTemp.Qty
            Else
                Return 0
            End If

            Return vQtyTemp
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TblStockTemp.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Function getTempQtyProduct(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double, ByVal vPrimaryLocation As Double, _
    ByVal vPrimaryWarehouse As Double, ByVal vPrimarySO As Double, Optional ByVal vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Decimal
        Dim vQtyTemp As Decimal = 0
        Try
            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & vPrimaryProduct & " AND [Primary Unit]= " & vPrimaryUnit & _
            " AND [Primary Storage Area]= " & vPrimaryLocation & " AND [Primary Warehouse]= " & vPrimaryWarehouse & " AND [Primary SO]= " & vPrimarySO & _
            " AND [Lot No]= '" & vBatchNo & "' AND [Expiry Date]= '" & vExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then
                Me.rowStockTemp = Me.TblStockTemp.DefaultView(0).Row
                vQtyTemp = Me.rowStockTemp.Qty
            Else
                Return 0
            End If

            Return vQtyTemp
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.TblStockTemp.DefaultView.RowFilter = ""
        End Try
    End Function

    Private Sub simpanStockTemp(ByVal rowDGV As DataGridViewRow, ByVal vQty As Decimal)
        Dim BatchNo As String = rowDGV.Cells(Me.ColBatchNo.Index).Value
        Dim ExpiryDate As Date = rowDGV.Cells(Me.ColExpiryDate.Index).Value
        Dim Qty As Decimal = rowDGV.Cells(Me.ColQtyAvail.Index).Value
        Try

            Me.TblStockTemp.DefaultView.RowFilter = "[Primary Product]= " & Me.PrimaryCode & " AND [Primary Unit]= " & Me.PrimaryUnit & _
                        " AND [Primary Storage Area]= " & Me.PrimaryLocation & " AND [Primary Warehouse]= " & Me.PrimaryWarehouse & " AND [Primary SO]= " & Me.PrimaryTransaction & _
                        " AND [Lot No]= '" & BatchNo & "' AND [Expiry Date]= '" & ExpiryDate.Date & "'"
            If Me.TblStockTemp.DefaultView.Count > 0 Then 'hanya update
                Me.rowStockTemp = Me.TblStockTemp.DefaultView(0).Row
                rowStockTemp.Qty = vQty
            Else
                Me.rowStockTemp = Me.TblStockTemp.NewTempStockDetailRow
                With rowDGV
                    rowStockTemp.Primary_Product = Me.PrimaryCode
                    rowStockTemp.Primary_Storage_Area = PrimaryLocation
                    rowStockTemp.Primary_SO = Me.PrimaryTransaction
                    rowStockTemp.Primary_Unit = PrimaryUnit
                    rowStockTemp.Primary_Warehouse = PrimaryWarehouse
                    rowStockTemp.Lot_No = BatchNo
                    rowStockTemp.Expiry_Date = ExpiryDate
                    rowStockTemp.Qty = vQty
                    rowStockTemp.Location = Me.ObjLoc.GetStorageAreaCode(Me.PrimaryLocation)
                    rowStockTemp.Warehouse = Me.ObjWR.GetWarehouseCode(Me.PrimaryWarehouse)
                    rowStockTemp._Decimal = rowDGV.Cells(Me.ColDecimal.Index).Value
                End With
                Me.TblStockTemp.Rows.Add(Me.rowStockTemp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function cekQtyDgvVSSOB(ByVal vQtyInput As Decimal, ByVal vRowIndex As Integer) As Boolean
        Dim rowdgv As DataGridViewRow
        Dim qtyDgv As Decimal = 0
        Dim qtyMaxInput As Decimal = 0
        Try
            For Each rowdgv In Me.Dgv.Rows
                If rowdgv.Index <> vRowIndex Then
                    qtyDgv = qtyDgv + CDec(rowdgv.Cells(Me.ColQtyIn.Index).Value)
                End If
            Next
            qtyMaxInput = Me.TxtQtyTotal.Text - qtyDgv
            If vQtyInput <= qtyMaxInput Then
                Return True
            Else
                MsgBox("Maximum qty can input : " & qtyMaxInput, MsgBoxStyle.Information, "Information")
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Function

    Private Sub UpdateQtyIn()
        Dim rowdgv As DataGridViewRow
        Dim QtyDgv As Decimal = 0
        Dim defprec As Integer = 0
        Try
            For Each rowdgv In Me.Dgv.Rows
                QtyDgv = QtyDgv + CDec(rowdgv.Cells(Me.ColQtyIn.Index).Value)
                If defprec < rowdgv.Cells(Me.ColDecimal.Index).Value Then defprec = rowdgv.Cells(Me.ColDecimal.Index).Value
            Next
            Me.txtQtyTotalIn.Text = Me.ObjCD.DigitGrouping(QtyDgv, defprec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class