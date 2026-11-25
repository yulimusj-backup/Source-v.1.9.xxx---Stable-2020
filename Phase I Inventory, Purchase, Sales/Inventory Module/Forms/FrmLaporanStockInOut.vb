Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmLaporanStockInOut
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Private Scope_Index As EnumScope
    Private LevelGroup As EnumLevel
    Private LevelIndex As EnumLevelIndex
    Private ShowSubTotal As EnumShowSubTotal

    Private Buffer As String = ""
    Private objGroup As New Core_Module.ClsMasterProductGroup
    Private ObjCategory As New Core_Module.ClsMasterCategory
    Private ObjProduct As New Core_Module.clsMasterProduct
    Private ObjLocation As New Core_Module.clsMasterStorageArea
    Private ObjCust As New Core_Module.clsMasterCustomer
    Private ObjSupp As New Core_Module.clsMasterSupplier
    Private ObjPO As New ClsTransaksiPO
    Private ObjSO As New ClsTransaksiSalesOrder
    Private ObjCD As New Core_Module.ClsCekData
    Private ObjWarehouse As New Core_Module.clsMasterWarehouse

    Private EraseRowindexLocation As Integer = 0
    Private EraseRowIndexProduct As Integer = 0
    Private EraseRowIndexBatchNo As Integer = 0
    Private EraseRowIndexSuppCust As Integer = 0
    Private TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut
    Private TBL As New dseTransaksiStockInOut.tbl_setting_adjustmentDataTable
    Dim ObjLicense As New Core_Module.ClsLicense

#Region "Enum saja"
    Public Enum EnumScope
        Goods_Only
        Service_Only
        Both
    End Enum

    Public Enum EnumLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum

    Public Enum EnumLevelIndex As Integer
        Index01
        Index02
        Index03
    End Enum

    Public Enum EnumShowSubTotal
        No
        Yes
    End Enum

    Public Enum EnumPilih
        No
        Yes
    End Enum

    Public Enum JenisProductGroup As Integer
        Product_Group
        Brand
        Sub_Brand
        Category
        Product
    End Enum
#End Region

#Region "Atur Tampilan"
    Public Sub IsiClbSetting(ByVal vTBL As dseTransaksiStockInOut.tbl_setting_adjustmentDataTable, ByVal vClb As CheckedListBox)
        Dim dta As New dseTransaksiStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim row As dseTransaksiStockInOut.tbl_setting_adjustmentRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(vTBL)

            'set tampilan
            For Each row In vTBL.Rows
                row.Tampilan = row.Code + " - " + row.Desc
            Next

            'insert baris kedua - stock taking
            row = vTBL.Newtbl_setting_adjustmentRow
            row.Code = "Stock Taking"
            row.Desc = ""
            row.Primary = -2
            row.PrimaryGain = -1
            row.PrimaryLoss = -1
            row.Remarks = ""
            row.Tampilan = "Stock Taking Adjustment"
            vTBL.Rows.InsertAt(row, 0)

            'insert baris pertama - no type
            row = vTBL.Newtbl_setting_adjustmentRow
            row.Code = "No Tipe"
            row.Desc = ""
            row.Primary = -1
            row.PrimaryGain = -1
            row.PrimaryLoss = -1
            row.Remarks = ""
            row.Tampilan = "No Adjustment Type"
            vTBL.Rows.InsertAt(row, 0)

            vClb.DataSource = vTBL.DefaultView
            vClb.DisplayMember = "Tampilan"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                Me.clbAdjustmentType.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                Me.clbAdjustmentType.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanStockDetail2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboPilihGroupLevel.SelectedIndex = 4
            Me.CboScope.SelectedIndex = 0
            Me.DtpMinExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Value = Now
            Me.DtpMaxExpiryDate.Enabled = True
            Me.DtpMinExpiryDate.Enabled = True
            Me.ChkShowAllItemsInTrans.Checked = False
            Me.ChkShowSubTotal.Checked = False
            Me.DgvLocation.AllowUserToAddRows = True
            Me.DgvGroup.AllowUserToAddRows = True
            Me.DgvBatchNo.AllowUserToAddRows = True
            Me.dgvCustomer.AllowUserToAddRows = True
            Me.DgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvBatchNo.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboFormat.SelectedIndex = 0
            Me.CboInPurchase.SelectedIndex = 0
            Me.LblPONumber.Text = "PO Number"
            Me.TabPage4.Text = "Supplier Data"
            Me.dgvCustomer.BringToFront()
            Me.dgvCustomer.Rows.Clear()
            'Isi Clb Adjustment Type
            Me.IsiClbSetting(Me.TBL, Me.clbAdjustmentType)
            For i As Integer = 0 To Me.clbAdjustmentType.CheckedItems.Count - 1
                Me.clbAdjustmentType.SetItemChecked(i, False)
            Next
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            Me.ChkFilterByDate.Checked = True
            Me.ChkShowAOPNo.Checked = True
            Me.ChkFilterByWareHouse.Checked = True
            Me.TabPage1.Text = "Warehouse"

            Me.CboScope.SelectedIndex = 2
            Me.SplitContainer1.SplitterDistance = Me.ChkShowSubTotal.Top + Me.ChkShowSubTotal.Height + 5
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboFormat.SelectedIndexChanged
        Try
            If Me.CboFormat.SelectedIndex = 0 Then
                'In Purchase
                Me.CboInPurchase.Enabled = True
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "PO Number"
                Me.TxtPONumber.Enabled = True
                Me.btnFindPONumber.Enabled = True
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Supplier Data"
                Me.dgvCustomer.BringToFront()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.CheckedItems.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, False)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                If Me.CboInPurchase.SelectedIndex = 0 Or Me.CboInPurchase.SelectedIndex = 2 Then
                    Me.ChkShowAOPNo.Checked = True
                    Me.ChkShowAOPNo.Enabled = True
                Else
                    Me.ChkShowAOPNo.Checked = False
                    Me.ChkShowAOPNo.Enabled = False
                End If
                Me.ChkFilterByDate.Checked = True
            ElseIf Me.CboFormat.SelectedIndex = 1 Then
                'In Sales Return
                Me.CboInPurchase.Enabled = False
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "SO Number"
                Me.TxtPONumber.Enabled = True
                Me.btnFindPONumber.Enabled = True
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Customer Data"
                Me.dgvCustomer.BringToFront()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.CheckedItems.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, False)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            ElseIf Me.CboFormat.SelectedIndex = 2 Then
                'In Adjust
                Me.CboInPurchase.Enabled = False
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "PO Number"
                Me.TxtPONumber.Enabled = False
                Me.btnFindPONumber.Enabled = False
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Adjustment Type"
                Me.dgvCustomer.SendToBack()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, True)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            ElseIf Me.CboFormat.SelectedIndex = 3 Then
                'Out Sales
                Me.CboInPurchase.Enabled = False
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "SO Number"
                Me.TxtPONumber.Enabled = True
                Me.btnFindPONumber.Enabled = True
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Customer Data"
                Me.dgvCustomer.BringToFront()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.CheckedItems.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, False)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            ElseIf Me.CboFormat.SelectedIndex = 4 Then
                'Out Purchase Return 
                Me.CboInPurchase.Enabled = False
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "PO Number"
                Me.TxtPONumber.Enabled = True
                Me.btnFindPONumber.Enabled = True
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Supplier Data"
                Me.dgvCustomer.BringToFront()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.CheckedItems.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, False)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            ElseIf Me.CboFormat.SelectedIndex = 5 Then
                'Out Adjust
                Me.CboInPurchase.Enabled = False
                Me.CboInPurchase.SelectedIndex = 0
                Me.LblPONumber.Text = "PO Number"
                Me.TxtPONumber.Enabled = False
                Me.btnFindPONumber.Enabled = False
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Clear()
                Me.TabPage4.Text = "Adjustment Type"
                Me.dgvCustomer.SendToBack()
                Me.dgvCustomer.Rows.Clear()
                For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                    Me.clbAdjustmentType.SetItemChecked(i, True)
                Next
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPilihGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPilihGroupLevel.SelectedIndexChanged
        Try
            If Me.CboPilihGroupLevel.SelectedIndex = 0 Then
                LevelGroup = EnumLevel.ProductGroup
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 1 Then
                LevelGroup = EnumLevel.Brand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 2 Then
                LevelGroup = EnumLevel.SubBrand
                LevelIndex = EnumLevelIndex.Index01
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 3 Then
                LevelGroup = EnumLevel.Category
                LevelIndex = EnumLevelIndex.Index02
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = 4 Then
                LevelIndex = EnumLevelIndex.Index03
                Me.DgvGroup.Rows.Clear()
                Me.DgvGroup.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboScope_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboScope.SelectedIndexChanged
        Try
            If Me.CboScope.SelectedIndex = 0 Then
                Me.Scope_Index = EnumScope.Goods_Only
            ElseIf Me.CboScope.SelectedIndex = 1 Then
                Me.Scope_Index = EnumScope.Service_Only
            ElseIf Me.CboScope.SelectedIndex = 2 Then
                Me.Scope_Index = EnumScope.Both
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkShowSubTotal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkShowSubTotal.CheckedChanged
        Try
            If Me.ChkShowSubTotal.Checked = False Then
                Me.ShowSubTotal = EnumShowSubTotal.No
            ElseIf Me.ChkShowSubTotal.Checked = True Then
                Me.ShowSubTotal = EnumShowSubTotal.Yes
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
            'System.Environment.GetFolderPath(SpecialFolder.CommonApplicationData)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpMaxExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMaxExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpMinExpiryDate.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpMinExpiryDate.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If


            If Me.DtpMinExpiryDate.Value > Me.DtpMaxExpiryDate.Value Then
                MsgBox("Period is not valid, please check again.", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Exit Sub
            End If
            Me.Cetak(Me.DtpMaxExpiryDate.Value.Date, Me.DtpMinExpiryDate.Value.Date)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Location"
    Private Sub DgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvLocation.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColLocation_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvLocation.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryLocation As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColLocation_Code.Index Then
                        Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ChkFilterByWareHouse.Checked = False Then
                            If Me.ObjLocation.ApakahSudahAda(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                PrimaryLocation = Me.ObjLocation.GetPrimary(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemLocation(PrimaryLocation, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = PrimaryLocation
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(PrimaryLocation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                            Else

                                Dim ij As New Core_Module.FrmCariMasterStorageArea
                                ij.txtCari.Text = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryStorageArea <> -1 Then

                                    index = Me.GetIndexItemLocation(ij.PrimaryStorageArea, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                        Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowindexLocation = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                                            Exit Sub
                                        Else
                                            Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                            Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                            Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                        End If
                                    End If
                                ElseIf ij.PrimaryStorageArea = -1 Then
                                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                        ElseIf Me.ChkFilterByWareHouse.Checked = True Then
                            If Me.ObjWareHouse.ApakahSudahAda(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                PrimaryLocation = Me.ObjWareHouse.GetPrimary(Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemLocation(PrimaryLocation, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = PrimaryLocation
                                    Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWareHouse.GetWarehouseCode(PrimaryLocation)
                                    Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWareHouse.GetDescription(PrimaryLocation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                            Else

                                Dim ij As New Core_Module.frmCariMasterWarehouse
                                ij.txtCari.Text = Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryWarehouse <> -1 Then

                                    index = Me.GetIndexItemLocation(ij.PrimaryWarehouse, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                        Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWareHouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                        Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWareHouse.GetDescription(ij.PrimaryWarehouse)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowindexLocation = e.RowIndex
                                                Me.BackgroundWorker1.RunWorkerAsync()
                                            Else
                                                Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                                            Exit Sub
                                        Else
                                            Me.DgvLocation.Item(Me.ColLocation_Primary.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                            Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value = Me.ObjWarehouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                            Me.DgvLocation.Item(Me.ColLocation_Desc.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(ij.PrimaryWarehouse)
                                        End If
                                    End If
                                ElseIf ij.PrimaryWarehouse = -1 Then
                                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowindexLocation = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvLocation)
                        End If
                    End If
                Else
                    If Me.DgvLocation.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                    If Me.DgvLocation.Item(Me.ColLocation_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowindexLocation = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                End If
                'Me.dgvLocation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvLocation.Columns(Me.ColLocation_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvLocation.Columns(Me.ColLocation_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemLocation(ByVal vPrimary As Double, _
          ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvLocation.Rows.Count - 2
                RowDGV = Me.DgvLocation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColLocation_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvLocation.Rows.RemoveAt(EraseRowindexLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Batch No"
    Private Sub DgvBatchNo_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvBatchNo.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                Me.Buffer = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBatchNo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvBatchNo.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvBatchNo.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim BatchNo As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColBatch_BatchNo.Index Then
                        'set ke huruf besar dan hapus spasi
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim

                        index = Me.GetIndexItemBatchNo(Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value, RowIdx)
                        If index >= 0 Then
                            If e.RowIndex <> index Then
                                MsgBox("Duplication Batch No, please input another Batch No.", MsgBoxStyle.Information, "Information")
                                Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                    Me.EraseRowIndexBatchNo = e.RowIndex
                                    Me.BackgroundWorker3.RunWorkerAsync()
                                Else
                                    Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                Else

                    If Me.DgvBatchNo.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If

                    If Me.DgvBatchNo.Item(Me.ColBatch_BatchNo.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexBatchNo = e.RowIndex
                            Me.BackgroundWorker3.RunWorkerAsync()
                        Else
                            Me.DgvBatchNo.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemBatchNo(ByVal vBatchNo As String, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvBatchNo.Rows.Count - 2
                RowDGV = Me.DgvBatchNo.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value = vBatchNo Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DgvBatchNo.Rows.RemoveAt(EraseRowIndexBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Product Group"
    Private Sub DgvProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvGroup.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    '''simpan nilai awal di buffer
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.Buffer = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim _Group As Boolean = True
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        If Me.LevelIndex = EnumLevelIndex.Index01 Then
                            If Me.objGroup.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup) = True Then
                                PrimaryGroup = Me.objGroup.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.LevelGroup)
                                index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(PrimaryGroup)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(PrimaryGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New Core_Module.FrmCariMasterProductGroup
                                ij.ProductGroupLevel = Me.LevelGroup
                                ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryProductGroup <> -1 Then

                                    index = Me.GetIndexItemGroup(ij.PrimaryProductGroup, e.RowIndex)
                                    If index < 0 Then
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexProduct = e.RowIndex
                                                Me.BackgroundWorker2.RunWorkerAsync()
                                            Else
                                                Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                                            Exit Sub
                                        Else
                                            Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                                            Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.objGroup.GetGroupCode(ij.PrimaryProductGroup)
                                            Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.objGroup.GetGroupDesc(ij.PrimaryProductGroup)
                                        End If
                                    End If
                                ElseIf ij.PrimaryProductGroup = -1 Then
                                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                            'Me.ObjCD.SetNumber(Me.dgvGroup)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index02 Then
                            EndEditCategory(sender, e)
                        ElseIf Me.LevelIndex = EnumLevelIndex.Index03 Then
                            EndEditProduct(sender, e)
                        End If
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjCategory.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value, LvlGroup) = True Then
                            PrimaryGroup = Me.ObjCategory.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCategory
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCategory <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryCategory, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryCategory
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                                    End If
                                End If
                            ElseIf ij.PrimaryCategory = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvGroup.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryGroup As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim LvlGroup As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel = Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        Try
            With row
                If IsNothing(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColGroup_Code.Index Then
                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryGroup = Me.ObjProduct.GetPrimary(Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemGroup(PrimaryGroup, e.RowIndex)
                            If index < 0 Then
                                Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = PrimaryGroup
                                Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryGroup)
                                Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryGroup)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.txtCari.Text = Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> -1 Then

                                index = Me.GetIndexItemGroup(ij.PrimaryProduct, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexProduct = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                                        Exit Sub
                                    Else
                                        Me.DgvGroup.Item(Me.ColGroup_Primary.Index, e.RowIndex).Value = ij.PrimaryProduct
                                        Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                        Me.DgvGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                    End If
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
                                If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                                If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndexProduct = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                    End If
                Else
                    If Me.DgvGroup.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                    If Me.DgvGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexProduct = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvGroup)
                        Exit Sub
                    End If
                End If
                'Me.dgvGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvGroup.Columns(Me.ColGroup_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            Me.DgvGroup.Columns(Me.ColGroup_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Function GetIndexItemGroup(ByVal vPrimary As Double, _
                    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                RowDGV = Me.DgvGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColGroup_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvGroup.Rows.RemoveAt(EraseRowIndexProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Dgv Customer"
    Private Sub dgvCustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCustomer_Code.Index Then
                Me.Buffer = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomer.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvCustomer.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            With Row
                If IsNothing(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColCustomer_Code.Index Then
                        Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.TabPage4.Text = "Customer Data" Then
                            If Me.ObjCust.ApakahSudahAda(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Index = Me.GetIndexItemInput_Customer(Me.ObjCust.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value))
                                If Index < 0 Then
                                    Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = Me.ObjCust.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value)
                                    'Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjCust.GetCustomerCode(Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value)
                                    Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = ""
                                        Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexSuppCust = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New Core_Module.FrmCariMasterCustomer
                                ij.txtCari.Text = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimaryCustomer <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    Index = Me.GetIndexItemInput_Customer(ij.PrimaryCustomer)
                                    If Index < 0 Then
                                        Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjCust.GetCustomerCode(ij.PrimaryCustomer)
                                        Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(ij.PrimaryCustomer)
                                        Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = ij.PrimaryCustomer
                                    Else
                                        If e.RowIndex <> Index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexSuppCust = e.RowIndex
                                                Me.BackgroundWorker4.RunWorkerAsync()
                                            Else
                                                Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCekData.SetNumber(Me.dgvItem)
                                            Exit Sub
                                        Else
                                            Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjCust.GetCustomerCode(ij.PrimaryCustomer)
                                            Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(ij.PrimaryCustomer)
                                            Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = ij.PrimaryCustomer
                                        End If
                                    End If
                                ElseIf ij.PrimaryCustomer = -1 Then
                                    If Me.dgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndexSuppCust = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    End If
                                End If
                                ij = Nothing
                            End If
                        ElseIf Me.TabPage4.Text = "Supplier Data" Then
                            If Me.ObjSupp.ApakahSudahAda(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Index = Me.GetIndexItemInput_Customer(Me.ObjSupp.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value))
                                If Index < 0 Then
                                    Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = Me.ObjSupp.GetPrimary(Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value)
                                    'Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierCode(Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value)
                                    Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierName(Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value)
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndexSuppCust = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New Core_Module.FrmCariMasterSupplier
                                ij.txtCari.Text = Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.PrimarySupplier <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    Index = Me.GetIndexItemInput_Customer(ij.PrimarySupplier)
                                    If Index < 0 Then
                                        Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierCode(ij.PrimarySupplier)
                                        Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierName(ij.PrimarySupplier)
                                        Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                    Else
                                        If e.RowIndex <> Index Then
                                            MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndexSuppCust = e.RowIndex
                                                Me.BackgroundWorker4.RunWorkerAsync()
                                            Else
                                                Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            'Me.ObjCekData.SetNumber(Me.dgvItem)
                                            Exit Sub
                                        Else
                                            Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierCode(ij.PrimarySupplier)
                                            Me.dgvCustomer.Item(Me.ColCustomer_Desc.Index, e.RowIndex).Value = Me.ObjSupp.GetSupplierName(ij.PrimarySupplier)
                                            Me.dgvCustomer.Item(Me.ColCustomer_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                        End If
                                    End If
                                ElseIf ij.PrimarySupplier = -1 Then
                                    If Me.dgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                        Me.EraseRowIndexSuppCust = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                        'Me.ObjCekData.SetNumber(Me.dgvItem)
                    End If
                Else
                    If Me.dgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexSuppCust = e.RowIndex
                            Me.BackgroundWorker4.RunWorkerAsync()
                        Else
                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If
                    If Me.dgvCustomer.Item(Me.ColCustomer_Code.Index, e.RowIndex).Value <> "" Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndexSuppCust = e.RowIndex
                            Me.BackgroundWorker4.RunWorkerAsync()
                        Else
                            Me.dgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvLocation)
                        Exit Sub
                    End If

                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput_Customer(ByVal vPrimaryCustomer As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvCustomer.Rows
                If rowDGV.Cells(Me.ColCustomer_Primary.Index).Value = vPrimaryCustomer Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.dgvCustomer.Rows.RemoveAt(Me.EraseRowIndexSuppCust)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "PO Number"
    Private Sub BtnFindPONumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPONumber.Click
        Dim FrmPO As New FrmCariTransaksiPO
        Dim FrmSO As New FrmCariTransaksiSO
        FrmPO.BackColorAll = Me.BackColorAll
        FrmSO.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPONumber.Text <> "" Then
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    FrmPO.txtCari.Text = Me.TxtPONumber.Text
                ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                    FrmSO.txtCari.Text = Me.TxtPONumber.Text
                End If
            End If

            If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                FrmPO.ShowDialog()
                If FrmPO.PrimaryPO > 0 Then
                    Me.TxtPrimaryPO.Text = FrmPO.PrimaryPO
                    Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                    Me.ChkFilterByDate.Focus()
                Else
                    Me.TxtPrimaryPO.Text = "-1"
                    Me.TxtPONumber.Focus()
                End If
            ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                FrmSO.ShowDialog()
                If FrmSO.PrimaryTransaksi > 0 Then
                    Me.TxtPrimaryPO.Text = FrmSO.PrimaryTransaksi
                    Me.TxtPONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryPO.Text)
                    Me.ChkFilterByDate.Focus()
                Else
                    Me.TxtPrimaryPO.Text = "-1"
                    Me.TxtPONumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPO.Dispose()
            FrmSO.Dispose()
        End Try
    End Sub

    Private Sub TxtPONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPONumber.KeyDown
        Dim FrmPO As New FrmCariTransaksiPO
        Dim FrmSO As New FrmCariTransaksiSO
        Try
            If e.KeyValue = 13 Then
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    If Me.TxtPONumber.Text <> "" Then
                        If Me.ObjPO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                            Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtPONumber.Text)
                            Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                            Me.ChkFilterByDate.Focus()
                        Else
                            Me.btnFindPONumber.PerformClick()
                        End If
                    Else
                        Me.btnFindPONumber.PerformClick()
                    End If
                ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
                Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                    If Me.TxtPONumber.Text <> "" Then
                        If Me.ObjSO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                            Me.TxtPrimaryPO.Text = Me.ObjSO.GetPrimary(Me.TxtPONumber.Text)
                            Me.TxtPONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryPO.Text)
                            Me.ChkFilterByDate.Focus()
                        Else
                            Me.btnFindPONumber.PerformClick()
                        End If
                    Else
                        Me.btnFindPONumber.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPO.Dispose()
            FrmSO.Dispose()
        End Try
    End Sub

    Private Sub TxtPONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.Leave
        Dim FrmPO As New FrmCariTransaksiPO
        Dim FrmSO As New FrmCariTransaksiSO
        FrmPO.BackColorAll = Me.BackColorAll
        FrmSO.BackColorAll = Me.BackColorAll
        Try
            If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                If Me.TxtPONumber.Text <> "" Then
                    If Me.ObjPO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                        Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtPONumber.Text)
                        Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                    Else
                        If Me.TxtPONumber.Text <> "" Then
                            FrmPO.txtCari.Text = Me.TxtPONumber.Text
                        End If

                        FrmPO.ShowDialog()
                        If FrmPO.PrimaryPO > 0 Then
                            Me.TxtPrimaryPO.Text = FrmPO.PrimaryPO
                            Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                        Else
                            Me.TxtPrimaryPO.Text = "-1"
                            Me.TxtPONumber.Clear()
                        End If
                    End If
                Else
                    Me.TxtPONumber.Clear()
                    Me.TxtPrimaryPO.Text = "-1"
                End If
            ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                If Me.TxtPONumber.Text <> "" Then
                    If Me.ObjSO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                        Me.TxtPrimaryPO.Text = Me.ObjSO.GetPrimary(Me.TxtPONumber.Text)
                        Me.TxtPONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimaryPO.Text)
                    Else
                        If Me.TxtPONumber.Text <> "" Then
                            FrmSO.txtCari.Text = Me.TxtPONumber.Text
                        End If

                        FrmSO.ShowDialog()
                        If FrmSO.PrimaryTransaksi > 0 Then
                            Me.TxtPrimaryPO.Text = FrmSO.PrimaryTransaksi
                            Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                        Else
                            Me.TxtPrimaryPO.Text = "-1"
                            Me.TxtPONumber.Clear()
                        End If
                    End If
                Else
                    Me.TxtPONumber.Clear()
                    Me.TxtPrimaryPO.Text = "-1"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPO.Dispose()
            FrmSO.Dispose()
        End Try
    End Sub

    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.GotFocus
        Try
            Dim txt As TextBox = TryCast(sender, TextBox)
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_lostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.LostFocus
        Try
            Dim txt As TextBox = TryCast(sender, TextBox)
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Product"
    Private Sub GetPrimaryMinMax(ByRef vPrimaryMin_Result As Double, ByRef vPrimaryMax_Result As Double, _
        ByVal vPrimary_Pembanding As Double)
        Try
            If vPrimary_Pembanding <= vPrimaryMin_Result Then
                vPrimaryMin_Result = vPrimary_Pembanding
            End If
            If vPrimary_Pembanding >= vPrimaryMax_Result Then
                vPrimaryMax_Result = vPrimary_Pembanding
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetFilterDGVProduct(ByRef vFilter As String, _
    ByVal vEnumJenisProductGroup As FrmLaporanStockInOut.JenisProductGroup, _
    ByVal vPrimaryParent As Double)
        Try
            If vEnumJenisProductGroup = JenisProductGroup.Category Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Category] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Category] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Sub_Brand Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Sub Brand] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Sub Brand] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Brand Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Brand] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Brand] = " & vPrimaryParent.ToString & ""
                End If
            ElseIf vEnumJenisProductGroup = JenisProductGroup.Product_Group Then
                If String.IsNullOrEmpty(vFilter) = False Then
                    vFilter &= " OR [Primary Group] = " & vPrimaryParent.ToString & ""
                Else
                    vFilter = "[Primary Group] = " & vPrimaryParent.ToString & ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetFilterProduct_Purchase(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_PurchaseDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_PurchaseRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub

    Public Sub GetFilterProduct_Sales(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_SalesDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_SalesRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub

    Public Sub GetFilterProduct_Adjustment(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_AdjustmentDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_AdjustmentRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub
#End Region

#Region "Get_Filter_Product_WAREHOUSE"
    Public Sub GetFilterProduct_Purchase_WAREHOUSE(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_Purchase_WHDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_Purchase_WHRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub

    Public Sub GetFilterProduct_Sales_WAREHOUSE(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_Sales_WHDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_Sales_WHRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub

    Public Sub GetFilterProduct_Adjustment_WAREHOUSE(ByRef vParamString As String, _
    ByRef TBL_Stock As DseLaporanStockInOut.TBL_Adjustment_WHDataTable)
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Ketemu As Boolean = False
        Dim DTA_Group As New DseLaporanStockInOutTableAdapters.TblGroupTableAdapter
        Dim TBL_Group As New DseLaporanStockInOut.TblGroupDataTable
        Dim Row_Group As DseLaporanStockInOut.TblGroupRow
        Dim DTA_Brand As New DseLaporanStockInOutTableAdapters.TblBrandTableAdapter
        Dim TBL_Brand As New DseLaporanStockInOut.TblBrandDataTable
        Dim Row_Brand As DseLaporanStockInOut.TblBrandRow
        Dim DTA_SubBrand As New DseLaporanStockInOutTableAdapters.TblSubBrandTableAdapter
        Dim TBL_SubBrand As New DseLaporanStockInOut.TblSubBrandDataTable
        Dim Row_SubBrand As DseLaporanStockInOut.TblSubBrandRow
        Dim DTA_Category As New DseLaporanStockInOutTableAdapters.TblCategoryTableAdapter
        Dim TBL_Category As New DseLaporanStockInOut.TblCategoryDataTable
        Dim Row_Category As DseLaporanStockInOut.TblCategoryRow
        Dim Row_Stock As DseLaporanStockInOut.TBL_Adjustment_WHRow
        Dim ListPrimaryGroupParent As New List(Of Double)
        Try
            vParamString = ""
            If Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product_Group Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Product_Group, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Group.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Group.FillBy_PrimaryProductGroup(TBL_Group, PrimaryMax, PrimaryMin)
                    If TBL_Group.Rows.Count > 0 Then
                        TBL_Group.DefaultView.RowFilter = Filter
                        TBL_Group.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Group.DefaultView.Count - 1
                            Row_Group = TBL_Group.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Group.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Group = Row_Group.Primary_Group
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Group.Primary_Group) = False Then
                                    ListPrimaryGroupParent.Add(Row_Group.Primary_Group)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    Else
                                        vParamString = "[Primary Group] = " & Row_Group.Primary_Group.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Group] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Brand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Brand.FillBy_PrimaryBrand(TBL_Brand, PrimaryMax, PrimaryMin)
                    If TBL_Brand.Rows.Count > 0 Then
                        TBL_Brand.DefaultView.RowFilter = Filter
                        TBL_Brand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Brand.DefaultView.Count - 1
                            Row_Brand = TBL_Brand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Brand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Brand = Row_Brand.Primary_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Brand.Primary_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_Brand.Primary_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary Brand] = " & Row_Brand.Primary_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Brand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Sub_Brand Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Sub_Brand, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_SubBrand.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_SubBrand.FillBy_PrimarySubBrand(TBL_SubBrand, PrimaryMax, PrimaryMin)
                    If TBL_SubBrand.Rows.Count > 0 Then
                        TBL_SubBrand.DefaultView.RowFilter = Filter
                        TBL_SubBrand.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_SubBrand.DefaultView.Count - 1
                            Row_SubBrand = TBL_SubBrand.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_SubBrand.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_SubBrand = Row_SubBrand.Primary_Sub_Brand
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_SubBrand.Primary_Sub_Brand) = False Then
                                    ListPrimaryGroupParent.Add(Row_SubBrand.Primary_Sub_Brand)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    Else
                                        vParamString = "[Primary SubBrand] = " & Row_SubBrand.Primary_Sub_Brand.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary SubBrand] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Category Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                        Me.GetFilterDGVProduct(Filter, JenisProductGroup.Category, CDbl(RowDGV.Cells(Me.ColGroup_Primary.Index).Value))
                    Next
                    DTA_Category.Connection.ConnectionString = Core_Module.GetConstring
                    DTA_Category.FillBy_PrimaryCategory(TBL_Category, PrimaryMax, PrimaryMin)
                    'Input Primary Parent Group
                    If TBL_Category.Rows.Count > 0 Then
                        TBL_Category.DefaultView.RowFilter = Filter
                        TBL_Category.DefaultView.Sort = "[Primary Product] ASC"
                        For ik As Integer = 0 To TBL_Category.DefaultView.Count - 1
                            Row_Category = TBL_Category.DefaultView(ik).Row
                            TBL_Stock.DefaultView.RowFilter = Nothing
                            TBL_Stock.DefaultView.RowFilter = "[Primary Product] = " & Row_Category.Primary_Product.ToString & ""
                            For il As Integer = 0 To TBL_Stock.DefaultView.Count - 1
                                Row_Stock = TBL_Stock.DefaultView(il).Row
                                'Ada data yang sesuai, di dalam TBLStockInOut
                                Ketemu = True
                                'Input ke dalam "TBLStock" Primary Parent Group Nya
                                Row_Stock.Primary_Category = Row_Category.Primary_Category
                                'Input Parameter By Primary Parent Group
                                If ListPrimaryGroupParent.Contains(Row_Category.Primary_Category) = False Then
                                    ListPrimaryGroupParent.Add(Row_Category.Primary_Category)
                                    If String.IsNullOrEmpty(vParamString) = False Then
                                        vParamString &= " OR [Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    Else
                                        vParamString = "[Primary Category] = " & Row_Category.Primary_Category.ToString & ""
                                    End If
                                End If
                            Next
                        Next
                        'Kalo DGV Product ada tetapi tidak satu pun sesuai
                        If Ketemu = False Then
                            vParamString = "[Primary Category] = -2"
                        End If
                    End If
                End If
            ElseIf Me.CboPilihGroupLevel.SelectedIndex = FrmLaporanStockTransfer.JenisProductGroup.Product Then
                If Me.DgvGroup.Rows.Count > 1 Then
                    For ij As Integer = 0 To Me.DgvGroup.Rows.Count - 2
                        RowDGV = Me.DgvGroup.Rows(ij)
                        If String.IsNullOrEmpty(vParamString) = False Then
                            vParamString &= " OR [Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        Else
                            vParamString = "[Primary Product] = " & RowDGV.Cells(Me.ColGroup_Primary.Index).Value.ToString & ""
                        End If
                    Next
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
                TBL_Stock.DefaultView.RowFilter = Nothing
                TBL_Stock.DefaultView.Sort = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Group.Dispose()
            TBL_Group.Dispose()
            DTA_Brand.Dispose()
            TBL_Brand.Dispose()
            DTA_SubBrand.Dispose()
            TBL_SubBrand.Dispose()
            DTA_Category.Dispose()
            TBL_Category.Dispose()
            ListPrimaryGroupParent = Nothing
        End Try
    End Sub
#End Region

#Region "Get_Filter_BatchNo"
    Public Sub GetFilterBatchNo(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.DgvBatchNo.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.DgvBatchNo.Rows.Count - 2
                    RowDGV = Me.DgvBatchNo.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR [Batch No] LIKE '" & RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value.ToString & "%'"
                    Else
                        vParamString = "[Batch No] LIKE '" & RowDGV.Cells(Me.ColBatch_BatchNo.Index).Value.ToString & "%'"
                    End If
                Next
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_GoodORService"
    Public Sub GetFilterGoodsORService(ByRef vParamString As String)
        Try
            vParamString = ""
            If Me.CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Goods_Only Then
                vParamString = "[GoodsService] = 0"
            ElseIf CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Service_Only Then
                vParamString = "[GoodsService] = 1"
            ElseIf CboScope.SelectedIndex = FrmLaporanStockTransfer.EnumScope.Both Then
                vParamString = "[GoodsService] = 0 OR [GoodsService] = 1"
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Location"
    Private Function GetLocationAndDestination(ByVal vPrimaryLocation As Double) As String
        Dim Retval As String = ""
        Try
            Retval = "[Primary Location] = " & vPrimaryLocation.ToString & ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterLocation(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.DgvLocation.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.DgvLocation.Rows.Count - 2
                    RowDGV = Me.DgvLocation.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR " & Me.GetLocationAndDestination(CDbl(RowDGV.Cells(Me.ColLocation_Primary.Index).Value))
                    Else
                        vParamString = Me.GetLocationAndDestination( CDbl(RowDGV.Cells(Me.ColLocation_Primary.Index).Value))
                    End If
                Next
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Customer"
    Private Function GetCustomer(ByVal vPrimaryCustomer As Double) As String
        Dim Retval As String = ""
        Try
            Retval = "[Primary Customer] = " & vPrimaryCustomer.ToString & ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterCustomer(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.dgvCustomer.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.dgvCustomer.Rows.Count - 2
                    RowDGV = Me.dgvCustomer.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR " & Me.GetCustomer(CDbl(RowDGV.Cells(Me.ColCustomer_Primary.Index).Value))
                    Else
                        vParamString = Me.GetCustomer(CDbl(RowDGV.Cells(Me.ColCustomer_Primary.Index).Value))
                    End If
                Next
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Supplier"
    Private Function GetSupplier(ByVal vPrimarySupplier As Double) As String
        Dim Retval As String = ""
        Try
            Retval = "[Primary Supplier] = " & vPrimarySupplier.ToString & ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterSupplier(ByRef vParamString As String)
        Dim RowDGV As DataGridViewRow
        Try
            vParamString = ""
            If Me.dgvCustomer.Rows.Count > 1 Then
                For ij As Integer = 0 To Me.dgvCustomer.Rows.Count - 2
                    RowDGV = Me.dgvCustomer.Rows(ij)
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR " & Me.GetSupplier(CDbl(RowDGV.Cells(Me.ColCustomer_Primary.Index).Value))
                    Else
                        vParamString = Me.GetSupplier(CDbl(RowDGV.Cells(Me.ColCustomer_Primary.Index).Value))
                    End If
                Next
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_Adjust"
    Private Function GetAdjust(ByVal vPrimaryAdjust As Double) As String
        Dim Retval As String = ""
        Try
            Retval = "[PrimaryAdjType] = " & vPrimaryAdjust.ToString & ""
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterAdjust(ByRef vParamString As String)
        Dim RowView As Data.DataRowView
        Dim PrimaryAdjustType As Double
        Try
            vParamString = ""
            If Me.clbAdjustmentType.Items.Count > 1 Then
                For Each RowView In Me.clbAdjustmentType.CheckedItems
                    PrimaryAdjustType = RowView.Row.Item("Primary")
                    If String.IsNullOrEmpty(vParamString) = False Then
                        vParamString &= " OR " & Me.GetAdjust(PrimaryAdjustType)
                    Else
                        vParamString = Me.GetAdjust(PrimaryAdjustType)
                    End If
                Next
            End If
            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Filter_POAtoSONumber"
    Private Function GetPOAtoSONumber(ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
        ByVal vPrimaryPOatoSO As Double) As String
        Dim Retval As String = ""
        Try
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Retval = "[Primary PO] = " & vPrimaryPOatoSO.ToString & ""
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Retval = "[Primary Sales] = " & vPrimaryPOatoSO.ToString & ""
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetFilterPOAtoSONumber(ByRef vParamString As String, _
    ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut)
        Try
            vParamString = ""
            If Me.TxtPrimaryPO.Text <> "-1" Then
                If String.IsNullOrEmpty(vParamString) = False Then
                    vParamString &= " OR " & Me.GetPOAtoSONumber(TypeStockInOut, CDbl(Me.TxtPrimaryPO.Text))
                Else
                    vParamString = Me.GetPOAtoSONumber(TypeStockInOut, CDbl(Me.TxtPrimaryPO.Text))
                End If
            End If

            'Return Value
            If String.IsNullOrEmpty(vParamString) = False Then
                vParamString = "(" & vParamString & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Overall_Filter"
    Public Function GetOverallFilter(ByRef TBLPurchase As DseLaporanStockInOut.TBL_PurchaseDataTable, _
        ByRef TBLSales As DseLaporanStockInOut.TBL_SalesDataTable, _
        ByRef TBLAdjust As DseLaporanStockInOut.TBL_AdjustmentDataTable, _
        ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut) As String
        Dim Retval As String = ""
        Dim ProductStr As String = ""
        Dim BatchNoStr As String = ""
        Dim LocationStr As String = ""
        Dim GoodsStr As String = ""
        Dim POSOStr As String = ""
        Dim SupplierStr As String = ""
        Dim CustomerStr As String = ""
        Dim AdjustStr As String = ""
        Try
            'Product
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Me.GetFilterProduct_Purchase(ProductStr, TBLPurchase)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Me.GetFilterProduct_Sales(ProductStr, TBLSales)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Me.GetFilterProduct_Adjustment(ProductStr, TBLAdjust)
            End If
            'Batch
            Me.GetFilterBatchNo(BatchNoStr)
            'location
            Me.GetFilterLocation(LocationStr)
            'Goods Or Service
            Me.GetFilterGoodsORService(GoodsStr)
            'PO Number
            Me.GetFilterPOAtoSONumber(POSOStr, TypeStockInOut)
            'Adjustment Type
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Me.GetFilterSupplier(SupplierStr)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Me.GetFilterCustomer(CustomerStr)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Me.GetFilterAdjust(AdjustStr)
            End If

            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(ProductStr) = False Then
                    Retval &= " AND " & ProductStr
                End If
            Else
                Retval = ProductStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(BatchNoStr) = False Then
                    Retval &= " AND " & BatchNoStr
                End If
            Else
                Retval = BatchNoStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(LocationStr) = False Then
                    Retval &= " AND " & LocationStr
                End If
            Else
                Retval = LocationStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(GoodsStr) = False Then
                    Retval &= " AND " & GoodsStr
                End If
            Else
                Retval = GoodsStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(POSOStr) = False Then
                    Retval &= " AND " & POSOStr
                End If
            Else
                Retval = POSOStr
            End If
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(SupplierStr) = False Then
                        Retval &= " AND " & SupplierStr
                    End If
                Else
                    Retval = SupplierStr
                End If
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(CustomerStr) = False Then
                        Retval &= " AND " & CustomerStr
                    End If
                Else
                    Retval = CustomerStr
                End If
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(AdjustStr) = False Then
                        Retval &= " AND " & AdjustStr
                    End If
                Else
                    Retval = AdjustStr
                End If
            End If

            'Return Value
            If Retval.StartsWith(" AND ") = True Then
                Retval = Retval.Remove(0, 4)
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetOverallFilter_WAREHOUSE(ByRef TBLPurchase As DseLaporanStockInOut.TBL_Purchase_WHDataTable, _
    ByRef TBLSales As DseLaporanStockInOut.TBL_Sales_WHDataTable, _
    ByRef TBLAdjust As DseLaporanStockInOut.TBL_Adjustment_WHDataTable, _
    ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut) As String
        Dim Retval As String = ""
        Dim ProductStr As String = ""
        Dim BatchNoStr As String = ""
        Dim LocationStr As String = ""
        Dim GoodsStr As String = ""
        Dim POSOStr As String = ""
        Dim SupplierStr As String = ""
        Dim CustomerStr As String = ""
        Dim AdjustStr As String = ""
        Try
            'Product
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Me.GetFilterProduct_Purchase_WAREHOUSE(ProductStr, TBLPurchase)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Me.GetFilterProduct_Sales_WAREHOUSE(ProductStr, TBLSales)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Me.GetFilterProduct_Adjustment_WAREHOUSE(ProductStr, TBLAdjust)
            End If
            'Batch
            Me.GetFilterBatchNo(BatchNoStr)
            'location
            Me.GetFilterLocation(LocationStr)
            'Goods Or Service
            Me.GetFilterGoodsORService(GoodsStr)
            'PO Number
            Me.GetFilterPOAtoSONumber(POSOStr, TypeStockInOut)
            'Adjustment Type
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                Me.GetFilterSupplier(SupplierStr)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Me.GetFilterCustomer(CustomerStr)
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Me.GetFilterAdjust(AdjustStr)
            End If

            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(ProductStr) = False Then
                    Retval &= " AND " & ProductStr
                End If
            Else
                Retval = ProductStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(BatchNoStr) = False Then
                    Retval &= " AND " & BatchNoStr
                End If
            Else
                Retval = BatchNoStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(LocationStr) = False Then
                    Retval &= " AND " & LocationStr
                End If
            Else
                Retval = LocationStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(GoodsStr) = False Then
                    Retval &= " AND " & GoodsStr
                End If
            Else
                Retval = GoodsStr
            End If
            If String.IsNullOrEmpty(Retval) = False Then
                If String.IsNullOrEmpty(POSOStr) = False Then
                    Retval &= " AND " & POSOStr
                End If
            Else
                Retval = POSOStr
            End If
            If TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(SupplierStr) = False Then
                        Retval &= " AND " & SupplierStr
                    End If
                Else
                    Retval = SupplierStr
                End If
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(CustomerStr) = False Then
                        Retval &= " AND " & CustomerStr
                    End If
                Else
                    Retval = CustomerStr
                End If
            ElseIf TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                If String.IsNullOrEmpty(Retval) = False Then
                    If String.IsNullOrEmpty(AdjustStr) = False Then
                        Retval &= " AND " & AdjustStr
                    End If
                Else
                    Retval = AdjustStr
                End If
            End If

            'Return Value
            If Retval.StartsWith(" AND ") = True Then
                Retval = Retval.Remove(0, 4)
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
#End Region

    Private Sub GetDecimal_Maximum_Untuk_SubTotal(ByRef vDecimalMax_Result As Double, _
    ByVal vDecimal_Pembanding As Double)
        Try
            If vDecimal_Pembanding >= vDecimalMax_Result Then
                vDecimalMax_Result = vDecimal_Pembanding
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Get_Data_Purchase"
    Private Sub GetDataPurchase_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
       ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
       ByRef TBLPurchase As DseLaporanStockInOut.TBL_Purchase_WHDataTable, _
       ByVal TBLSales As DseLaporanStockInOut.TBL_Sales_WHDataTable, _
       ByVal TBLAdjust As DseLaporanStockInOut.TBL_Adjustment_WHDataTable, _
       ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
       ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
       ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTAPurchase As New DseLaporanStockInOutTableAdapters.TBL_Purchase_WHTableAdapter
        Dim RowPurchase As DseLaporanStockInOut.TBL_Purchase_WHRow
        Dim ij As Integer = 0
        Try
            'Database
            DTAPurchase.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkFilterByDate.Checked = True Then
                DTAPurchase.FillBy_TanggalDariSampai_INOUTPurchase(TBLPurchase, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTAPurchase.FillBy_TipeTransaksi(TBLPurchase, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLPurchase.Rows.Count > 0 Then
                TBLPurchase.DefaultView.RowFilter = Nothing
                TBLPurchase.DefaultView.Sort = Nothing
                TBLPurchase.DefaultView.RowFilter = Me.GetOverallFilter_WAREHOUSE(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLPurchase.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLPurchase.DefaultView.Count > 0 Then
                    While ij <= TBLPurchase.DefaultView.Count - 1
                        RowPurchase = TBLPurchase.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowPurchase.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowPurchase.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowPurchase.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            '.Primary_AutoNumber = RowPurchase.Primary
                            .Primary_Stock = RowPurchase.Primary_Transaksi
                            .Primary_Product = RowPurchase.Primary_Product
                            ._Date = RowPurchase._Date
                            .Document_No = RowPurchase.Document_No
                            .Remarks = RowPurchase.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowPurchase.Product_Code
                            .Description = RowPurchase.Description
                            ._Decimal = RowPurchase._Decimal
                            .Primary_Unit = RowPurchase.Primary_Unit
                            .Unit = RowPurchase.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowPurchase.Qty
                            .Batch_No = RowPurchase.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowPurchase.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowPurchase.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowPurchase._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = RowPurchase.Primary_PO
                            .PO_SO_Number = RowPurchase.PO_Number
                            .PrimaryAdjType = -1
                            .AdjTypeCode = ""
                            .Primary_Supp_Cust = RowPurchase.Primary_Supplier
                            .Cust_Supp_Code = RowPurchase.Supplier_Name
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_Purchase_WAREHOUSE(TBLSubTotal, RowPurchase)
                        'Hapus Baris
                        TBLPurchase.Rows.Remove(RowPurchase)
                    End While
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Purchase_WAREHOUSE(TBLSort, TBLPurchase, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTAPurchase.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Purchase_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLDummy As DseLaporanStockInOut.TBL_Purchase_WHDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_Purchase_WHRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Dim ExpDate As Date = #1/1/1950#
        Dim idx As Integer = 0
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Filter TBLDummy, input sisa baris By Primary Stock
            ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & ""
            TBLDummy.DefaultView.RowFilter = ParamStr
            If TBLDummy.DefaultView.Count > 0 Then
                While idx <= TBLDummy.DefaultView.Count - 1
                    RowDummy = TBLDummy.DefaultView(idx).Row
                    'Primary Purchase
                    Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                    'New Data
                    RowSort = TBLSort.NewTblCetakRow
                    With RowSort
                        .Primary_Stock = RowDummy.Primary_Transaksi
                        .Primary_Product = RowDummy.Primary_Product
                        ._Date = RowDummy._Date
                        .Document_No = RowDummy.Document_No
                        .Remarks = RowDummy.Remarks
                        If String.IsNullOrEmpty(.Remarks) = True Then
                            .Remarks = "-"
                        End If
                        .Code_No = RowDummy.Product_Code
                        .Description = RowDummy.Description
                        ._Decimal = RowDummy._Decimal
                        .Primary_Unit = RowDummy.Primary_Unit
                        .Unit = RowDummy.Unit
                        If String.IsNullOrEmpty(.Unit) = True Then
                            .Unit = "-"
                        End If
                        .Qty = RowDummy.Qty
                        .Batch_No = RowDummy.Batch_No
                        If String.IsNullOrEmpty(.Batch_No) = True Then
                            .Batch_No = "-"
                        End If
                        .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                        If CDate(.Expiry_Date) = #1/1/1950# Then
                            .Expiry_Date = "-"
                        End If
                        .Location = RowDummy.Location
                        If String.IsNullOrEmpty(.Location) = True Then
                            .Location = "-"
                        End If
                        .Rounding = .Rounding - RowDummy._Decimal
                        .SubTotal = ""
                        '--Tambahan Field--
                        .Primary_PO_SO = RowDummy.Primary_PO
                        .PO_SO_Number = RowDummy.PO_Number
                        .PrimaryAdjType = -1
                        .AdjTypeCode = ""
                        .Primary_Supp_Cust = RowDummy.Primary_Supplier
                        .Cust_Supp_Code = RowDummy.Supplier_Name
                    End With
                    TBLSort.Rows.Add(RowSort)
                    'Input Sub Total
                    Me.InputTBLSubTotal_Purchase_WAREHOUSE(TBLSubTotal, RowDummy)
                    'hapus baris yang udah di input
                    TBLDummy.Rows.Remove(RowDummy)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_Purchase_WAREHOUSE(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_Purchase_WHRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataPurchase(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
      ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
      ByRef TBLPurchase As DseLaporanStockInOut.TBL_PurchaseDataTable, _
      ByVal TBLSales As DseLaporanStockInOut.TBL_SalesDataTable, _
      ByVal TBLAdjust As DseLaporanStockInOut.TBL_AdjustmentDataTable, _
      ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
      ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
      ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTAPurchase As New DseLaporanStockInOutTableAdapters.TBL_PurchaseTableAdapter
        Dim RowPurchase As DseLaporanStockInOut.TBL_PurchaseRow
        Dim TBLDummyPurchase As New DseLaporanStockInOut.TBL_PurchaseDataTable
        Try
            'Database
            DTAPurchase.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkFilterByDate.Checked = True Then
                DTAPurchase.FillBy_TanggalDariSampai_INOUTPurchase(TBLPurchase, vEndDate, vStartDate, Me.TypeStockInOut)
                DTAPurchase.FillBy_TanggalDariSampai_INOUTPurchase(TBLDummyPurchase, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTAPurchase.FillBy_TipeStock(TBLPurchase, Me.TypeStockInOut)
                DTAPurchase.FillBy_TipeStock(TBLDummyPurchase, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLPurchase.Rows.Count > 0 Then
                TBLPurchase.DefaultView.RowFilter = Nothing
                TBLPurchase.DefaultView.Sort = Nothing
                TBLPurchase.DefaultView.RowFilter = Me.GetOverallFilter(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLPurchase.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLPurchase.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBLPurchase.DefaultView.Count - 1
                        RowPurchase = TBLPurchase.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowPurchase.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowPurchase.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowPurchase.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            .Primary_AutoNumber = RowPurchase.Primary
                            .Primary_Stock = RowPurchase.Primary_Transaksi
                            .Primary_Product = RowPurchase.Primary_Product
                            ._Date = RowPurchase._Date
                            .Document_No = RowPurchase.Document_No
                            .Remarks = RowPurchase.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowPurchase.Product_Code
                            .Description = RowPurchase.Description
                            ._Decimal = RowPurchase._Decimal
                            .Primary_Unit = RowPurchase.Primary_Unit
                            .Unit = RowPurchase.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowPurchase.Qty
                            .Batch_No = RowPurchase.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowPurchase.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowPurchase.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowPurchase._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = RowPurchase.Primary_PO
                            .PO_SO_Number = RowPurchase.PO_Number
                            .PrimaryAdjType = -1
                            .AdjTypeCode = ""
                            .Primary_Supp_Cust = RowPurchase.Primary_Supplier
                            .Cust_Supp_Code = RowPurchase.Supplier_Name
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_Purchase(TBLSubTotal, RowPurchase)
                    Next
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    TBLDummyPurchase.DefaultView.RowFilter = ""
                    TBLDummyPurchase.DefaultView.Sort = ""
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Purchase(TBLSort, TBLDummyPurchase, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTAPurchase.Dispose()
            TBLDummyPurchase.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Purchase(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByVal TBLDummy As DseLaporanStockInOut.TBL_PurchaseDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_PurchaseRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Dapatkan primary auto number dari tbl cetak
            TBLSort.DefaultView.RowFilter = "[Primary Stock] = " & vPrimaryStock.ToString & ""
            If TBLSort.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                    RowSort = TBLSort.DefaultView(ij).Row
                    If String.IsNullOrEmpty(Filter) = False Then
                        Filter &= " AND [Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    Else
                        Filter = "[Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    End If
                Next
                '2. Input sisa data item ke TBL Sort (Cetak)
                If String.IsNullOrEmpty(Filter) = False Then
                    TBLDummy.DefaultView.RowFilter = ""
                    TBLDummy.DefaultView.Sort = ""
                    ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & " AND " & Filter & ""
                    TBLDummy.DefaultView.RowFilter = ParamStr
                    If TBLDummy.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TBLDummy.DefaultView.Count - 1
                            RowDummy = TBLDummy.DefaultView(ij).Row
                            'Primary Purchase
                            Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                            'New Data
                            RowSort = TBLSort.NewTblCetakRow
                            With RowSort
                                .Primary_AutoNumber = RowDummy.Primary
                                .Primary_Stock = RowDummy.Primary_Transaksi
                                .Primary_Product = RowDummy.Primary_Product
                                ._Date = RowDummy._Date
                                .Document_No = RowDummy.Document_No
                                .Remarks = RowDummy.Remarks
                                If String.IsNullOrEmpty(.Remarks) = True Then
                                    .Remarks = "-"
                                End If
                                .Code_No = RowDummy.Product_Code
                                .Description = RowDummy.Description
                                ._Decimal = RowDummy._Decimal
                                .Primary_Unit = RowDummy.Primary_Unit
                                .Unit = RowDummy.Unit
                                If String.IsNullOrEmpty(.Unit) = True Then
                                    .Unit = "-"
                                End If
                                .Qty = RowDummy.Qty
                                .Batch_No = RowDummy.Batch_No
                                If String.IsNullOrEmpty(.Batch_No) = True Then
                                    .Batch_No = "-"
                                End If
                                .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                                If CDate(.Expiry_Date) = #1/1/1950# Then
                                    .Expiry_Date = "-"
                                End If
                                .Location = RowDummy.Location
                                If String.IsNullOrEmpty(.Location) = True Then
                                    .Location = "-"
                                End If
                                .Rounding = .Rounding - RowDummy._Decimal
                                .SubTotal = ""
                                '--Tambahan Field--
                                .Primary_PO_SO = RowDummy.Primary_PO
                                .PO_SO_Number = RowDummy.PO_Number
                                .PrimaryAdjType = -1
                                .AdjTypeCode = ""
                                .Primary_Supp_Cust = RowDummy.Primary_Supplier
                                .Cust_Supp_Code = RowDummy.Supplier_Name
                            End With
                            TBLSort.Rows.Add(RowSort)
                            'Input Sub Total
                            Me.InputTBLSubTotal_Purchase(TBLSubTotal, RowDummy)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_Purchase(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_PurchaseRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanTBLSubTotal_Purchase(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TblCetakRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Stock.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty -= RowStock.Qty
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Stock
                RowSubTotal.Total_Qty = RowStock.Qty * -1
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CompareDgnGRN(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLGRN As DseLaporanStockInOut.tbltrans_grnDataTable, _
    ByVal PrimaryMax As Double, ByVal PrimaryMin As Double, _
    ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim Filter As String = ""
        Dim ij As Integer = 0
        Dim RowGRN As DseLaporanStockInOut.tbltrans_grnRow
        Dim WithAOP As Boolean = False
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim DTAGRN As New DseLaporanStockInOutTableAdapters.tbltrans_grnTableAdapter
        Dim DTAAOP As New DseLaporanStockInOutTableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New DseLaporanStockInOut.tbltrans_aopDataTable
        Dim RowAOP As DseLaporanStockInOut.tbltrans_aopRow
        Try
            DTAGRN.Connection.ConnectionString = Core_Module.GetConstring
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTAGRN.FillBy_PrimaryMinMax_INPurchase(TBLGRN, PrimaryMax, PrimaryMin)
            DTAAOP.Fill(TBLAOP)
            If TBLGRN.Rows.Count > 0 Then
                While ij <= TBLSort.Rows.Count - 1
                    RowSort = TBLSort.Rows(ij)
                    WithAOP = False
                    If Me.CboInPurchase.SelectedIndex = 0 Then
                        'With AOP
                        Filter = "[primary aop] <> -1" & " AND [Primary Stock In Purchase] = " & RowSort.Primary_Stock.ToString & ""
                        WithAOP = True
                    ElseIf Me.CboInPurchase.SelectedIndex = 1 Then
                        'W/O AOP
                        Filter = "[primary aop] = -1" & " AND [Primary Stock In Purchase] = " & RowSort.Primary_Stock.ToString & ""
                        WithAOP = False
                    End If
                    TBLGRN.DefaultView.RowFilter = Nothing
                    TBLGRN.DefaultView.RowFilter = Filter
                    If TBLGRN.DefaultView.Count <= 0 Then
                        'SubTotal-nya dikurangi
                        Me.KurangkanTBLSubTotal_Purchase(TBLSubTotal, RowSort)
                        'Hapus, kalo ga ketemu data compare-nya
                        TBLSort.Rows.Remove(RowSort)
                    Else
                        ij += 1
                        'Tambahkan No AOP
                        If WithAOP = True Then
                            RowGRN = TBLGRN.DefaultView(0).Row
                            TBLAOP.DefaultView.RowFilter = "[Primary] = " & RowGRN.primary_aop & ""
                            If TBLAOP.DefaultView.Count > 0 Then
                                RowAOP = TBLAOP.DefaultView(0).Row
                                RowSort.GRN_No = RowAOP.No_Transaksi
                            End If
                        End If
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGRN.Dispose()
            DTAAOP.Dispose()
            TBLAOP.Dispose()
        End Try
    End Sub

    Private Sub GetAOPNumber_Both(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLGRN As DseLaporanStockInOut.tbltrans_grnDataTable, _
    ByVal PrimaryMax As Double, ByVal PrimaryMin As Double)
        Dim Filter As String = ""
        Dim ij As Integer = 0
        Dim RowGRN As DseLaporanStockInOut.tbltrans_grnRow
        Dim WithAOP As Boolean = False
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim DTAGRN As New DseLaporanStockInOutTableAdapters.tbltrans_grnTableAdapter
        Dim DTAAOP As New DseLaporanStockInOutTableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New DseLaporanStockInOut.tbltrans_aopDataTable
        Dim RowAOP As DseLaporanStockInOut.tbltrans_aopRow
        Try
            If Me.CboInPurchase.SelectedIndex = 2 Then
                DTAGRN.Connection.ConnectionString = Core_Module.GetConstring
                DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
                DTAGRN.FillBy_PrimaryMinMax_INPurchase(TBLGRN, PrimaryMax, PrimaryMin)
                DTAAOP.Fill(TBLAOP)
                If TBLGRN.Rows.Count > 0 Then
                    While ij <= TBLSort.Rows.Count - 1
                        RowSort = TBLSort.Rows(ij)
                        WithAOP = False
                        If Me.CboInPurchase.SelectedIndex = 2 Then
                            'With AOP
                            Filter = "[primary aop] <> -1" & " AND [Primary Stock In Purchase] = " & RowSort.Primary_Stock.ToString & ""
                            WithAOP = True
                        End If
                        TBLGRN.DefaultView.RowFilter = Nothing
                        TBLGRN.DefaultView.RowFilter = Filter
                        If TBLGRN.DefaultView.Count > 0 Then
                            'Tambahkan No AOP
                            If WithAOP = True Then
                                RowGRN = TBLGRN.DefaultView(0).Row
                                TBLAOP.DefaultView.RowFilter = "[Primary] = " & RowGRN.primary_aop & ""
                                If TBLAOP.DefaultView.Count > 0 Then
                                    RowAOP = TBLAOP.DefaultView(0).Row
                                    RowSort.GRN_No = RowAOP.No_Transaksi
                                End If
                            End If
                        End If
                        ij += 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGRN.Dispose()
            DTAAOP.Dispose()
            TBLAOP.Dispose()
        End Try
    End Sub
#End Region

#Region "Get_Data_Sales"
    Private Sub GetDataSales_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
       ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
       ByRef TBLPurchase As DseLaporanStockInOut.TBL_Purchase_WHDataTable, _
       ByVal TBLSales As DseLaporanStockInOut.TBL_Sales_WHDataTable, _
       ByVal TBLAdjust As DseLaporanStockInOut.TBL_Adjustment_WHDataTable, _
       ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
       ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
       ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTASales As New DseLaporanStockInOutTableAdapters.TBL_Sales_WHTableAdapter
        Dim RowSales As DseLaporanStockInOut.TBL_Sales_WHRow
        Dim ij As Integer = 0
        Try
            'Database
            DTASales.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkFilterByDate.Checked = True Then
                DTASales.FillBy_TanggalDariSampai_INOUTSales(TBLSales, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTASales.FillBy_TipeTransaksi(TBLSales, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLSales.Rows.Count > 0 Then
                TBLSales.DefaultView.RowFilter = Nothing
                TBLSales.DefaultView.Sort = Nothing
                TBLSales.DefaultView.RowFilter = Me.GetOverallFilter_WAREHOUSE(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLSales.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLSales.DefaultView.Count > 0 Then
                    While ij <= TBLSales.DefaultView.Count - 1
                        RowSales = TBLSales.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowSales.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowSales.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowSales.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            '.Primary_AutoNumber = RowSales.Primary
                            .Primary_Stock = RowSales.Primary_Transaksi
                            .Primary_Product = RowSales.Primary_Product
                            ._Date = RowSales._Date
                            .Document_No = RowSales.Document_No
                            .Remarks = RowSales.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowSales.Product_Code
                            .Description = RowSales.Description
                            ._Decimal = RowSales._Decimal
                            .Primary_Unit = RowSales.Primary_Unit
                            .Unit = RowSales.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowSales.Qty
                            .Batch_No = RowSales.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowSales.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowSales.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowSales._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = RowSales.Primary_Sales
                            .PO_SO_Number = RowSales.Sales_Number
                            .PrimaryAdjType = -1
                            .AdjTypeCode = ""
                            .Primary_Supp_Cust = RowSales.Primary_Customer
                            .Cust_Supp_Code = RowSales.Customer_Name
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_Sales_WAREHOUSE(TBLSubTotal, RowSales)
                        'Hapus Baris-nya
                        TBLSales.Rows.Remove(RowSales)
                    End While
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Sales_WAREHOUSE(TBLSort, TBLSales, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTASales.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Sales_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLDummy As DseLaporanStockInOut.TBL_Sales_WHDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_Sales_WHRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Dim ExpDate As Date = #1/1/1950#
        Dim idx As Integer = 0
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Filter TBLDummy, input sisa baris By Primary Stock
            ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & ""
            TBLDummy.DefaultView.RowFilter = ParamStr
            If TBLDummy.DefaultView.Count > 0 Then
                While idx <= TBLDummy.Rows.Count - 1
                    RowDummy = TBLDummy.DefaultView(idx).Row
                    'Primary Purchase
                    Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                    'New Data
                    RowSort = TBLSort.NewTblCetakRow
                    With RowSort
                        '.Primary_AutoNumber = RowDummy.Primary
                        .Primary_Stock = RowDummy.Primary_Transaksi
                        .Primary_Product = RowDummy.Primary_Product
                        ._Date = RowDummy._Date
                        .Document_No = RowDummy.Document_No
                        .Remarks = RowDummy.Remarks
                        If String.IsNullOrEmpty(.Remarks) = True Then
                            .Remarks = "-"
                        End If
                        .Code_No = RowDummy.Product_Code
                        .Description = RowDummy.Description
                        ._Decimal = RowDummy._Decimal
                        .Primary_Unit = RowDummy.Primary_Unit
                        .Unit = RowDummy.Unit
                        If String.IsNullOrEmpty(.Unit) = True Then
                            .Unit = "-"
                        End If
                        .Qty = RowDummy.Qty
                        .Batch_No = RowDummy.Batch_No
                        If String.IsNullOrEmpty(.Batch_No) = True Then
                            .Batch_No = "-"
                        End If
                        .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                        If CDate(.Expiry_Date) = #1/1/1950# Then
                            .Expiry_Date = "-"
                        End If
                        .Location = RowDummy.Location
                        If String.IsNullOrEmpty(.Location) = True Then
                            .Location = "-"
                        End If
                        .Rounding = .Rounding - RowDummy._Decimal
                        .SubTotal = ""
                        '--Tambahan Field--
                        .Primary_PO_SO = RowDummy.Primary_Sales
                        .PO_SO_Number = RowDummy.Sales_Number
                        .PrimaryAdjType = -1
                        .AdjTypeCode = ""
                        .Primary_Supp_Cust = RowDummy.Primary_Customer
                        .Cust_Supp_Code = RowDummy.Customer_Name
                    End With
                    TBLSort.Rows.Add(RowSort)
                    'Input Sub Total
                    Me.InputTBLSubTotal_Sales_WAREHOUSE(TBLSubTotal, RowDummy)
                    'Hapus Baris yang sudah di input
                    TBLDummy.Rows.Remove(RowDummy)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_Sales_WAREHOUSE(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_Sales_WHRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataSales(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByRef TBLPurchase As DseLaporanStockInOut.TBL_PurchaseDataTable, _
    ByVal TBLSales As DseLaporanStockInOut.TBL_SalesDataTable, _
    ByVal TBLAdjust As DseLaporanStockInOut.TBL_AdjustmentDataTable, _
    ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
    ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTASales As New DseLaporanStockInOutTableAdapters.TBL_SalesTableAdapter
        Dim RowSales As DseLaporanStockInOut.TBL_SalesRow
        Dim TBLDummySales As New DseLaporanStockInOut.TBL_SalesDataTable
        Try
            'Database
            DTASales.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ChkFilterByDate.Checked = True Then
                DTASales.FillBy_TanggalDariSampai_INOUTSales(TBLSales, vEndDate, vStartDate, Me.TypeStockInOut)
                DTASales.FillBy_TanggalDariSampai_INOUTSales(TBLDummySales, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTASales.FillBy_TipeStock(TBLSales, Me.TypeStockInOut)
                DTASales.FillBy_TipeStock(TBLDummySales, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLSales.Rows.Count > 0 Then
                TBLSales.DefaultView.RowFilter = Nothing
                TBLSales.DefaultView.Sort = Nothing
                TBLSales.DefaultView.RowFilter = Me.GetOverallFilter(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLSales.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLSales.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBLSales.DefaultView.Count - 1
                        RowSales = TBLSales.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowSales.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowSales.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowSales.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            .Primary_AutoNumber = RowSales.Primary
                            .Primary_Stock = RowSales.Primary_Transaksi
                            .Primary_Product = RowSales.Primary_Product
                            ._Date = RowSales._Date
                            .Document_No = RowSales.Document_No
                            .Remarks = RowSales.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowSales.Product_Code
                            .Description = RowSales.Description
                            ._Decimal = RowSales._Decimal
                            .Primary_Unit = RowSales.Primary_Unit
                            .Unit = RowSales.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowSales.Qty
                            .Batch_No = RowSales.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowSales.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowSales.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowSales._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = RowSales.Primary_Sales
                            .PO_SO_Number = RowSales.Sales_Number
                            .PrimaryAdjType = -1
                            .AdjTypeCode = ""
                            .Primary_Supp_Cust = RowSales.Primary_Customer
                            .Cust_Supp_Code = RowSales.Customer_Name
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_Sales(TBLSubTotal, RowSales)
                    Next
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    TBLDummySales.DefaultView.RowFilter = ""
                    TBLDummySales.DefaultView.Sort = ""
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Sales(TBLSort, TBLDummySales, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTASales.Dispose()
            TBLDummySales.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_Sales(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByVal TBLDummy As DseLaporanStockInOut.TBL_SalesDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_SalesRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Dapatkan primary auto number dari tbl cetak
            TBLSort.DefaultView.RowFilter = "[Primary Stock] = " & vPrimaryStock.ToString & ""
            If TBLSort.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                    RowSort = TBLSort.DefaultView(ij).Row
                    If String.IsNullOrEmpty(Filter) = False Then
                        Filter &= " AND [Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    Else
                        Filter = "[Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    End If
                Next
                '2. Input sisa data item ke TBL Sort (Cetak)
                If String.IsNullOrEmpty(Filter) = False Then
                    TBLDummy.DefaultView.RowFilter = ""
                    TBLDummy.DefaultView.Sort = ""
                    ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & " AND " & Filter & ""
                    TBLDummy.DefaultView.RowFilter = ParamStr
                    If TBLDummy.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TBLDummy.DefaultView.Count - 1
                            RowDummy = TBLDummy.DefaultView(ij).Row
                            'Primary Purchase
                            Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                            'New Data
                            RowSort = TBLSort.NewTblCetakRow
                            With RowSort
                                .Primary_AutoNumber = RowDummy.Primary
                                .Primary_Stock = RowDummy.Primary_Transaksi
                                .Primary_Product = RowDummy.Primary_Product
                                ._Date = RowDummy._Date
                                .Document_No = RowDummy.Document_No
                                .Remarks = RowDummy.Remarks
                                If String.IsNullOrEmpty(.Remarks) = True Then
                                    .Remarks = "-"
                                End If
                                .Code_No = RowDummy.Product_Code
                                .Description = RowDummy.Description
                                ._Decimal = RowDummy._Decimal
                                .Primary_Unit = RowDummy.Primary_Unit
                                .Unit = RowDummy.Unit
                                If String.IsNullOrEmpty(.Unit) = True Then
                                    .Unit = "-"
                                End If
                                .Qty = RowDummy.Qty
                                .Batch_No = RowDummy.Batch_No
                                If String.IsNullOrEmpty(.Batch_No) = True Then
                                    .Batch_No = "-"
                                End If
                                .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                                If CDate(.Expiry_Date) = #1/1/1950# Then
                                    .Expiry_Date = "-"
                                End If
                                .Location = RowDummy.Location
                                If String.IsNullOrEmpty(.Location) = True Then
                                    .Location = "-"
                                End If
                                .Rounding = .Rounding - RowDummy._Decimal
                                .SubTotal = ""
                                '--Tambahan Field--
                                .Primary_PO_SO = RowDummy.Primary_Sales
                                .PO_SO_Number = RowDummy.Sales_Number
                                .PrimaryAdjType = -1
                                .AdjTypeCode = ""
                                .Primary_Supp_Cust = RowDummy.Primary_Customer
                                .Cust_Supp_Code = RowDummy.Customer_Name
                            End With
                            TBLSort.Rows.Add(RowSort)
                            'Input Sub Total
                            Me.InputTBLSubTotal_Sales(TBLSubTotal, RowDummy)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_Sales(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_SalesRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Get_Data_AdjustType"
    Private Function GetAdjTypeCode(ByVal TBL As DseLaporanStockInOut.tbl_setting_adjustmentDataTable, _
    ByVal vPrimaryAdjType As Double) As String
        Dim row As DseLaporanStockInOut.tbl_setting_adjustmentRow
        Dim retval As String = ""
        Try
            If vPrimaryAdjType = -1 Then
                retval = "No Tipe"
                Return retval
            End If
            If vPrimaryAdjType = -2 Then
                retval = "Stock Taking"
            End If
            If vPrimaryAdjType <> -1 Or vPrimaryAdjType <> -2 Then
                TBL.DefaultView.RowFilter = Nothing
                TBL.DefaultView.RowFilter = "[Primary] = " & vPrimaryAdjType.ToString & ""
                If TBL.DefaultView.Count > 0 Then
                    row = TBL.DefaultView(0).Row
                    retval = row.Code
                    Return retval
                End If
            End If
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub GetDataAdjustType_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByRef TBLPurchase As DseLaporanStockInOut.TBL_Purchase_WHDataTable, _
    ByVal TBLSales As DseLaporanStockInOut.TBL_Sales_WHDataTable, _
    ByVal TBLAdjust As DseLaporanStockInOut.TBL_Adjustment_WHDataTable, _
    ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
    ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTAAdjType As New DseLaporanStockInOutTableAdapters.TBL_Adjustment_WHTableAdapter
        Dim RowAdjType As DseLaporanStockInOut.TBL_Adjustment_WHRow
        Dim DTASetting As New DseLaporanStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim TBLSetting As New DseLaporanStockInOut.tbl_setting_adjustmentDataTable
        '---
        Dim ij As Integer = 0
        Try
            'Database
            DTAAdjType.Connection.ConnectionString = Core_Module.GetConstring
            DTASetting.Connection.ConnectionString = Core_Module.GetConstring
            DTASetting.Fill(TBLSetting)
            If Me.ChkFilterByDate.Checked = True Then
                DTAAdjType.FillBy_TanggalDariSampai_INOutAdjust(TBLAdjust, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTAAdjType.FillBy_TipeTransaksi(TBLAdjust, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLAdjust.Rows.Count > 0 Then
                TBLAdjust.DefaultView.RowFilter = Nothing
                TBLAdjust.DefaultView.Sort = Nothing
                TBLAdjust.DefaultView.RowFilter = Me.GetOverallFilter_WAREHOUSE(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLAdjust.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLAdjust.DefaultView.Count > 0 Then
                    While ij <= TBLAdjust.DefaultView.Count - 1
                        RowAdjType = TBLAdjust.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowAdjType.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowAdjType.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowAdjType.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            '.Primary_AutoNumber = RowAdjType.Primary
                            .Primary_Stock = RowAdjType.Primary_Transaksi
                            .Primary_Product = RowAdjType.Primary_Product
                            ._Date = RowAdjType._Date
                            .Document_No = RowAdjType.Document_No
                            .Remarks = RowAdjType.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowAdjType.Product_Code
                            .Description = RowAdjType.Description
                            ._Decimal = RowAdjType._Decimal
                            .Primary_Unit = RowAdjType.Primary_Unit
                            .Unit = RowAdjType.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowAdjType.Qty
                            .Batch_No = RowAdjType.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowAdjType.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowAdjType.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowAdjType._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = -1
                            .PO_SO_Number = ""
                            .PrimaryAdjType = RowAdjType.PrimaryAdjType
                            .AdjTypeCode = Me.GetAdjTypeCode(TBLSetting, .PrimaryAdjType)
                            .Primary_Supp_Cust = -1
                            .Cust_Supp_Code = ""
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_AdjustType_WAREHOUSE(TBLSubTotal, RowAdjType)
                        'Hapus Baris yang dicatat
                        TBLAdjust.Rows.Remove(RowAdjType)
                    End While
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_AdjustType_WAREHOUSE(TBLSort, TBLAdjust, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal, TBLSetting)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTAAdjType.Dispose()
            DTASetting.Dispose()
            TBLSetting.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_AdjustType_WAREHOUSE(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLDummy As DseLaporanStockInOut.TBL_Adjustment_WHDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal TBLSetting As DseLaporanStockInOut.tbl_setting_adjustmentDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_Adjustment_WHRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Dim ExpDate As Date = #1/1/1950#
        Dim idx As Integer = 0
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Filter TBLDummy, input sisa baris By Primary Stock
            ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & ""
            TBLDummy.DefaultView.RowFilter = ParamStr
            If TBLDummy.DefaultView.Count > 0 Then
                While idx <= TBLDummy.DefaultView.Count - 1
                    RowDummy = TBLDummy.DefaultView(idx).Row
                    'Primary Purchase
                    Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                    'New Data
                    RowSort = TBLSort.NewTblCetakRow
                    With RowSort
                        '.Primary_AutoNumber = RowDummy.Primary
                        .Primary_Stock = RowDummy.Primary_Transaksi
                        .Primary_Product = RowDummy.Primary_Product
                        ._Date = RowDummy._Date
                        .Document_No = RowDummy.Document_No
                        .Remarks = RowDummy.Remarks
                        If String.IsNullOrEmpty(.Remarks) = True Then
                            .Remarks = "-"
                        End If
                        .Code_No = RowDummy.Product_Code
                        .Description = RowDummy.Description
                        ._Decimal = RowDummy._Decimal
                        .Primary_Unit = RowDummy.Primary_Unit
                        .Unit = RowDummy.Unit
                        If String.IsNullOrEmpty(.Unit) = True Then
                            .Unit = "-"
                        End If
                        .Qty = RowDummy.Qty
                        .Batch_No = RowDummy.Batch_No
                        If String.IsNullOrEmpty(.Batch_No) = True Then
                            .Batch_No = "-"
                        End If
                        .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                        If CDate(.Expiry_Date) = #1/1/1950# Then
                            .Expiry_Date = "-"
                        End If
                        .Location = RowDummy.Location
                        If String.IsNullOrEmpty(.Location) = True Then
                            .Location = "-"
                        End If
                        .Rounding = .Rounding - RowDummy._Decimal
                        .SubTotal = ""
                        '--Tambahan Field--
                        .Primary_PO_SO = -1
                        .PO_SO_Number = ""
                        .PrimaryAdjType = RowDummy.PrimaryAdjType
                        .AdjTypeCode = Me.GetAdjTypeCode(TBLSetting, .PrimaryAdjType)
                        .Primary_Supp_Cust = -1
                        .Cust_Supp_Code = ""
                    End With
                    TBLSort.Rows.Add(RowSort)
                    'Input Sub Total
                    Me.InputTBLSubTotal_AdjustType_WAREHOUSE(TBLSubTotal, RowDummy)
                    'Hapus Baris yang sudah di Input ke TBL Sort
                    TBLDummy.Rows.Remove(RowDummy)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_AdjustType_WAREHOUSE(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_Adjustment_WHRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataAdjustType(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByRef TBLPurchase As DseLaporanStockInOut.TBL_PurchaseDataTable, _
    ByVal TBLSales As DseLaporanStockInOut.TBL_SalesDataTable, _
    ByVal TBLAdjust As DseLaporanStockInOut.TBL_AdjustmentDataTable, _
    ByVal TypeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut, _
    ByRef TampilkanSubTotal As Double, ByVal vEndDate As Date, ByVal vStartDate As Date, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double)
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        '---
        Dim DTAAdjType As New DseLaporanStockInOutTableAdapters.TBL_AdjustmentTableAdapter
        Dim RowAdjType As DseLaporanStockInOut.TBL_AdjustmentRow
        Dim TBLDummyAdjType As New DseLaporanStockInOut.TBL_AdjustmentDataTable
        Dim DTASetting As New DseLaporanStockInOutTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim TBLSetting As New DseLaporanStockInOut.tbl_setting_adjustmentDataTable
        Try
            'Database
            DTAAdjType.Connection.ConnectionString = Core_Module.GetConstring
            DTASetting.Connection.ConnectionString = Core_Module.GetConstring
            DTASetting.Fill(TBLSetting)
            If Me.ChkFilterByDate.Checked = True Then
                DTAAdjType.FillBy_TanggalDariSampai_INOUTAdjust(TBLAdjust, vEndDate, vStartDate, Me.TypeStockInOut)
                DTAAdjType.FillBy_TanggalDariSampai_INOUTAdjust(TBLDummyAdjType, vEndDate, vStartDate, Me.TypeStockInOut)
            Else
                DTAAdjType.FillBy_TipeStock(TBLAdjust, Me.TypeStockInOut)
                DTAAdjType.FillBy_TipeStock(TBLDummyAdjType, Me.TypeStockInOut)
            End If

            'TampilkanSubTotal
            If Me.ChkShowSubTotal.Checked = True Then
                TampilkanSubTotal = 1
            End If

            'Hitung Stock
            If TBLAdjust.Rows.Count > 0 Then
                TBLAdjust.DefaultView.RowFilter = Nothing
                TBLAdjust.DefaultView.Sort = Nothing
                TBLAdjust.DefaultView.RowFilter = Me.GetOverallFilter(TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut)
                TBLAdjust.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Product Code] ASC"
                If TBLAdjust.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To TBLAdjust.DefaultView.Count - 1
                        RowAdjType = TBLAdjust.DefaultView(ij).Row
                        'Primary Purchase
                        Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowAdjType.Primary_Transaksi)
                        'List Primary Transaksi Stock Transfer
                        If ListPrimaryStock.Contains(RowAdjType.Primary_Transaksi) = False Then
                            ListPrimaryStock.Add(RowAdjType.Primary_Transaksi)
                        End If
                        'add DS_Cetak
                        RowSort = TBLSort.NewTblCetakRow
                        With RowSort
                            .Primary_AutoNumber = RowAdjType.Primary
                            .Primary_Stock = RowAdjType.Primary_Transaksi
                            .Primary_Product = RowAdjType.Primary_Product
                            ._Date = RowAdjType._Date
                            .Document_No = RowAdjType.Document_No
                            .Remarks = RowAdjType.Remarks
                            If String.IsNullOrEmpty(.Remarks) = True Then
                                .Remarks = "-"
                            End If
                            .Code_No = RowAdjType.Product_Code
                            .Description = RowAdjType.Description
                            ._Decimal = RowAdjType._Decimal
                            .Primary_Unit = RowAdjType.Primary_Unit
                            .Unit = RowAdjType.Unit
                            If String.IsNullOrEmpty(.Unit) = True Then
                                .Unit = "-"
                            End If
                            .Qty = RowAdjType.Qty
                            .Batch_No = RowAdjType.Batch_No
                            If String.IsNullOrEmpty(.Batch_No) = True Then
                                .Batch_No = "-"
                            End If
                            .Expiry_Date = Format(RowAdjType.Expiry_Date, "dd MMM yyyy")
                            If CDate(.Expiry_Date) = #1/1/1950# Then
                                .Expiry_Date = "-"
                            End If
                            .Location = RowAdjType.Location
                            If String.IsNullOrEmpty(.Location) = True Then
                                .Location = "-"
                            End If
                            .Rounding = .Rounding - RowAdjType._Decimal
                            .SubTotal = ""
                            '--Tambahan Field--
                            .Primary_PO_SO = -1
                            .PO_SO_Number = ""
                            .PrimaryAdjType = RowAdjType.PrimaryAdjType
                            .AdjTypeCode = Me.GetAdjTypeCode(TBLSetting, .PrimaryAdjType)
                            .Primary_Supp_Cust = -1
                            .Cust_Supp_Code = ""
                        End With
                        TBLSort.Rows.Add(RowSort)
                        'Input Sub Total
                        Me.InputTBLSubTotal_AdjustType(TBLSubTotal, RowAdjType)
                    Next
                End If
                'All Items
                If Me.ChkShowAllItemsInTrans.Checked = True Then
                    TBLDummyAdjType.DefaultView.RowFilter = ""
                    TBLDummyAdjType.DefaultView.Sort = ""
                    For Each PrimaryStock In ListPrimaryStock
                        Me.GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_AdjustType(TBLSort, TBLDummyAdjType, PrimaryStock, _
                        PrimaryMax, PrimaryMin, TBLSubTotal)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryStock = Nothing
            DTAAdjType.Dispose()
            TBLDummyAdjType.Dispose()
            DTASetting.Dispose()
            TBLSetting.Dispose()
        End Try
    End Sub

    Public Sub GetFilter_TidakSamaDengan_PrimaryAutoNumber_Stock_AdjustType(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable, _
    ByVal TBLDummy As DseLaporanStockInOut.TBL_AdjustmentDataTable, ByVal vPrimaryStock As Double, _
    ByRef PrimaryMax As Double, ByRef PrimaryMin As Double, ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable)
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim RowDummy As DseLaporanStockInOut.TBL_AdjustmentRow
        Dim Filter As String = ""
        Dim ParamStr As String = ""
        Try
            'Masukan sisa item yang ada di TBLStock Per Primary Stock
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = ""
            '1. Dapatkan primary auto number dari tbl cetak
            TBLSort.DefaultView.RowFilter = "[Primary Stock] = " & vPrimaryStock.ToString & ""
            If TBLSort.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                    RowSort = TBLSort.DefaultView(ij).Row
                    If String.IsNullOrEmpty(Filter) = False Then
                        Filter &= " AND [Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    Else
                        Filter = "[Primary] <> " & RowSort.Primary_AutoNumber.ToString & ""
                    End If
                Next
                '2. Input sisa data item ke TBL Sort (Cetak)
                If String.IsNullOrEmpty(Filter) = False Then
                    TBLDummy.DefaultView.RowFilter = ""
                    TBLDummy.DefaultView.Sort = ""
                    ParamStr = "[Primary Transaksi] = " & vPrimaryStock.ToString & " AND " & Filter & ""
                    TBLDummy.DefaultView.RowFilter = ParamStr
                    If TBLDummy.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TBLDummy.DefaultView.Count - 1
                            RowDummy = TBLDummy.DefaultView(ij).Row
                            'Primary Purchase
                            Me.GetPrimaryMinMax(PrimaryMin, PrimaryMax, RowDummy.Primary_Transaksi)
                            'New Data
                            RowSort = TBLSort.NewTblCetakRow
                            With RowSort
                                .Primary_AutoNumber = RowDummy.Primary
                                .Primary_Stock = RowDummy.Primary_Transaksi
                                .Primary_Product = RowDummy.Primary_Product
                                ._Date = RowDummy._Date
                                .Document_No = RowDummy.Document_No
                                .Remarks = RowDummy.Remarks
                                If String.IsNullOrEmpty(.Remarks) = True Then
                                    .Remarks = "-"
                                End If
                                .Code_No = RowDummy.Product_Code
                                .Description = RowDummy.Description
                                ._Decimal = RowDummy._Decimal
                                .Primary_Unit = RowDummy.Primary_Unit
                                .Unit = RowDummy.Unit
                                If String.IsNullOrEmpty(.Unit) = True Then
                                    .Unit = "-"
                                End If
                                .Qty = RowDummy.Qty
                                .Batch_No = RowDummy.Batch_No
                                If String.IsNullOrEmpty(.Batch_No) = True Then
                                    .Batch_No = "-"
                                End If
                                .Expiry_Date = Format(RowDummy.Expiry_Date, "dd MMM yyyy")
                                If CDate(.Expiry_Date) = #1/1/1950# Then
                                    .Expiry_Date = "-"
                                End If
                                .Location = RowDummy.Location
                                If String.IsNullOrEmpty(.Location) = True Then
                                    .Location = "-"
                                End If
                                .Rounding = .Rounding - RowDummy._Decimal
                                .SubTotal = ""
                                '--Tambahan Field--
                                .Primary_PO_SO = -1
                                .PO_SO_Number = ""
                                .PrimaryAdjType = RowDummy.PrimaryAdjType
                                .AdjTypeCode = RowDummy.AdjTypeCode
                                .Primary_Supp_Cust = -1
                                .Cust_Supp_Code = ""
                            End With
                            TBLSort.Rows.Add(RowSort)
                            'Input Sub Total
                            Me.InputTBLSubTotal_AdjustType(TBLSubTotal, RowDummy)
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputTBLSubTotal_AdjustType(ByRef TBLSubTotal As DseLaporanStockInOut.TblTotalDataTable, _
    ByVal RowStock As DseLaporanStockInOut.TBL_AdjustmentRow)
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim Filter As String = ""
        Try
            Filter = "[Primary Transaksi] = " & RowStock.Primary_Transaksi.ToString & " AND [Primary Unit] = " & RowStock.Primary_Unit.ToString & ""
            TBLSubTotal.DefaultView.RowFilter = Filter
            If TBLSubTotal.DefaultView.Count > 0 Then
                RowSubTotal = TBLSubTotal.DefaultView(0).Row
                RowSubTotal.Total_Qty += RowStock.Qty
                Me.GetDecimal_Maximum_Untuk_SubTotal(RowSubTotal._Decimal, RowStock._Decimal)
            Else
                RowSubTotal = TBLSubTotal.NewRow
                RowSubTotal.Primary_Transaksi = RowStock.Primary_Transaksi
                RowSubTotal.Total_Qty = RowStock.Qty
                RowSubTotal._Decimal = RowStock._Decimal
                RowSubTotal.Primary_Unit = RowStock.Primary_Unit
                RowSubTotal.Unit = RowStock.Unit
                TBLSubTotal.Rows.Add(RowSubTotal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Public Sub Cetak(ByVal vEndDate As Date, ByVal vStartDate As Date)
        Dim Report1 As New CrpLaporanStockInOutPurchase_SMESS
        Dim Report2 As New CrpLaporanStockInOutAdjust_SMESS
        Dim Report_AOP As New CrpLaporanStockInOutPurchase_SMESS_WithAOP
        Dim TampilkanSubTotal As Double = 0
        Dim ListPrimaryStock As New List(Of Double)
        Dim PrimaryStock As Double = -1
        Dim SubTotalStr As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DS As New DseLaporanStockInOut
        Dim TBL As DseLaporanStockInOut.TblCetakDataTable = DS.TblCetak
        Dim TBLSort As New DseLaporanStockInOut.TblCetakDataTable
        Dim Row As DseLaporanStockInOut.TblCetakRow
        Dim RowSort As DseLaporanStockInOut.TblCetakRow
        Dim PrimaryMinStockPurchase As Double = Double.MaxValue
        Dim PrimaryMaxStockPurchase As Double = Double.MinValue
        Dim JudulReports As String = ""
        Dim SupplierCustomer As String = ""
        Dim POatoSONumber As String = ""
        Dim LocationWarehouseString As String = ""
        '-------
        Dim TBLSubTotal As New DseLaporanStockInOut.TblTotalDataTable
        Dim RowSubTotal As DseLaporanStockInOut.TblTotalRow
        Dim TBLGRN As New DseLaporanStockInOut.tbltrans_grnDataTable
        '-------
        'Pecah sesuai dengan Tipe Stock Yang dipilih
        '1.
        Dim DTAPurchase As New DseLaporanStockInOutTableAdapters.TBL_PurchaseTableAdapter
        Dim TBLPurchase As New DseLaporanStockInOut.TBL_PurchaseDataTable
        Dim TBLDummyPurchase As New DseLaporanStockInOut.TBL_PurchaseDataTable
        Dim TBLPurchase_WH As New DseLaporanStockInOut.TBL_Purchase_WHDataTable
        Dim TBLDummyPurchase_WH As New DseLaporanStockInOut.TBL_Purchase_WHDataTable
        '2. 
        Dim DTASales As New DseLaporanStockInOutTableAdapters.TBL_SalesTableAdapter
        Dim TBLSales As New DseLaporanStockInOut.TBL_SalesDataTable
        Dim TBLDummySales As New DseLaporanStockInOut.TBL_SalesDataTable
        Dim TBLSales_WH As New DseLaporanStockInOut.TBL_Sales_WHDataTable
        Dim TBLDummySales_WH As New DseLaporanStockInOut.TBL_Sales_WHDataTable
        '3.
        Dim DTAAdjust As New DseLaporanStockInOutTableAdapters.TBL_AdjustmentTableAdapter
        Dim TBLAdjust As New DseLaporanStockInOut.TBL_AdjustmentDataTable
        Dim TBLDummyAdjust As New DseLaporanStockInOut.TBL_AdjustmentDataTable
        Dim TBLAdjust_WH As New DseLaporanStockInOut.TBL_Adjustment_WHDataTable
        Dim TBLDummyAdjust_WH As New DseLaporanStockInOut.TBL_Adjustment_WHDataTable
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'Fill data
            If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                If Me.ChkFilterByWareHouse.Checked = False Then
                    'By Location Code
                    Me.GetDataPurchase(TBLSort, TBLSubTotal, TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                Else
                    'By Warehouse Code
                    Me.GetDataPurchase_WAREHOUSE(TBLSort, TBLSubTotal, TBLPurchase_WH, TBLSales_WH, TBLDummyAdjust_WH, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                End If
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                    'Compare dengan GRN
                    If TBLSort.Rows.Count > 0 Then
                        If Me.CboInPurchase.SelectedIndex <> 2 Then
                            'If Me.ChkShowAllItemsInTrans.Checked = False Then
                            Me.CompareDgnGRN(TBLSort, TBLGRN, PrimaryMaxStockPurchase, PrimaryMinStockPurchase, _
                            TBLSubTotal)
                            'End If
                        Else
                            'Jika In Purchase = Both, isi AOP Number
                            Me.GetAOPNumber_Both(TBLSort, TBLGRN, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                        End If
                    End If
                End If
            ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                If Me.ChkFilterByWareHouse.Checked = False Then
                    'By Location Code
                    Me.GetDataSales(TBLSort, TBLSubTotal, TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                Else
                    'By Warehouse Code 
                    Me.GetDataSales_WAREHOUSE(TBLSort, TBLSubTotal, TBLPurchase_WH, TBLSales_WH, TBLAdjust_WH, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                End If
            ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                If Me.ChkFilterByWareHouse.Checked = False Then
                    'By Location Code 
                    Me.GetDataAdjustType(TBLSort, TBLSubTotal, TBLPurchase, TBLSales, TBLAdjust, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                Else
                    'By Warehouse Code
                    Me.GetDataAdjustType_WAREHOUSE(TBLSort, TBLSubTotal, TBLPurchase_WH, TBLSales_WH, TBLAdjust_WH, Me.TypeStockInOut, _
                    TampilkanSubTotal, vEndDate, vStartDate, PrimaryMaxStockPurchase, PrimaryMinStockPurchase)
                End If
            End If

            'Input Sub Total
            If TampilkanSubTotal = 1 Then
                If TBLSubTotal.Rows.Count > 0 Then
                    For Each RowSort In TBLSort.Rows
                        SubTotalStr = ""
                        TBLSubTotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSort.Primary_Stock.ToString & ""
                        If TBLSubTotal.DefaultView.Count > 0 Then
                            For ij As Integer = 0 To TBLSubTotal.DefaultView.Count - 1
                                RowSubTotal = TBLSubTotal.DefaultView(ij).Row
                                If String.IsNullOrEmpty(SubTotalStr) = False Then
                                    SubTotalStr &= ", " & ObjCD.DigitGrouping(RowSubTotal.Total_Qty, RowSubTotal._Decimal) & " " & RowSubTotal.Unit
                                Else
                                    SubTotalStr = ObjCD.DigitGrouping(RowSubTotal.Total_Qty, RowSubTotal._Decimal) & " " & RowSubTotal.Unit
                                End If
                            Next
                            RowSort.SubTotal = SubTotalStr
                        End If
                    Next
                End If
            End If

            'Sorting 
            If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                TBLSort.DefaultView.RowFilter = ""
                TBLSort.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [AdjTypeCode] ASC, [Code No] ASC"
            Else
                TBLSort.DefaultView.RowFilter = ""
                TBLSort.DefaultView.Sort = "[Date] ASC, [Document No] ASC, [Cust Supp Code] ASC, [PO SO Number] ASC, [Code No] ASC"
            End If
            If TBLSort.DefaultView.Count > 0 Then
                For ij As Integer = 0 To TBLSort.DefaultView.Count - 1
                    RowSort = TBLSort.DefaultView(ij).Row
                    Row = TBL.NewTblCetakRow
                    For ik As Integer = 0 To TBL.Columns.Count - 1
                        Row.Item(ik) = RowSort.Item(ik)
                    Next
                    TBL.Rows.Add(Row)
                Next
            End If

            'Jika Report Kosong berikan Message Box
            If TBL.Rows.Count <= 0 Then
                Windows.Forms.Cursor.Current = Cursors.Arrow
                MsgBox("There is no data.", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Exit Sub
            End If

            Me.SetDecimal(DS.TblCetak)
            'Reports Source
            If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Report2.SetDataSource(DS)
                Report2.SetParameterValue(Report2.Parameter_TampilkanSubTotal.ParameterFieldName, TampilkanSubTotal)
                Report2.SetParameterValue(Report2.Parameter_StartDate.ParameterFieldName, Format(vStartDate, "dd MMM yyyy"))
                Report2.SetParameterValue(Report2.Parameter_EndDate.ParameterFieldName, Format(vEndDate, "dd MMM yyyy"))
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Then
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, CStr("In Adjust").ToUpper)
                ElseIf Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                    Report2.SetParameterValue(Report2.Parameter_Judul.ParameterFieldName, CStr("Out Adjust").ToUpper)
                End If
                If Me.ChkFilterByWareHouse.Checked = False Then
                    LocationWarehouseString = "Location"
                    Report2.SetParameterValue(Report2.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                Else
                    LocationWarehouseString = "Warehouse"
                    Report2.SetParameterValue(Report2.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                End If
                Me.crv1.ReportSource = Report2
            Else
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                    JudulReports = CStr("In Purchase").ToUpper
                    SupplierCustomer = CStr("Supplier")
                    POatoSONumber = CStr("PO Number")
                End If
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                    JudulReports = CStr("Out Purchase").ToUpper
                    SupplierCustomer = CStr("Supplier")
                    POatoSONumber = CStr("PO Number")
                End If
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                    JudulReports = CStr("In Sales Return").ToUpper
                    SupplierCustomer = CStr("Customer")
                    POatoSONumber = CStr("SO Number")
                End If
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                    JudulReports = CStr("Out Sales").ToUpper
                    SupplierCustomer = CStr("Customer")
                    POatoSONumber = CStr("SO Number")
                End If
                If Me.TypeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                    If Me.CboInPurchase.SelectedIndex = 0 Or Me.CboInPurchase.SelectedIndex = 2 Then
                        If Me.ChkShowAOPNo.Checked = True Then
                            Report_AOP.SetDataSource(DS)
                            Report_AOP.SetParameterValue(Report1.Parameter_TampilkanSubTotal.ParameterFieldName, TampilkanSubTotal)
                            Report_AOP.SetParameterValue(Report1.Parameter_StartDate.ParameterFieldName, Format(vStartDate, "dd MMM yyyy"))
                            Report_AOP.SetParameterValue(Report1.Parameter_EndDate.ParameterFieldName, Format(vEndDate, "dd MMM yyyy"))
                            Report_AOP.SetParameterValue(Report1.Parameter_Judul.ParameterFieldName, JudulReports)
                            Report_AOP.SetParameterValue(Report1.Parameter_SupplierCustomer.ParameterFieldName, SupplierCustomer)
                            Report_AOP.SetParameterValue(Report1.Parameter_POSONumber.ParameterFieldName, POatoSONumber)
                            If Me.ChkFilterByWareHouse.Checked = False Then
                                LocationWarehouseString = "Location"
                                Report_AOP.SetParameterValue(Report_AOP.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                            Else
                                LocationWarehouseString = "Warehouse"
                                Report_AOP.SetParameterValue(Report_AOP.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                            End If
                            Me.crv1.ReportSource = Report_AOP
                            ''
                            Windows.Forms.Cursor.Current = Cursors.Arrow
                            Me.crv1.Show()
                            Exit Sub
                        End If
                    End If
                End If
                Report1.SetDataSource(DS)
                Report1.SetParameterValue(Report1.Parameter_TampilkanSubTotal.ParameterFieldName, TampilkanSubTotal)
                Report1.SetParameterValue(Report1.Parameter_StartDate.ParameterFieldName, Format(vStartDate, "dd MMM yyyy"))
                Report1.SetParameterValue(Report1.Parameter_EndDate.ParameterFieldName, Format(vEndDate, "dd MMM yyyy"))
                Report1.SetParameterValue(Report1.Parameter_Judul.ParameterFieldName, JudulReports)
                Report1.SetParameterValue(Report1.Parameter_SupplierCustomer.ParameterFieldName, SupplierCustomer)
                Report1.SetParameterValue(Report1.Parameter_POSONumber.ParameterFieldName, POatoSONumber)
                If Me.ChkFilterByWareHouse.Checked = False Then
                    LocationWarehouseString = "Location"
                    Report1.SetParameterValue(Report1.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                Else
                    LocationWarehouseString = "Warehouse"
                    Report1.SetParameterValue(Report1.Parameter_LocationWarehouse.ParameterFieldName, LocationWarehouseString)
                End If
                Me.crv1.ReportSource = Report1
            End If
            ''
            Windows.Forms.Cursor.Current = Cursors.Arrow
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAPurchase.Dispose()
            TBLPurchase.Dispose()
            TBLDummyPurchase.Dispose()
            TBLPurchase_WH.Dispose()
            TBLDummyPurchase_WH.Dispose()
            DTASales.Dispose()
            TBLSales.Dispose()
            TBLDummySales.Dispose()
            TBLSales_WH.Dispose()
            TBLDummySales_WH.Dispose()
            DTAAdjust.Dispose()
            TBLAdjust.Dispose()
            TBLDummyAdjust.Dispose()
            TBLAdjust_WH.Dispose()
            TBLDummyAdjust_WH.Dispose()
            ListPrimaryStock = Nothing
            ObjCD = Nothing
            TBLSort.Dispose()
            TBLSubTotal.Dispose()
            TBLGRN.Dispose()
        End Try
    End Sub

    Private Function SetDecimal(ByRef TBLSort As DseLaporanStockInOut.TblCetakDataTable) As Boolean
        Dim Row As DseLaporanStockInOut.TblCetakRow
        Dim dt As New DseLaporanStockInOutTableAdapters.DataTable4TableAdapter
        Dim t As New DseLaporanStockInOut.DataTable4DataTable
        Dim r As DseLaporanStockInOut.DataTable4Row
        Try
            dt.Connection.ConnectionString = Core_Module.GetConstring
            dt.Fill(t)
            For Each Row In TBLSort.Rows
                t.DefaultView.RowFilter = "[Primary Product]=" & Row.Primary_Product & " AND [Primary Unit]=" & Row.Primary_Unit 'maksudnya primary multi unit salah penamaan saja.
                If t.DefaultView.Count > 0 Then
                    r = t.DefaultView(0).Row
                    Row._Decimal = r.Decimal_Multi
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ChkFilterByDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFilterByDate.CheckedChanged
        Try
            If Me.ChkFilterByDate.Checked = True Then
                Me.DtpMinExpiryDate.Enabled = True
                Me.DtpMaxExpiryDate.Enabled = True
            Else
                Me.DtpMinExpiryDate.Enabled = False
                Me.DtpMaxExpiryDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboInPurchase_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboInPurchase.SelectedIndexChanged
        Try
            If Me.CboInPurchase.SelectedIndex = 0 Then
                Me.ChkShowAOPNo.Enabled = True
                Me.ChkShowAOPNo.Checked = True
            ElseIf Me.CboInPurchase.SelectedIndex = 1 Then
                Me.ChkShowAOPNo.Enabled = False
                Me.ChkShowAOPNo.Checked = False
            ElseIf Me.CboInPurchase.SelectedIndex = 2 Then
                Me.ChkShowAOPNo.Enabled = True
                Me.ChkShowAOPNo.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkFilterByWareHouse_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkFilterByWareHouse.CheckedChanged
        Try
            If Me.ChkFilterByWareHouse.Checked = True Then
                Me.TabPage1.Text = "Warehouse"
                Me.DgvLocation.Rows.Clear()
            Else
                Me.TabPage1.Text = "Location"
                Me.DgvLocation.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class