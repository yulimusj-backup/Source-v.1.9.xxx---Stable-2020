Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmMasterKondisiStock
    Dim ObjCd As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjPackingListItem As New ClsTransaksiPackingListItem
    Dim ObjBookingInv As New ClsTransaksiBookingInvItem
    Dim ObjDOLossItem As New ClsTransaksiDOLossItem
    Dim ObjDOItem As New ClsTransaksiDOItem
    Dim ObjPO As New ClsTransaksiPOItem
    Dim ObjPurcReq As New ClsTransaksiPR
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjProses As New Core_Module.ClsMasterProses

    Dim TBLBookingInv As New DseTransaksiBookingInv.tbltrans_bookinginv_item_viewDataTable
    Dim TBLGIT As New DseTransaksiPackingList.DataTable22DataTable
    Dim TBLPendingDO As New DseTransaksiDO.pendingdo_view0DataTable
    Dim TBLProses As New Core_Module.DseMasterProses.tblmaster_proses1DataTable
    Dim TBLUnitPrice As New Core_Module.dseMasterProduct.DataTable3DataTable

    Public BackcolorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public FormNo As Integer = 0

    Dim DTA As New dseLaporanStockDetailTableAdapters.stock_detail_viewTableAdapter
    Dim TBL As New dseLaporanStockDetail.stock_detail_viewDataTable
    Dim TBL_POB As New DseTransaksiPO2.tbltrans_po_item9DataTable
    Dim TBL_SOB As New DseTransaksiSO_OrderBalance.tbltrans_so_item1DataTable
    Dim buffer As String = ""

    'aul
    Dim TblPUO2 As New DseTransaksiPR.Tbl_PickUpOrder2DataTable

    Dim PrimaryCodeNo_ShowQty As Double = -1

    Enum EnumPOBSort As Integer
        NoSort
        Asc
        Desc
    End Enum

    'aul
    Dim UnAuthPO As ClsTransaksiPOItem.IncludeUnAuthPO = ClsTransaksiPOItem.IncludeUnAuthPO.ExcludeUnAuthPO
    '///

    Dim SortDirection As EnumPOBSort = EnumPOBSort.NoSort
    Dim SortDirection_Date As EnumPOBSort = EnumPOBSort.NoSort

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub txtCodeNo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Enter
        Try
            Me.buffer = Me.txtCodeNo.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus
        Try
            If Me.txtCodeNo.ReadOnly = False And Me.txtCodeNo.Enabled = True Then
                Me.txtCodeNo.BackColor = Me.BackcolorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                i.BackColorAll = Me.BackcolorAll
                Jml = Me.ObjProduct.GetJml(Me.txtCodeNo.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGBO()
                    Else
                        Me.txtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.KosongkanItem()
                        Me.txtCodeNo.Focus()
                        Me.gboDetail.Text = ""
                        Me.DataGridView1.BringToFront()
                        Me.Stock_detail_viewDataGridView.Visible = False
                    End If
                End If

                'If Me.txtPrimaryCodeNo.Text <> Me.PrimaryCodeNo_ShowQty Then
                '    Me.ShowQtyOnHand()
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            Me.ObjCd.KosongkanControl(Me.Controls, True, False, False)
            Me.txtBookingGIT.Text = "0"
            Me.txtBookingInv.Text = "0"
            Me.txtFreeStock.Text = "0"
            Me.txtGIT.Text = "0"
            Me.txtQtyOnHand.Text = "0"
            Me.TxtPOB.Text = "0"
            Me.TxtSOB.Text = "0"
            Me.txtPendingDO.Text = "0"
            Me.txtAvailable.Text = "0"
            Me.TxtPurcReqBalance.Text = "0"
            Me.TxtMOB.Text = ""
            Me.TxtRawMaterialRequest.Text = "0"

            Me.lblUnit1.Text = "UNIT"
            Me.lblUnit2.Text = "UNIT"
            Me.lblUnit3.Text = "UNIT"
            Me.lblUnit4.Text = "UNIT"
            Me.lblUnit5.Text = "UNIT"
            Me.LblUnit6.Text = "UNIT"
            Me.LblUnit7.Text = "UNIT"
            Me.LblUnit8.Text = "UNIT"
            Me.lblUnit9.Text = "UNIT"
            Me.lblUnit10.Text = "UNIT"
            Me.lblUnitAvailable.Text = "UNIT"
            Me.lblUnitPendingDO.Text = "UNIT"

            Me.txtPrimaryCodeNo.Text = "-1"
            Me.gboDetail.Text = ""
            Me.Stock_detail_viewDataGridView.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGBO()
        Dim mDecimal As Integer = 0

        Try
            Cursor.Current = Cursors.WaitCursor

            mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            'set unit
            Me.lblUnit1.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
            Me.lblUnit2.Text = Me.lblUnit1.Text
            Me.lblUnit3.Text = Me.lblUnit1.Text
            Me.lblUnit4.Text = Me.lblUnit1.Text
            Me.lblUnit5.Text = Me.lblUnit1.Text
            Me.LblUnit6.Text = Me.lblUnit1.Text
            Me.LblUnit7.Text = Me.lblUnit1.Text
            Me.LblUnit8.Text = Me.lblUnit1.Text
            Me.lblUnit9.Text = Me.lblUnit1.Text
            Me.lblUnit10.Text = Me.lblUnit1.Text
            Me.lblUnitPendingDO.Text = Me.lblUnit1.Text
            Me.lblUnitAvailable.Text = Me.lblUnit1.Text

            'isikan data stock
            Me.txtGIT.Text = Me.ObjCd.DigitGrouping(Me.ObjPackingListItem.GetJmlAktif(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtQtyOnHand.Text = Me.ObjCd.DigitGrouping(Me.ObjProduct.GetJmlByPrimary(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtBookingInv.Text = Me.ObjCd.DigitGrouping(Me.ObjBookingInv.GetJml(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtBookingGIT.Text = Me.ObjCd.DigitGrouping(0, mDecimal)
            Me.TxtMOB.Text = Me.ObjCd.DigitGrouping(Me.ObjMO.GetMOB(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.TxtRawMaterialRequest.Text = Me.ObjCd.DigitGrouping(Me.ObjMO.GetRawMaterialRequest(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtFreeStock.Text = Me.ObjCd.DigitGrouping(CDec(Me.txtQtyOnHand.Text) - CDec(Me.txtBookingInv.Text) - CDec(Me.TxtRawMaterialRequest.Text), mDecimal)
            Me.txtPendingDO.Text = Me.ObjCd.DigitGrouping(Me.ObjDOItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text) + _
            Me.ObjDOLossItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtAvailable.Text = Me.ObjCd.DigitGrouping(CDec(Me.txtFreeStock.Text) - CDec(Me.txtPendingDO.Text), mDecimal)
            Me.TxtPOB.Text = Me.ObjCd.DigitGrouping(Me.ObjPO.GetOrderBalance(Me.txtPrimaryCodeNo.Text, Me.UnAuthPO), mDecimal)
            Me.TxtSOB.Text = Me.ObjCd.DigitGrouping(Me.ObjSO.GetSalesOrderBalance(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.TxtPurcReqBalance.Text = Me.ObjCd.DigitGrouping(Me.ObjPurcReq.GetOrderBalance_Total2(Me.txtPrimaryCodeNo.Text), mDecimal)

            'kosongkan gbo
            Me.gboDetail.Text = ""
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.DataGridView1.BringToFront()
            Me.IsiUnitPrice()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackcolorAll
            Me.txtCodeNo.Text = Me.txtCodeNo.Text.Trim

            If Me.txtCodeNo.Text <> "" Then
                If Me.buffer <> Me.txtCodeNo.Text Then
                    Cursor.Current = Cursors.WaitCursor

                    If Me.ObjProduct.GetJml(Me.txtCodeNo.Text) = 1 Then
                        Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGBO()
                    Else
                        If Me.txtCodeNo.Text <> "" Then
                            i.txtCari.Text = Me.txtCodeNo.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryProduct > 0 Then
                            Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                            Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                            Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                            Me.SetGBO()
                        Else
                            Me.txtCodeNo.Clear()
                            Me.txtPrimaryCodeNo.Text = "-1"
                            Me.KosongkanItem()
                            Me.txtCodeNo.Focus()
                            Me.gboDetail.Text = ""
                            Me.DataGridView1.BringToFront()
                            Me.Stock_detail_viewDataGridView.Visible = False
                        End If
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.KosongkanItem()
                Me.gboDetail.Text = ""
                Me.DataGridView1.BringToFront()
                Me.Stock_detail_viewDataGridView.Visible = False
                Me.IsiUnitPrice()
            End If

            If Me.txtPrimaryCodeNo.Text <> Me.PrimaryCodeNo_ShowQty Then
                Me.ShowQtyOnHand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackcolorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.txtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.SetGBO()
            Else
                Me.txtCodeNo.Clear()
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
                Me.txtCodeNo.Focus()
                Me.gboDetail.Text = ""
                Me.DataGridView1.BringToFront()
                Me.Stock_detail_viewDataGridView.Visible = False
            End If

            If Me.txtPrimaryCodeNo.Text <> Me.PrimaryCodeNo_ShowQty Then
                Me.ShowQtyOnHand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus
        Try
            If Me.txtCodeNo.ReadOnly = False And Me.txtCodeNo.Enabled = True Then
                Me.txtCodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim mDecimal As Integer = 0
        Try
            If Me.txtPrimaryCodeNo.Text <> -1 Then
                mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                Me.TxtPurcReqBalance.Text = Me.ObjCd.DigitGrouping(Me.ObjPurcReq.GetOrderBalance_Total2(Me.txtPrimaryCodeNo.Text), mDecimal)
                Me.ObjPurcReq.IsiDGVCekStock(Me.DgvPurcReq, Me.txtPrimaryCodeNo.Text, Me, TblPUO2)
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.UpdateSetting_ChkRMReq()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            Me.Text = "Check Stock"
            Me.Stock_detail_viewDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                End If
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItemPendingDO.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItemSOB.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
           Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowWONo, Me.UserCode)
            If Dummy = "1" Then
                Me.chkShowWONo.Checked = True
            Else
                Me.chkShowWONo.Checked = False
            End If
            LoadData_CheckChange = True
            Me.ColNo.Visible = False
            Me.ColRMReq_No.Visible = False
            Me.ColPurcReq_PickUpOrder.Visible = False
            Me.ColPurcReq_BalanceQtyUsing.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbQtyOnHand_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbQtyOnHand.LinkClicked
        Try
            Me.ShowQtyOnHand()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowQtyOnHand()
        Dim rowDGV As DataGridViewRow
        Dim SecNo As Boolean = True
        Dim LotNo As Boolean = False
        Dim Location As Boolean = True
        Dim ExpiryDate As Boolean = True
        Dim col As DataGridViewColumn
        Dim rowStock As dseLaporanStockDetail.stock_detail_viewRow
        Dim DefPrec As Integer = 0
        Dim KontrolLokasi, KontrolBatch, KontrolExpiry As Boolean

        Try
            Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'TODO: This line of code loads data into the 'DseLaporanStockDetail.stock_detail_view' table. You can move, or remove it, as needed.
            If Me.txtCodeNo.Text <> "" Then
                DTA.FillBy_PrimaryProduct(TBL, Me.txtPrimaryCodeNo.Text)
            Else
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If
            KontrolBatch = Me.ObjProduct.ApakahBatchNoControlled(Me.txtPrimaryCodeNo.Text)
            KontrolExpiry = Me.ObjProduct.ApakahExpiryDateControlled(Me.txtPrimaryCodeNo.Text)
            KontrolLokasi = Me.ObjProduct.ApakahLocationControlled(Me.txtPrimaryCodeNo.Text)

            If KontrolExpiry = True Then
                TBL.DefaultView.Sort = "[Expiry Date], [Lot No], Location"
            Else
                TBL.DefaultView.Sort = "[Lot No], [Expiry Date], Location"
            End If

            Me.Stock_detail_viewDataGridView.DataSource = TBL.DefaultView
            ' Me.Stock_detail_viewDataGridView.Columns(TBL.QtyColumn.ColumnName).DefaultCellStyle.Format = "N" + DefPrec.ToString

            For Each rowDGV In Me.Stock_detail_viewDataGridView.Rows
                If CStr(rowDGV.Cells(TBL.Lot_NoColumn.ColumnName).Value) <> "" Then
                    LotNo = True
                End If
                If CDate(rowDGV.Cells(TBL.Expiry_DateColumn.ColumnName).Value) <> #1/1/1950# Then
                    ExpiryDate = False
                End If
                If CStr(rowDGV.Cells(TBL.LocationColumn.ColumnName).Value) <> "No Location" Then
                    Location = False
                End If
            Next

            If KontrolBatch = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Lot_NoColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.Lot_NoColumn.ColumnName).Visible = False
            End If
            If KontrolExpiry = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = False
            End If
            If KontrolLokasi = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.LocationColumn.ColumnName).Visible = True
                Me.Stock_detail_viewDataGridView.Columns(TBL.WarehouseColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.LocationColumn.ColumnName).Visible = False
                Me.Stock_detail_viewDataGridView.Columns(TBL.WarehouseColumn.ColumnName).Visible = False
            End If


            'format dgv
            Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).DefaultCellStyle.Format = "dd MMM yyyy"
            Me.Stock_detail_viewDataGridView.Columns(0).DefaultCellStyle.Format = "N" + Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text).ToString
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_Storage_AreaColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_UnitColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_WarehouseColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.DecimalColumn.ColumnName).Visible = False
            If KontrolExpiry = False Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = False
            End If
            For Each col In Me.Stock_detail_viewDataGridView.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.Stock_detail_viewDataGridView.Columns(TBL.QtyColumn.ColumnName).DefaultCellStyle.Alignment = _
            DataGridViewContentAlignment.MiddleRight

            Me.Stock_detail_viewDataGridView.Visible = True
            Me.Stock_detail_viewDataGridView.BringToFront()
            Me.gboDetail.Text = "Qty On Hand"

            Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.PrimaryCodeNo_ShowQty = Me.PrimaryCodeNo_ShowQty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub llbBookingGIT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBookingGIT.LinkClicked
    '    Dim col As DataGridViewColumn

    '    Try
    '        Me.ObjBookingGIT.IsiDgvCheckStock(Me.dgvItem, Me.TBLBookingGIT, Me.txtPrimaryCodeNo.Text)
    '        Me.dgvItem.BringToFront()
    '        Me.Stock_detail_viewDataGridView.Visible = False
    '        Me.gboDetail.Text = "Booking GIT"

    '        'format tampilan
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        For Each col In Me.dgvItem.Columns
    '            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    '        Next
    '        Me.dgvItem.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub llbBookingStockInv_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBookingStockInv.LinkClicked
        Dim row As DataGridViewRow
        Dim col As DataGridViewColumn

        Try
            Me.dgvItem.BringToFront()
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.gboDetail.Text = "Booking Inventory"
            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                Me.ObjBookingInv.IsiDgv(Me.dgvItem, Me.TBLBookingInv, Me.txtPrimaryCodeNo.Text)
            End If

            'format tampilan
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each col In Me.dgvItem.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.dgvItem.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'Me.dgvItem.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            For Each row In Me.dgvItem.Rows
                'If row.Cells(9).Value <> -1 Then
                'row.Cells(9).Value = Me.ObjCd.DigitGrouping(row.Cells(9).Value)
                'End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbGIT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbGIT.LinkClicked
        Dim col As DataGridViewColumn

        Try
            Me.dgvItem.BringToFront()
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.gboDetail.Text = "Goods In Transit"

            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                Me.ObjPackingListItem.IsiDgvCheckStock(Me.dgvItem, Me.TBLGIT, Me.txtPrimaryCodeNo.Text)
            End If

            'format tampilan
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each col In Me.dgvItem.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.dgvItem.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub lblPendingDO_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblPendingDO.LinkClicked
        Dim DTA As New DseTransaksiDOTableAdapters.pendingdo_view0TableAdapter
        Dim DTA1 As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_pending_view0TableAdapter
        Dim TBL As New DseTransaksiDOLoss.tbltrans_doloss_pending_view0DataTable
        Dim rowPending As DseTransaksiDO.pendingdo_view0Row
        Dim rowDOLoss As DseTransaksiDOLoss.tbltrans_doloss_pending_view0Row

        Try
            Me.dgvItemPendingDO.DataSource = Nothing
            Me.dgvItemPendingDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Me.gboDetail.Text = "Pending DO"
            Me.dgvItemPendingDO.BringToFront()
            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If

            If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                DTA.Fill(Me.TBLPendingDO)
                DTA1.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(Me.TBLPendingDO, Me.txtPrimaryCodeNo.Text)
                DTA1.FillBy_PrimaryProduct(TBL, Me.txtPrimaryCodeNo.Text)
            End If

            'tambahkan baris pending do ke dalam laporan
            For Each rowDOLoss In TBL.Rows
                rowPending = Me.TBLPendingDO.Newpendingdo_view0Row
                With rowPending
                    .DO_Number = rowDOLoss.DO_Number
                    .Product_Code = rowDOLoss.Product_Code
                    .Description = rowDOLoss.Description
                    .Sales_Order_No = rowDOLoss.Sales_Order_Number
                    .Batch_No = rowDOLoss.Batch_No
                    .Expiry_Date = rowDOLoss.Expiry_Date
                    .Total_Qty = rowDOLoss.Total_Qty
                    .Unit_Code = rowDOLoss.Unit_Code
                    .Secondary_Qty = rowDOLoss.Secondary_Qty
                    .Secondary_Unit = rowDOLoss.Secondary_Unit
                    .Remarks = rowDOLoss.Remarks
                End With
                Me.TBLPendingDO.Rows.Add(rowPending)
            Next

            For Each rowPending In Me.TBLPendingDO.Rows
                rowPending.Quantity = Me.ObjCd.DigitGrouping(rowPending.Total_Qty, rowPending._Decimal)
            Next

            'format tampilan
            Me.dgvItemPendingDO.DataSource = Me.TBLPendingDO.DefaultView
            Dim row As DataGridViewRow
            Dim AdaBatch As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaSecondary As Boolean = False
            If Me.dgvItemPendingDO.Columns.Count > 0 Then
                Me.dgvItemPendingDO.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.dgvItemPendingDO.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If

            For Each row In Me.dgvItemPendingDO.Rows
                If row.Cells(4).Value <> "" Then
                    AdaBatch = True
                End If
                If row.Cells(5).Value <> #1/1/1950# Then
                    AdaExpiry = True
                End If
                If row.Cells(9).Value <> -1 Then
                    AdaSecondary = True
                End If
            Next

            If AdaBatch = False Then
                Me.dgvItemPendingDO.Columns(4).Visible = False
            Else
                Me.dgvItemPendingDO.Columns(4).Visible = True
            End If
            If AdaExpiry = False Then
                Me.dgvItemPendingDO.Columns(5).Visible = False
            Else
                Me.dgvItemPendingDO.Columns(5).Visible = True
            End If
            If AdaSecondary = True Then
                Me.dgvItemPendingDO.Columns(9).Visible = True
                Me.dgvItemPendingDO.Columns(10).Visible = True
            Else
                Me.dgvItemPendingDO.Columns(9).Visible = False
                Me.dgvItemPendingDO.Columns(10).Visible = False
            End If
            Me.dgvItemPendingDO.Columns(6).Visible = False
            Me.dgvItemPendingDO.Columns(12).Visible = False
            Me.dgvItemPendingDO.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            Me.dgvItemPendingDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub IsiUnitPrice()
        Try
            If Me.txtPrimaryCodeNo.Text <> "-1" And Me.txtPrimaryCodeNo.Text <> "" Then
                Me.cboUnitPrice.Enabled = True
                Me.ObjProduct.IsiTBLUnitPrice(Me.txtPrimaryCodeNo.Text, Me.TBLUnitPrice)
                Me.cboUnitPrice.DataSource = Me.TBLUnitPrice.DefaultView
                Me.cboUnitPrice.DisplayMember = "Currency Code"
                If Me.cboUnitPrice.Items.Count > 0 Then
                    Me.cboUnitPrice.SelectedIndex = -1
                    Me.cboUnitPrice.SelectedIndex = 0
                End If
            Else
                Me.TxtUnitPrice.Text = 0
                Me.cboUnitPrice.Enabled = False
                Me.cboUnitPrice.DataSource = Nothing
                Me.cboUnitPrice.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbPOB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbPOB.LinkClicked
        Try
            Me.gboDetail.Text = "Purchase Order Balance"
            'dgv bring to front
            Me.dgvItemPOB.BringToFront()

            If Me.txtCodeNo.Text = "" Then
                Me.dgvItemPOB.DataSource = Nothing
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                'isikan dgv POB
                If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                    Me.txtPrimaryCodeNo.Text = "-1"
                End If
                Me.ObjPO.IsiDGVPOB(Me.txtPrimaryCodeNo.Text, Me.dgvItemPOB, Me.TBL_POB, Me.UnAuthPO)
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemPOB_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItemPOB.ColumnHeaderMouseClick
        Dim IndexPOB As Integer = 7
        Dim IndexPOBStr As Integer = 11
        Dim IndexPOB_Date As Integer = 10
        Dim IndexPOB_Date_Str As Integer = 12

        Try
            Me.dgvItemPOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemPOB.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If e.ColumnIndex = IndexPOBStr Then
                If Me.SortDirection = EnumPOBSort.Desc Or Me.SortDirection = EnumPOBSort.NoSort Then
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection = EnumPOBSort.Asc
                Else
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection = EnumPOBSort.Desc
                End If
                Me.SortDirection_Date = EnumPOBSort.NoSort
            ElseIf e.ColumnIndex = IndexPOB_Date_Str Then
                If Me.SortDirection_Date = EnumPOBSort.Desc Or Me.SortDirection_Date = EnumPOBSort.NoSort Then
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection_Date = EnumPOBSort.Asc
                Else
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection_Date = EnumPOBSort.Desc
                End If
                Me.SortDirection = EnumPOBSort.NoSort
            Else
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If

            Me.dgvItemPOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSOB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSOB.LinkClicked
        Try
            Me.gboDetail.Text = "Sales Order Balance"
            'dgv bring to front
            Me.dgvItemSOB.BringToFront()
            Me.dgvItemSOB.Rows.Clear()

            'isikan dgv POB
            If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                Me.txtPrimaryCodeNo.Text = "-1"
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If

            If Me.dgvItemSOB.Rows.Count > 0 Then
                Me.dgvItemSOB.Rows.Clear()
            End If

            Me.ObjSO.IsiDGVSalesOrderBalance(Me.txtPrimaryCodeNo.Text, Me.TBL_SOB, Me.dgvItemSOB)
            Me.SortDirection = EnumPOBSort.NoSort
            Me.SortDirection_Date = EnumPOBSort.NoSort
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSOB_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemSOB.CellContentClick
        Dim IndexPOB As Integer = 9
        Dim IndexPOBStr As Integer = 11
        Dim IndexPOB_Date As Integer = 10
        Dim IndexPOB_Date_Str As Integer = 12

        Dim ij As New FrmTransaksiSO_SetDelivery

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColSOB_PromiseToDeliver.Index Then
                If Me.dgvItemSOB.Item(Me.ColSOB_PromiseToDeliver.Index, e.RowIndex).Value <> "-" Then
                    ij.vPrimarySO = Me.dgvItemSOB.Item(Me.ColSOB_PrimarySO.Index, e.RowIndex).Value
                    ij.VIdentifier = Me.dgvItemSOB.Item(Me.ColSOB_IdentifierSOItem.Index, e.RowIndex).Value
                    ij.vPrimaryCodeNo = Me.txtPrimaryCodeNo.Text
                    ij.vPrimaryUnitDefault = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                    ij.vPrimaryUnitMulti = Me.dgvItemSOB.Item(Me.ColSOB_PrimaryUnitMulti.Index, e.RowIndex).Value
                    ij.FrmRef = Me
                    ij.vReadOnly = True
                    ij.ShowDialog()
                End If
            End If

            Me.dgvItemSOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemSOB.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If e.ColumnIndex = IndexPOBStr Then
                If Me.SortDirection = EnumPOBSort.Desc Or Me.SortDirection = EnumPOBSort.NoSort Then
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection = EnumPOBSort.Asc
                Else
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection = EnumPOBSort.Desc
                End If
                Me.SortDirection_Date = EnumPOBSort.NoSort
            ElseIf e.ColumnIndex = IndexPOB_Date_Str Then
                If Me.SortDirection_Date = EnumPOBSort.Desc Or Me.SortDirection_Date = EnumPOBSort.NoSort Then
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection_Date = EnumPOBSort.Asc
                Else
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection_Date = EnumPOBSort.Desc
                End If
                Me.SortDirection = EnumPOBSort.NoSort
            Else
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If

            Me.dgvItemSOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub cboUnitPrice_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUnitPrice.SelectedValueChanged
        Dim rowUnitPrice As Core_Module.dseMasterProduct.DataTable3Row

        Try
            If Me.cboUnitPrice.SelectedIndex <> -1 Then
                If Me.cboUnitPrice.Enabled = True And Me.txtPrimaryCodeNo.Text <> "-1" Then
                    rowUnitPrice = Me.TBLUnitPrice.DefaultView(Me.cboUnitPrice.SelectedIndex).Row
                    Me.TxtUnitPrice.Text = Me.ObjCd.DigitGrouping(ObjCd.SetDecimal(rowUnitPrice.Unit_Price, _
                    rowUnitPrice.Default_Precision))
                End If
            Else
                Me.TxtUnitPrice.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbPurcReqBalance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbPurcReqBalance.LinkClicked
        Try
            Me.gboDetail.Text = "Purchase Request Balance"
            'dgv bring to front
            Me.DgvPurcReq.BringToFront()

            If Me.txtCodeNo.Text = "" Then
                If Me.DgvPurcReq.RowCount > 0 Then
                    Me.DgvPurcReq.Rows.Clear()
                End If
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                'isikan dgv POB
                If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                    Me.txtPrimaryCodeNo.Text = "-1"
                End If
                Me.ObjPurcReq.IsiDGVCekStock(Me.DgvPurcReq, Me.txtPrimaryCodeNo.Text, Me, Me.TblPUO2)
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul
    Private Sub ChkIncludeUnAuthPO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkIncludeUnAuthPO.CheckedChanged
        Dim mDecimal As Integer = 0
        Try
            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                If Me.ChkIncludeUnAuthPO.Checked = True Then
                    Me.UnAuthPO = ClsTransaksiPOItem.IncludeUnAuthPO.IncludeUnAuthPO
                    If Me.dgvItemPOB.RowCount > 0 Then
                        Me.dgvItemPOB.Columns(13).Visible = True
                    End If
                Else
                    Me.UnAuthPO = ClsTransaksiPOItem.IncludeUnAuthPO.ExcludeUnAuthPO
                    If Me.dgvItemPOB.RowCount > 0 Then
                        Me.dgvItemPOB.Columns(13).Visible = False
                    End If
                End If

                mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                Me.ObjPO.IsiDGVPOB(Me.txtPrimaryCodeNo.Text, Me.dgvItemPOB, Me.TBL_POB, Me.UnAuthPO)
                Me.TxtPOB.Text = Me.ObjCd.DigitGrouping(Me.ObjPO.GetOrderBalance(Me.txtPrimaryCodeNo.Text, Me.UnAuthPO), mDecimal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul
    Private Sub DgvPurcReq_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvPurcReq.CellContentClick
        Dim ij As New FrmMasterKondisiStock_PickUpOrderList
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiPR.Tbl_PickUpOrderRow
        Dim vPrimary As Double = -1

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColPurcReq_PickUpOrder.Index Then
                'jika hanya ada dua atau lebih pick up dalam satu transaksi purch req maka tampilkan form pick up list
                If Me.DgvPurcReq.Item(Me.ColPurcReq_PickUpOrder.Index, e.RowIndex).Value = "Multiple" And _
                Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPickUpOrder.Index, e.RowIndex).Value = -1 Then
                    ij.vPrimaryCodeNo = Me.txtPrimaryCodeNo.Text
                    ij.VPrimaryPR = Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPR.Index, e.RowIndex).Value
                    ij.vPrimaryPrItemDel = Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPrItemDel.Index, e.RowIndex).Value
                    ij.vTblPuo2 = TblPUO2

                    Dim j As Windows.Forms.Form

                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, FrmMasterKondisiStock_PickUpOrderList)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.Show()
                            Exit Sub
                        End If
                    Next

                    Dim i As New FrmMasterKondisiStock_PickUpOrderList
                    i.vPrimaryCodeNo = Me.txtPrimaryCodeNo.Text
                    i.VPrimaryPR = Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPR.Index, e.RowIndex).Value
                    i.vPrimaryPrItemDel = Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPrItemDel.Index, e.RowIndex).Value
                    i.vTblPuo2 = TblPUO2
                    ' Increment the caption counter.
                    i.UserCode = Me.UserCode.ToString
                    ' Set the caption to be unique.
                    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                    i.MdiParent = Me.ParentForm
                    i.Show()

                    'jika hanya ada satu pick up dalam satu transaksi purch req maka langsung tampilkan form pick up
                ElseIf Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPickUpOrder.Index, e.RowIndex).Value <> -2 Then
                    If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColPurcReq_PickUpOrder.Index Then
                        vPrimary = Me.DgvPurcReq.Item(Me.ColPurcReq_PrimaryPickUpOrder.Index, e.RowIndex).Value
                        Dim ijk As FrmTransaksiPR_PickUpOrder2
                        Dim j As Windows.Forms.Form

                        For Each j In Me.ParentForm.MdiChildren
                            ijk = TryCast(j, FrmTransaksiPR_PickUpOrder2)
                            If IsNothing(ijk) = False Then
                                ijk.BringToFront()
                                ijk.Focus()
                                ijk.ShowTransaction(vPrimary)
                                Exit Sub
                            End If
                        Next

                        Dim i As New FrmTransaksiPR_PickUpOrder2
                        ' Increment the caption counter.
                        i.UserCode = Me.UserCode.ToString
                        ' Set the caption to be unique.
                        i.BAckColorAll = Core_Module.GetSMESSBackColorAll
                        i.MdiParent = Me.ParentForm
                        i.Show()
                        i.ShowTransaction(vPrimary)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowD = Nothing
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_itemTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_mo_itemRow
        Dim UnitCode As String = Me.lblUnit1.Text
        Dim DTA1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item4TableAdapter
        Dim TBL1 As New DseTrans_ManufacturingOrder.tbltrans_mo_item4DataTable
        Dim row1 As DseTrans_ManufacturingOrder.tbltrans_mo_item4Row
        Dim mDecimal As Integer = 0

        Try
            'Me.Text = Me.Text + " - " + Me.ObjProduct.GetProductCode(Me.vPrimaryProduct)
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTA1.Connection.ConnectionString
            Me.ObjProses.FillTabelProses(Me.TBLProses)
            DTA.FillBy_PrimaryProduct(TBL, Me.txtPrimaryCodeNo.Text)

            Me.dgvDataMOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDataMOB.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvDataMOB.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.ColMODate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColTargetFinishDate.DefaultCellStyle.Format = "dd MMM yyyy"
            mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            Me.ColBalanceQty.DefaultCellStyle.Format = "N" + mDecimal.ToString

            For Each row In TBL.Rows
                Me.dgvDataMOB.Rows.Add()

                rowDGV = Me.dgvDataMOB.Rows(Me.dgvDataMOB.Rows.Count - 1)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDataMOB.RowCount
                rowDGV.Cells(Me.ColBalanceQty.Index).Value = row.Balance_Qty
                rowDGV.Cells(Me.ColJONo.Index).Value = row.JO_Number
                rowDGV.Cells(Me.ColMODate.Index).Value = row.MO_Date
                rowDGV.Cells(Me.ColMONo.Index).Value = row.MO_Number
                rowDGV.Cells(Me.ColProgress.Index).Value = "Show"
                rowDGV.Cells(Me.ColTargetFinishDate.Index).Value = #1/1/1950#
                rowDGV.Cells(Me.ColUnitCode.Index).Value = UnitCode
                rowDGV.Cells(Me.ColIdentifierJO.Index).Value = row.Identifier_JO
                rowDGV.Cells(Me.ColIdentifierWO.Index).Value = row.Identifier_WO
                rowDGV.Cells(Me.ColWONo.Index).Value = row.WO_Number
                rowDGV.Cells(Me.ColProses.Index).Value = Me.GetProses(row.Primary_Process)
            Next

            Me.dgvDataMOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvDataMOB.BringToFront()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetProses(ByVal vPrimaryProses As Double) As String
        Dim rowProses As Core_Module.DseMasterProses.tblmaster_proses1Row
        Try
            For Each rowProses In Me.TBLProses.Rows
                If rowProses.Primary = vPrimaryProses Then
                    Return rowProses.Code
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class