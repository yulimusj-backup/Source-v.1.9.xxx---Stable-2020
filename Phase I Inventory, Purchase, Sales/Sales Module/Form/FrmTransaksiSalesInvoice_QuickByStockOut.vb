Option Explicit On
Option Strict Off

Public Class FrmTransaksiSalesInvoice_QuickByStockOut
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjCustomer As New Core_Module.clsMasterCustomer

    Dim rowIndex As Double = -1
    Dim EraseIndex_C As Integer = -1
    Dim EraseIndex_SG As Integer = -1
    Dim EraseIndex_SA As Integer = -1
    Dim EraseIndex_MS As Integer = -1
    Dim EraseIndex_CT As Integer = -1
    Dim EraseIndex_DO As Integer = -1
    Dim EraseIndex_SOS As Integer = -1
    Dim buffer As String = ""

    Dim tbl As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerDataTable

    Public Enum EnumTipeStockInOut As Integer
        StockIn_Purchase
        StockIn_SalesReturn
        StockIn_Adjust
        StockOut_Sales
        StockOut_PurchaseReturn
        StockOut_Adjust
        StockIn_Process
        StockIn_BackProcess
        StockOut_Process
    End Enum

    ''1. sebelum import data, perhatikan sisa no faktur pajak, jangan sampai tidak cukup;.
    ''2. group data berdasarkan: customer, ship to customer, ppn/tidak/, nilai % pajak, currency
    Private Sub FrmTransaksiSalesInvoice_QuickByStockOut_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvHeader.MultiSelect = False
            Me.DgvItem.MultiSelect = False
            Me.Panel1.Visible = False

            Me.DgvHeader.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCust.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCustomerGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvFilter_CreatedBy.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvOutSales.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvDO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.DgvHeader.AllowUserToAddRows = False
            Me.DgvHeader.AllowUserToDeleteRows = False
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvHeader.AllowUserToDeleteRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboOutSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboOutSales.CheckedChanged
        Try
            Me.SetFormByRBO(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboDO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboDO.CheckedChanged
        Try
            Me.SetFormByRBO(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFormByRBO(ByVal vRboOutSales As Boolean)

        Try
            If vRboOutSales = True Then
                Me.TabPage6.Text = "Stock Out Sales"
                Me.DgvOutSales.BringToFront()
                Me.GroupBox3.Text = "Stock Out Sales Created By"
                Me.DgvDO.Rows.Clear()
            Else
                Me.TabPage6.Text = "Delivery Order"
                Me.DgvDO.BringToFront()
                Me.GroupBox3.Text = "Delivery Order Created By"
                Me.DgvOutSales.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll.LinkClicked

        Try
            Me.CheckOrUnCheckCreatedInvoice(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll.LinkClicked

        Try
            Me.CheckOrUnCheckCreatedInvoice(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckOrUnCheckCreatedInvoice(ByVal vCheck As Boolean)
        Dim rd As DataGridViewRow

        Try
            For i As Integer = 0 To Me.DgvHeader.Rows.Count - 1
                rd = Me.DgvHeader.Rows(i)
                rd.Cells(Me.ColCreate.Index).Value = vCheck
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            Me.GetDataFilter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetDataFilter()
        Dim dta As New DseTransaksiSalesInvoice_QuickByStockOutTableAdapters.tbl_headerTableAdapter
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerRow
        Dim rd As DataGridViewRow
        Dim DateFrom_OutSales As Date = Now
        Dim DateTo_OutSales As Date = Now
        Dim DateFrom_Do As Date = Now
        Dim DateTo_DO As Date = Now

        Try
            If Me.RboOutSales.Checked = True Then
                DateFrom_OutSales = Me.DtpFrom.Value.Date
                DateTo_OutSales = Me.DtpTo.Value.Date
                DateFrom_Do = #1/1/1950#
                DateTo_DO = #1/1/2200#
            Else
                DateFrom_OutSales = #1/1/1950#
                DateTo_OutSales = #1/1/2200#
                DateFrom_Do = Me.DtpFrom.Value.Date
                DateTo_DO = Me.DtpTo.Value.Date
            End If

            tbl.Rows.Clear()
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, EnumTipeStockInOut.StockOut_Sales, DateFrom_OutSales, DateFrom_Do, DateTo_OutSales, DateTo_DO)

            tbl.DefaultView.RowFilter = "[primary invoice] <> -1"
            While tbl.DefaultView.Count > 0
                row = tbl.DefaultView(0).Row
                tbl.Rows.Remove(row)
            End While

            Me.FilterByDgv(tbl)
            Me.GetShipToAddress(tbl)

            tbl.DefaultView.RowFilter = ""
            tbl.DefaultView.Sort = "[stockout sales date], [stockout sales no]"
            Me.DgvItem.Rows.Clear()
            Me.DgvHeader.Rows.Clear()
            Me.DgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To tbl.DefaultView.Count - 1
                row = tbl.DefaultView(i).Row
                With row
                    rd = Me.DgvHeader.Rows(Me.DgvHeader.Rows.Add)
                    rd.Cells(ColNo.Index).Value = rd.Index + 1
                    rd.Cells(ColCreate.Index).Value = False
                    rd.Cells(ColStockOut_Number.Index).Value = .stockout_sales_no
                    rd.Cells(ColStockOut_Date.Index).Value = .stockout_sales_date
                    rd.Cells(ColDONo.Index).Value = .do_no
                    rd.Cells(ColDoDate.Index).Value = .do_date
                    rd.Cells(ColCustomerCode.Index).Value = .cust_code
                    rd.Cells(ColCustomerDesc.Index).Value = .cust_name
                    rd.Cells(ColCustomerShipTo.Index).Value = .cust_ship_to_code
                    rd.Cells(ColShipToAddress.Index).Value = .Ship_to_address
                    rd.Cells(ColRemarks.Index).Value = .Remarks_OutSales
                    rd.Cells(ColCreatedBy.Index).Value = .Created_By_OutSales
                    rd.Cells(ColCreatedDate.Index).Value = .Created_Date_OutSales
                    rd.Cells(ColCreatedByDO.Index).Value = .Created_By_DO
                    rd.Cells(ColCreatedDateDO.Index).Value = .Created_Date_DO
                    rd.Cells(ColPrimaryOutSales.Index).Value = .Primary
                    rd.Cells(ColPrimaryDo.Index).Value = .primary_do
                    rd.Cells(ColPrimaryCurrency.Index).Value = .primary_curr
                End With
            Next

            Me.DgvHeader.ClearSelection()
            If Me.DgvHeader.Rows.Count > 0 Then
                Me.DgvHeader.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.DgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                Me.DgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        End Try
    End Sub

    Private Sub GetShipToAddress(ByVal tbl As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerDataTable)
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerRow
        Dim DtaST As New DseTransaksiSalesInvoice_QuickByStockOutTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TblST As New DseTransaksiSalesInvoice_QuickByStockOut.tblmaster_customer_item_addressDataTable
        Dim rowST As DseTransaksiSalesInvoice_QuickByStockOut.tblmaster_customer_item_addressRow

        Try
            DtaST.Fill(TblST)

            For Each row In tbl.Rows
                With row
                    If .primary_customer_ship_to = -1 Then
                        .primary_customer_ship_to = .primary_cust
                        .cust_ship_to_code = "" '.cust_code
                        .cust_ship_to_name = "" '.cust_name
                        .identifier_address_customer_ship_to = .identifier_address_ship_to
                    End If
                End With
            Next

            TblST.DefaultView.RowFilter = ""
            TblST.DefaultView.Sort = "[Primary Customer]"
            tbl.DefaultView.RowFilter = ""
            tbl.DefaultView.Sort = "[primary customer ship to]"
            For i As Integer = 0 To tbl.DefaultView.Count - 1
                row = tbl.DefaultView(i).Row
                With row
                    Dim ij As Integer = 0
                    While TblST.DefaultView.Count > 0
                        rowST = TblST.DefaultView(ij).Row
                        If .primary_customer_ship_to < rowST.Primary_Customer Then
                            Exit While
                        ElseIf .primary_customer_ship_to = rowST.Primary_Customer Then
                            If .identifier_address_customer_ship_to = rowST.Identifier Then
                                .Ship_to_address = rowST.Line_1
                                Exit While
                            Else
                                ij += 1
                            End If
                        Else
                            TblST.Rows.Remove(rowST)
                        End If
                    End While
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterByDgv(ByVal tbl As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerDataTable)
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerRow
        Dim ListPrimaryAcuan As New List(Of Double)

        Try
            If Me.RboOutSales.Checked = True Then
                'filter by stock out sales
                ListPrimaryAcuan.Clear()
                For Each rd As DataGridViewRow In Me.DgvOutSales.Rows
                    If String.IsNullOrEmpty(rd.Cells(ColfilterOutSales_Primary.Index).Value) = False Then
                        If rd.Cells(ColfilterOutSales_Primary.Index).Value <> -1 Then
                            If ListPrimaryAcuan.Contains(rd.Cells(ColfilterOutSales_Primary.Index).Value) = False Then
                                ListPrimaryAcuan.Add(rd.Cells(ColfilterOutSales_Primary.Index).Value)
                            End If
                        End If
                    End If
                Next

                If ListPrimaryAcuan.Count > 0 Then
                    For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                        row = tbl.Rows(i)
                        If ListPrimaryAcuan.Contains(row.Primary) = False Then
                            tbl.Rows.Remove(row)
                        End If
                    Next
                End If
                '---------------------------------------------------------------------------------------------------------------
            Else
                'filter by delivery order
                ListPrimaryAcuan.Clear()
                For Each rd As DataGridViewRow In Me.DgvDO.Rows
                    If String.IsNullOrEmpty(rd.Cells(ColDO_Primary.Index).Value) = False Then
                        If rd.Cells(ColDO_Primary.Index).Value <> -1 Then
                            If ListPrimaryAcuan.Contains(rd.Cells(ColDO_Primary.Index).Value) = False Then
                                ListPrimaryAcuan.Add(rd.Cells(ColDO_Primary.Index).Value)
                            End If
                        End If
                    End If
                Next

                If ListPrimaryAcuan.Count > 0 Then
                    For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                        row = tbl.Rows(i)
                        If ListPrimaryAcuan.Contains(row.primary_do) = False Then
                            tbl.Rows.Remove(row)
                        End If
                    Next
                End If
                '---------------------------------------------------------------------------------------------------------------
            End If

            'filter by customer
            ListPrimaryAcuan.Clear()
            For Each rd As DataGridViewRow In Me.dgvCust.Rows
                If String.IsNullOrEmpty(rd.Cells(ColCustPrimary.Index).Value) = False Then
                    If rd.Cells(ColCustPrimary.Index).Value <> -1 Then
                        If ListPrimaryAcuan.Contains(rd.Cells(ColCustPrimary.Index).Value) = False Then
                            ListPrimaryAcuan.Add(rd.Cells(ColCustPrimary.Index).Value)
                        End If
                    End If
                End If
            Next

            If ListPrimaryAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimaryAcuan.Contains(row.primary_cust) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------

            'filter by customer group
            ListPrimaryAcuan.Clear()
            For Each rd As DataGridViewRow In Me.dgvCustomerGroup.Rows
                If String.IsNullOrEmpty(rd.Cells(ColSG_Primary.Index).Value) = False Then
                    If rd.Cells(ColSG_Primary.Index).Value <> -1 Then
                        If ListPrimaryAcuan.Contains(rd.Cells(ColSG_Primary.Index).Value) = False Then
                            ListPrimaryAcuan.Add(rd.Cells(ColSG_Primary.Index).Value)
                        End If
                    End If
                End If
            Next

            If ListPrimaryAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimaryAcuan.Contains(row.Primary_Customer_Group) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------

            'filter by customer segmentation
            ListPrimaryAcuan.Clear()
            For Each rd As DataGridViewRow In Me.DgvMarketSegmentation.Rows
                If String.IsNullOrEmpty(rd.Cells(ColMS_Primary.Index).Value) = False Then
                    If rd.Cells(ColMS_Primary.Index).Value <> -1 Then
                        If ListPrimaryAcuan.Contains(rd.Cells(ColMS_Primary.Index).Value) = False Then
                            ListPrimaryAcuan.Add(rd.Cells(ColMS_Primary.Index).Value)
                        End If
                    End If
                End If
            Next

            If ListPrimaryAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimaryAcuan.Contains(row.Primary_Customer_Segmentasi) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------

            'filter by customer type
            ListPrimaryAcuan.Clear()
            For Each rd As DataGridViewRow In Me.DgvCustomerType.Rows
                If String.IsNullOrEmpty(rd.Cells(ColCT_Primary.Index).Value) = False Then
                    If rd.Cells(ColCT_Primary.Index).Value <> -1 Then
                        If ListPrimaryAcuan.Contains(rd.Cells(ColCT_Primary.Index).Value) = False Then
                            ListPrimaryAcuan.Add(rd.Cells(ColCT_Primary.Index).Value)
                        End If
                    End If
                End If
            Next

            If ListPrimaryAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimaryAcuan.Contains(row.Primary_Customer_Type) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------

            'filter by sales area
            ListPrimaryAcuan.Clear()
            For Each rd As DataGridViewRow In Me.DgvSalesArea.Rows
                If String.IsNullOrEmpty(rd.Cells(ColSA_Primary.Index).Value) = False Then
                    If rd.Cells(ColSA_Primary.Index).Value <> -1 Then
                        If ListPrimaryAcuan.Contains(rd.Cells(ColSA_Primary.Index).Value) = False Then
                            ListPrimaryAcuan.Add(rd.Cells(ColSA_Primary.Index).Value)
                        End If
                    End If
                End If
            Next

            If ListPrimaryAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If ListPrimaryAcuan.Contains(row.Primary_Sales_Area) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------

            'filter by created by
            Dim ListCodeAcuan As New List(Of String)
            Dim CreatedByAcuan As String = ""
            For Each rd As DataGridViewRow In Me.DgvFilter_CreatedBy.Rows
                If String.IsNullOrEmpty(rd.Cells(ColUser_Code.Index).Value) = False Then
                    If ListCodeAcuan.Contains(rd.Cells(ColUser_Code.Index).Value) = False Then
                        ListCodeAcuan.Add(rd.Cells(ColUser_Code.Index).Value)
                    End If
                End If
            Next

            If ListCodeAcuan.Count > 0 Then
                For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                    row = tbl.Rows(i)
                    If Me.RboOutSales.Checked = True Then
                        CreatedByAcuan = row.Created_By_OutSales
                    Else
                        CreatedByAcuan = row.Created_By_DO
                    End If

                    If ListCodeAcuan.Contains(CreatedByAcuan) = False Then
                        tbl.Rows.Remove(row)
                    End If
                Next
            End If
            '---------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCreate.Click
        Dim tblCreate, tblCreateToInv As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable
        Dim rowCreate, rowCreateAft, rowCreateToInv As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceRow
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerRow
        Dim ij As New FrmTransaksiSalesInvoice_QuickByStockOut_Option
        Dim FrmInv As New FrmTransaksiSalesInvoice
        Dim ListPrimaryDO As New List(Of Double)
        Dim ListPrimaryStockOutSales As New List(Of Double)
        Dim ListInvoiceGroup As New List(Of Double)
        Dim InvGroup As Integer = 0
        Dim TotalInvGroup As Integer = 0

        Try
            For Each rd As DataGridViewRow In Me.DgvHeader.Rows
                If rd.Cells(Me.ColCreate.Index).Value = True Then
                    If ListPrimaryDO.Contains(rd.Cells(Me.ColPrimaryDo.Index).Value) = False Then
                        ListPrimaryDO.Add(rd.Cells(Me.ColPrimaryDo.Index).Value)
                    End If
                    If ListPrimaryStockOutSales.Contains(rd.Cells(Me.ColPrimaryOutSales.Index).Value) = False Then
                        ListPrimaryStockOutSales.Add(rd.Cells(Me.ColPrimaryOutSales.Index).Value)
                    End If
                End If
            Next

            If ListPrimaryDO.Count = 0 Then
                MsgBox("Please select at least one stock out sales transaction.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each row In tbl
                With row
                    InvGroup = -1
                    If ListPrimaryDO.Contains(.primary_do) = True Then
                        tblCreate.DefaultView.RowFilter = "[primary cust] = " & .primary_cust & " AND [primary customer ship to] = " & .primary_customer_ship_to _
                        & " AND [use tax] = " & .use_tax & " AND [tax perc] = " & .tax_perc & " AND [primary curr] = " & .primary_curr '& " AND [address identifier] = '" &  .address_identifier & "'"
                        If tblCreate.DefaultView.Count > 0 Then
                            rowCreate = tblCreate.DefaultView(0).Row
                            InvGroup = rowCreate.invoice_group
                        Else
                            TotalInvGroup += 1
                        End If

                        rowCreate = tblCreate.Newtbl_create_invoiceRow
                        rowCreate.primary_cust = .primary_cust
                        rowCreate.primary_customer_ship_to = .primary_customer_ship_to
                        rowCreate.use_tax = .use_tax
                        rowCreate.tax_perc = .tax_perc
                        rowCreate.primary_curr = .primary_curr
                        rowCreate.primary_do = .primary_do
                        If InvGroup <> -1 Then
                            rowCreate.invoice_group = InvGroup
                        Else
                            rowCreate.invoice_group = tblCreate.Count + 1
                        End If
                        If ListInvoiceGroup.Contains(rowCreate.invoice_group) = False Then
                            ListInvoiceGroup.Add(rowCreate.invoice_group)
                        End If
                        tblCreate.Rows.Add(rowCreate)
                    End If
                End With
            Next

            ij.ListPrimaryStockOutSales = ListPrimaryStockOutSales
            ij.tblOutSales = tbl.Copy
            ij.tblCreate = tblCreate.Copy
            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim SimpanInvoice As Boolean = False

                Me.Panel1.Visible = True
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = ListInvoiceGroup.Count
                Me.Label1.Visible = True
                Me.Label1.Text = ""

                Me.SimpanByAntrianInvoice(ListInvoiceGroup, FrmInv, ij, tblCreate, tblCreateToInv)

                Me.Panel1.Visible = False
                Me.BtnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Overridable Function SimpanByAntrianInvoice(ByVal ListInvoiceGroup As List(Of Double), ByVal FrmInv As FrmTransaksiSalesInvoice, _
    ByVal ij As FrmTransaksiSalesInvoice_QuickByStockOut_Option, ByVal tblCreate As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable, _
    ByVal tblCreateToInv As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'SW.Start()

            ''ambil no antrian
            'Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, -1, _
            'Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_SimpanBaru, Me.UserCode)
            ''Me.GetNoAntrian()

            ''coba simpan data baru
            'Me.Clock_Antri.Reset()
            'Me.Clock_Antri.Start()

            'While Start = 0
            '    'hanya mencoba ke database setelah interval waktu tertentu saja
            '    If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
            '        If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
            '            MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
            '            Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
            '            Exit Function
            '        Else
            '            If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
            '                'clock antri matikan
            '                Me.Clock_Antri.Stop()

            '                'sudah saatnya melakukan proses simpan baru
            For Each Inv As Double In ListInvoiceGroup
                Me.ProgressBar1.Value += 1
                Me.Label1.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString '+ " - PO Number: " + ij.txtPONO.Text

                FrmInv = New FrmTransaksiSalesInvoice
                FrmInv.ProgressBar1.Visible = True
                FrmInv.lblProgress.Visible = True
                FrmInv.ProgressBar1.Maximum = Me.ProgressBar1.Maximum
                FrmInv.ProgressBar1.Value = Me.ProgressBar1.Value
                FrmInv.lblProgress.Text = Me.Label1.Text

                FrmInv.UserCode = Me.UserCode
                FrmInv.QuickInput = True
                FrmInv.LinkLabel1.Enabled = False
                FrmInv.DirectlySave = ij.ChkDirectlySave.Checked
                FrmInv.DirectlyPrint = ij.ChkDirectlyPrint.Checked
                FrmInv.PathPrintCustom = ij.TxtUseCustomPrinter.Text
                FrmInv.SetDtpDate = ij.DtpDate.Value.Date


                tblCreate.DefaultView.RowFilter = "[invoice group] = " & Inv
                tblCreateToInv.Rows.Clear()
                tblCreateToInv.Merge(tblCreate.DefaultView.ToTable.Copy)
                'CreateInvoiceByQuickInputStockOut(tblCreateToInv, ij.ChkDirectlySave.Checked, ij.ChkDirectlyPrint.Checked, ij.TxtUseCustomPrinter.Text, FrmInv)
                FrmInv.TblQuickInput = tblCreateToInv.Copy
                FrmInv.ShowDialog()

                If FrmInv.StopQuickInput = True Then
                    Exit For
                End If
            Next

            '                'hapus data antri untuk transaksi ini
            '                If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
            '                    MsgBox("Logic error(1), contact your provider when you see this message", _
            '                    MsgBoxStyle.Information, "Information")
            '                End If
            '                'nilai balikan
            '                Return retval
            '            Else
            '                'kontrol interval pengecekan
            '                TryCounter += 1
            '                CheckInterval = ((TryCounter \ 5) + 1) * 500
            '                Me.Clock_Antri.Reset()
            '                Me.Clock_Antri.Start()
            '            End If
            '        End If
            '    End If
            'End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function
    
    Private Sub DgvHeader_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvHeader.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCreate.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHeader_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvHeader.DataError

    End Sub

    Private Sub DgvHeader_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvHeader.SelectionChanged
        Try
            If Me.DgvHeader.SelectedCells.Count > 0 Then
                Dim PrimaryStockOutSales As Double = -1
                Dim PrimaryCurr As Double = -1
                If Me.DgvHeader(Me.ColPrimaryOutSales.Index, Me.DgvHeader.SelectedCells(0).RowIndex).Value <> -1 Then
                    PrimaryStockOutSales = Me.DgvHeader(Me.ColPrimaryOutSales.Index, Me.DgvHeader.SelectedCells(0).RowIndex).Value
                End If
                If Me.DgvHeader(Me.ColPrimaryCurrency.Index, Me.DgvHeader.SelectedCells(0).RowIndex).Value <> -1 Then
                    PrimaryCurr = Me.DgvHeader(Me.ColPrimaryCurrency.Index, Me.DgvHeader.SelectedCells(0).RowIndex).Value
                End If

                If PrimaryStockOutSales = -1 Or PrimaryCurr = -1 Then
                    Me.DgvItem.Rows.Clear()
                    Exit Sub
                End If

                Me.IsiDgvItem(PrimaryStockOutSales, PrimaryCurr)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgvItem(ByVal vPrimaryStockOutSales As Double, ByVal vPrimaryCurr As Double)
        Dim dta As New DseTransaksiSalesInvoice_QuickByStockOutTableAdapters.tbl_itemTableAdapter
        Dim tbl As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_itemDataTable
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_itemRow
        Dim rd As DataGridViewRow
        Dim DecPrice As Integer = 0
        Dim ShowBatch As Boolean = False
        Dim ShowExpDate As Boolean = False
        Dim ShowRatio As Boolean = False

        Try
            DecPrice = Me.ObjCurr.GetDefaultPrecission(vPrimaryCurr)

            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, vPrimaryStockOutSales)
            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To tbl.Rows.Count - 1
                row = tbl.Rows(i)
                With row
                    rd = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                    rd.Cells(Me.ColItem_No.Index).Value = rd.Index + 1
                    rd.Cells(Me.ColItem_ProductCode.Index).Value = .prod_code
                    rd.Cells(Me.ColItem_ProductDescription.Index).Value = .prod_desc
                    rd.Cells(Me.ColItem_SO.Index).Value = .so_no
                    rd.Cells(Me.ColItem_Batch.Index).Value = .batch
                    rd.Cells(Me.ColItem_Expiry.Index).Value = .expiry
                    If .expiry = #1/1/1950# Then
                        rd.Cells(Me.ColItem_Expiry.Index).Style.ForeColor = Color.Transparent
                        rd.Cells(Me.ColItem_Expiry.Index).Style.SelectionForeColor = Color.Transparent
                    End If
                    rd.Cells(Me.ColItem_TotalQty.Index).Value = Me.ObjCD.DigitGrouping(.qty_multi, .decmulti)
                    rd.Cells(Me.ColItem_Unit.Index).Value = .unit_code_multi
                    If .ratio_multi <= 1 And .ratio_def <= 1 Then
                        rd.Cells(Me.ColItem_Ratio.Index).Value = ""
                    Else
                        rd.Cells(Me.ColItem_Ratio.Index).Value = Me.ObjCD.DigitGrouping(.ratio_multi, .decmulti) & " " & .unit_code_multi & " = " & Me.ObjCD.DigitGrouping(.ratio_def, .decdef) & " " & .unit_code_def
                    End If
                    rd.Cells(Me.ColItem_QtyDefault.Index).Value = Me.ObjCD.DigitGrouping(.qty_def, .decdef)
                    rd.Cells(Me.ColItem_UnitDefault.Index).Value = .unit_code_def
                    rd.Cells(Me.ColItem_UnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price, DecPrice)
                    rd.Cells(Me.ColItem_Disc.Index).Value = Me.ObjCD.DigitGrouping(.Discount, 2)
                    rd.Cells(Me.ColItem_DiscAmount.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, DecPrice)
                    Dim QtyTemp As Decimal = .qty_multi
                    Dim PriceStr As String = .unit_code_multi
                    If .primary_unit <> .primary_unit_multi Then
                        If .price_by_multi_unit = 0 Then
                            QtyTemp = .qty_def
                            PriceStr = .unit_code_def
                        End If
                    End If
                    rd.Cells(Me.ColItem_PricePer.Index).Value = PriceStr
                    rd.Cells(Me.ColItem_NettPrice.Index).Value = Me.ObjCD.DigitGrouping(.Nett_Price, DecPrice)
                    rd.Cells(Me.ColItem_Amount.Index).Value = Me.ObjCD.DigitGrouping(QtyTemp * .Nett_Price, DecPrice)
                    rd.Cells(Me.ColItem_Remarks.Index).Value = .remarks_item_stockoutsales

                    If String.IsNullOrEmpty(rd.Cells(Me.ColItem_Batch.Index).Value) = False Then
                        ShowBatch = True
                    End If
                    If String.IsNullOrEmpty(rd.Cells(Me.ColItem_Expiry.Index).Value) = False Then
                        If rd.Cells(Me.ColItem_Expiry.Index).Value <> #1/1/1950# Then
                            ShowExpDate = True
                        End If
                    End If
                    If String.IsNullOrEmpty(rd.Cells(Me.ColItem_Ratio.Index).Value) = False Then
                        ShowRatio = True
                    End If
                End With
            Next

            Me.ColItem_Batch.Visible = ShowBatch
            Me.ColItem_Expiry.Visible = ShowExpDate
            Me.ColItem_Ratio.Visible = ShowRatio
            Me.ColItem_QtyDefault.Visible = ShowRatio
            Me.ColItem_UnitDefault.Visible = ShowRatio
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None Then
                Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        End Try
    End Sub

    Private Sub DgvHeader_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvHeader.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvHeader)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvItem.DataError

    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "filter"
    Private Sub DgvFilter_CreatedBy_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvFilter_CreatedBy.CellBeginEdit
        If e.ColumnIndex = Me.ColUser_Code.Index Then
            Me.buffer = Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
        Else
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvFilter_CreatedBy.Rows.Count - 2
                RowDGV = Me.DgvFilter_CreatedBy.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColUser_Code.Index).Value = vKodeUser Then
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

    Private Sub DgvFilter_CreatedBy_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvFilter_CreatedBy.CellEndEdit
        Dim KodeUser As String = ""
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim RowIdx As Integer = e.RowIndex
        Try
            If IsNothing(Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColUser_Code.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.DgvFilter_CreatedBy.RowCount - 1 Then
                        Me.rowIndex = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            Else
                ''data account number
                If e.ColumnIndex = Me.ColUser_Code.Index Then
                    Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUser.ApakahSudahAda(Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        KodeUser = Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        index = Me.GetIndexItem(KodeUser, e.RowIndex)
                        If index < 0 Then
                            Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value = Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value
                            Me.DgvFilter_CreatedBy.Item(Me.ColUser_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value)
                        Else
                            If e.RowIndex <> index Then
                                MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.rowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterUser
                        ij.backColorAll = Me.BackColorAll
                        ij.TxtCari.Text = Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.Kodeuser <> "" Then
                            index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value = ij.Kodeuser
                                Me.DgvFilter_CreatedBy.Item(Me.ColUser_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.rowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                Else
                                    Me.DgvFilter_CreatedBy.Item(Me.ColUser_Code.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvFilter_CreatedBy.Item(Me.ColUser_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                End If
                            End If
                            ''2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            'Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = ij.Kodeuser.ToUpper
                            'Me.DgvFilter_CreatedBy.Item(Me.ColUser_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value)
                        ElseIf ij.Kodeuser = "" Then
                            If e.ColumnIndex = Me.ColUser_Code.Index Then
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    If Me.DgvFilter_CreatedBy.Rows.Count - 1 <> e.RowIndex Then
                                        Me.rowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                Else
                                    Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            Else
                                Me.DgvFilter_CreatedBy.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
            Me.ObjCD.SetNumber(Me.DgvFilter_CreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvFilter_CreatedBy_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvFilter_CreatedBy.DataError

    End Sub

    Private Sub DgvFilter_CreatedBy_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvFilter_CreatedBy.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.DgvFilter_CreatedBy.Rows(Me.DgvFilter_CreatedBy.RowCount - 1)
            With rowDGV
                .Cells(Me.ColUser_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColUser_Code.Index).Value = ""
                .Cells(Me.ColUser_Name.Index).Value = ""
                .Cells(Me.ColUser_Primary.Index).Value = "-1"
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvFilter_CreatedBy_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvFilter_CreatedBy.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvFilter_CreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCreatedBy_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvFilter_CreatedBy.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvFilter_CreatedBy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCust_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCust.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCustCode.Index Then
                e.Cancel = True
                Exit Sub
            Else
                'col account
                If e.ColumnIndex = Me.ColCustCode.Index Then
                    Me.buffer = Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIdxCust(ByVal vRefIdx As Integer, ByVal vCode As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvCust.Rows
                If rowDGV.Index <> Me.dgvCust.Rows.Count - 1 Then
                    If rowDGV.Cells(Me.ColCustCode.Index).Value = vCode And rowDGV.Index <> vRefIdx Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Sub dgvCust_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCust.CellEndEdit
        Try
            If IsNothing(Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                'kalo kode cust dihapus, maka hapus satu baris semuanya
                If e.RowIndex <> Me.dgvCust.RowCount - 1 Then
                    Me.EraseIndex_C = e.RowIndex
                    Me.bgwCustomer.RunWorkerAsync()
                End If
                Exit Sub
            Else
                ''data cust code
                If e.ColumnIndex = Me.ColCustCode.Index Then
                    Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjCust.ApakahSudahAda(Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode cust yang diketik ada --> load data deskripsi + primary cust
                        Me.dgvCust.Item(Me.ColCustPrimary.Index, e.RowIndex).Value = Me.ObjCust.GetPrimary(Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvCust.Item(Me.ColCustDesc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(Me.dgvCust.Item(Me.ColCustPrimary.Index, e.RowIndex).Value)
                    Else
                        '2. kode cust yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterCustomer
                        ij.txtCari.Text = Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryCustomer <> -1 Then
                            '2.a. kode cust yang dicari ada --> load data deskripsi + primary cust + kode cust
                            Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCust.GetCustomerCode(ij.PrimaryCustomer)
                            Me.dgvCust.Item(Me.ColCustDesc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(ij.PrimaryCustomer)
                            Me.dgvCust.Item(Me.ColCustPrimary.Index, e.RowIndex).Value = ij.PrimaryCustomer
                        ElseIf ij.PrimaryCustomer = -1 Then
                            '2.b. kode cust yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            If Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.EraseIndex_C = e.RowIndex
                                Me.bgwCustomer.RunWorkerAsync()
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If

            'pastikan tidak double dengan yang lainnya
            rowIndex = Me.GetIdxCust(e.RowIndex, Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value)
            If rowIndex <> -1 Then
                MsgBox("Code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.EraseIndex_C = e.RowIndex
                    Me.bgwCustomer.RunWorkerAsync()
                Else
                    Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvCust.Item(Me.ColCustPrimary.Index, e.RowIndex).Value = Me.ObjCust.GetPrimary(Me.dgvCust.Item(e.ColumnIndex, e.RowIndex).Value)
                    Me.dgvCust.Item(Me.ColCustDesc.Index, e.RowIndex).Value = Me.ObjCust.GetCustName(Me.dgvCust.Item(Me.ColCustPrimary.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCust_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCust.DataError

    End Sub

    Private Sub dgvCustProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCust.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvCust.Rows(Me.dgvCust.RowCount - 1)
            rowDGV.Cells(Me.ColCustPrimary.Index).Value = -1
            rowDGV.Cells(Me.ColCustCode.Index).Value = ""
            rowDGV.Cells(Me.ColCustDesc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.rowIndex <> -1 And Me.rowIndex <> Me.dgvCust.RowCount - 1 Then
                Me.DgvFilter_CreatedBy.Rows.RemoveAt(Me.rowIndex)
                Me.rowIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustomer.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustomer.RunWorkerCompleted
        Try
            If Me.EraseIndex_C <> -1 And Me.EraseIndex_C <> Me.dgvCust.RowCount - 1 Then
                Me.dgvCust.Rows.RemoveAt(Me.EraseIndex_C)
                Me.rowIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomerGroup.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                Me.buffer = Me.dgvCustomerGroup(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvCustomerGroup.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerGroup
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                        .Cells(Me.ColSG_SalesGroup.Index).Value = .Cells(Me.ColSG_SalesGroup.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSG_SalesGroup.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSG_SalesGroup.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CG(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSG_Primary.Index).Value = primary
                                .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(primary)
                                .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerGroup
                            ij.txtCari.Text = Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerGroup <> -1 Then
                                index = Me.GetIndexItem_CG(ij.PrimaryCustomerGroup, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                    .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                    .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SG = e.RowIndex
                                            Me.bgwCustGroup.RunWorkerAsync()
                                        Else
                                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                        .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                        .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerGroup = -1 Then
                                If Me.dgvCustomerGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                        Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value = Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.dgvCustomerGroup.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SG = e.RowIndex
                                Me.bgwCustGroup.RunWorkerAsync()
                            Else
                                Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                            Exit Sub
                        Else
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSG_SalesGroup.Index Then
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_CG(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvCustomerGroup.Rows.Count - 1
                RowDGV = Me.dgvCustomerGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSG_Primary.Index).Value = vPrimary Then
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

    Private Sub bgwCustGroup_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustGroup.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustGroup_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustGroup.RunWorkerCompleted
        Try
            Me.dgvCustomerGroup.Rows.RemoveAt(EraseIndex_SG)
            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvCustomerGroup.DataError

    End Sub

    Private Sub dgvCustomerGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCustomerGroup.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvCustomerGroup.Rows(Me.dgvCustomerGroup.Rows.Count - 1)
            dr.Cells(Me.ColSG_No.Index).Value = Me.dgvCustomerGroup.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvCustomerGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCustomerGroup.Sorted
        Try
            ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomerType.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCT_CustType.Index Then
                Me.buffer = Me.DgvCustomerType(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomerType.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomerType.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjCT As New Core_Module.clsMasterCustomerType
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCT_CustType.Index Then
                        .Cells(Me.ColCT_CustType.Index).Value = .Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper
                        If ObjCT.ApakahSudahAda(.Cells(Me.ColCT_CustType.Index).Value) = True Then
                            primary = ObjCT.GetPrimary(.Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CT(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCT_Primary.Index).Value = primary
                                .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(primary)
                                .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCustomerType
                            ij.txtCari.Text = Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerType <> -1 Then
                                index = Me.GetIndexItem_CT(ij.PrimaryCustomerType, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                    .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                    .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_CT = e.RowIndex
                                            Me.bgwCustType.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                        .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                        .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerType = -1 Then
                                If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                        Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCT_CustType.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_CT = e.RowIndex
                                Me.bgwCustType.RunWorkerAsync()
                            Else
                                Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomerType)
                            Exit Sub
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCT_CustType.Index Then
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_CT(ByVal vPrimary As Double, _
       ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomerType.Rows.Count - 1
                RowDGV = Me.DgvCustomerType.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCT_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvCustomerType_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvCustomerType.DataError

    End Sub

    Private Sub DgvCustomerType_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvCustomerType.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvCustomerType.Rows(Me.DgvCustomerType.Rows.Count - 1)
            dr.Cells(Me.ColCT_No.Index).Value = Me.DgvCustomerType.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomerType.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomerType.Sorted
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustType_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustType.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustType_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustType.RunWorkerCompleted
        Try
            Me.DgvCustomerType.Rows.RemoveAt(EraseIndex_CT)
            Me.ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMarketSegmentation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                Me.buffer = Me.DgvMarketSegmentation(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMarketSegmentation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvMarketSegmentation.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                        .Cells(Me.ColMS_MarketSeg.Index).Value = .Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper
                        If ObjMS.ApakahSudahAda(.Cells(Me.ColMS_MarketSeg.Index).Value) = True Then
                            primary = ObjMS.GetPrimary(.Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_MS(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColMS_Primary.Index).Value = primary
                                .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(primary)
                                .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerSegmentation
                            ij.txtCari.Text = Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerMarketSegmentation <> -1 Then
                                index = Me.GetIndexItem_MS(ij.PrimaryCustomerMarketSegmentation, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                    .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                    .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_MS = e.RowIndex
                                            Me.bgwCustSegmen.RunWorkerAsync()
                                        Else
                                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                        .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                        .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerMarketSegmentation = -1 Then
                                If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()

                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                        Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_MS = e.RowIndex
                                Me.bgwCustSegmen.RunWorkerAsync()
                            Else
                                Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                            Exit Sub
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColMS_MarketSeg.Index Then
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_MS(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvMarketSegmentation.Rows.Count - 1
                RowDGV = Me.DgvMarketSegmentation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColMS_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvMarketSegmentation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvMarketSegmentation.DataError

    End Sub

    Private Sub DgvMarketSegmentation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvMarketSegmentation.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvMarketSegmentation.Rows(Me.DgvMarketSegmentation.Rows.Count - 1)
            dr.Cells(Me.ColMS_No.Index).Value = Me.DgvMarketSegmentation.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvMarketSegmentation.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustSegmen_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustSegmen.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustSegmen_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustSegmen.RunWorkerCompleted
        Try
            Me.DgvMarketSegmentation.Rows.RemoveAt(EraseIndex_MS)
            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalesArea.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                Me.buffer = Me.DgvSalesArea(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalesArea.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalesArea.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                        .Cells(Me.ColSA_SalesArea.Index).Value = .Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSA_SalesArea.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_SA(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSA_Primary.Index).Value = primary
                                .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(primary)
                                .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSalesArea
                            ij.txtCari.Text = Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerSalesArea <> -1 Then
                                index = Me.GetIndexItem_SA(ij.PrimaryCustomerSalesArea, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                    .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                    .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SA = e.RowIndex
                                            Me.bgwSalesArea.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                        .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                        .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerSalesArea = -1 Then
                                If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                        Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SA = e.RowIndex
                                Me.bgwSalesArea.RunWorkerAsync()
                            Else
                                Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesArea)
                            Exit Sub
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSA_SalesArea.Index Then
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_SA(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalesArea.Rows.Count - 1
                RowDGV = Me.DgvSalesArea.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSA_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvSalesArea_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvSalesArea.DataError
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSalesArea.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvSalesArea.Rows(Me.DgvSalesArea.Rows.Count - 1)
            dr.Cells(Me.ColSA_No.Index).Value = Me.DgvSalesArea.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalesArea.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSalesArea.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvSalesArea, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwSalesArea_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwSalesArea.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwSalesArea_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwSalesArea.RunWorkerCompleted
        Try
            Me.DgvSalesArea.Rows.RemoveAt(EraseIndex_SA)
            Me.ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub DgvDO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvDO.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColDO_DONo.Index Then
                Me.buffer = Me.DgvDO(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_DO(ByVal vPrimary As Double, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvDO.Rows.Count - 1
                RowDGV = Me.DgvDO.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColDO_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvDO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDO.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvDO.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New ClsTransaksiDO

        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColDO_DONo.Index Then
                        .Cells(Me.ColDO_DONo.Index).Value = .Cells(Me.ColDO_DONo.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColDO_DONo.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColDO_DONo.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_DO(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColDO_Primary.Index).Value = primary
                                .Cells(Me.ColDO_DONo.Index).Value = ObjSA.GetDONO(primary)
                                .Cells(Me.ColDO_Date.Index).Value = Format(ObjSA.GetTanggalDO(primary), "dd MMM yyyy")
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication transaction no, please input another transaction no.", MsgBoxStyle.Information, "Information")
                                    Me.DgvDO.Item(Me.ColDO_DONo.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_DO = e.RowIndex
                                        Me.bgwDO.RunWorkerAsync()
                                    Else
                                        Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvDO)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New FrmCariTransaksiDO
                            ij.txtCari.Text = Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryTransaksi <> -1 Then
                                index = Me.GetIndexItem_DO(ij.PrimaryTransaksi, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColDO_Primary.Index).Value = ij.PrimaryTransaksi
                                    .Cells(Me.ColDO_DONo.Index).Value = ObjSA.GetDONO(ij.PrimaryTransaksi)
                                    .Cells(Me.ColDO_Date.Index).Value = Format(ObjSA.GetTanggalDO(ij.PrimaryTransaksi), "dd MMM yyyy")
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication transaction no, please input another transaction no.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_DO = e.RowIndex
                                            Me.bgwDO.RunWorkerAsync()
                                        Else
                                            Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvDO)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColDO_Primary.Index).Value = ij.PrimaryTransaksi
                                        .Cells(Me.ColDO_DONo.Index).Value = ObjSA.GetDONO(ij.PrimaryTransaksi)
                                        .Cells(Me.ColDO_Date.Index).Value = Format(ObjSA.GetTanggalDO(ij.PrimaryTransaksi), "dd MMM yyyy")
                                    End If
                                End If
                            ElseIf ij.PrimaryTransaksi = -1 Then
                                If Me.DgvDO.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_DO = e.RowIndex
                                        Me.bgwDO.RunWorkerAsync()
                                    Else
                                        Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvDO)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvDO)
                        Me.DgvDO.Item(Me.ColDO_DONo.Index, e.RowIndex).Value = Me.DgvDO.Item(Me.ColDO_DONo.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvDO.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColDO_DONo.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_DO = e.RowIndex
                                Me.bgwDO.RunWorkerAsync()
                            Else
                                Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvDO)
                            Exit Sub
                        Else
                            Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColDO_DONo.Index Then
                            Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvDO.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvDO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvDO.DataError

    End Sub

    Private Sub DgvDO_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvDO.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvDO.Rows(Me.DgvDO.Rows.Count - 1)
            dr.Cells(Me.ColDO_No.Index).Value = Me.DgvDO.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDO_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvDO.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDO_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvDO.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvDO, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwDO_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwDO.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwDO_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwDO.RunWorkerCompleted
        Try
            Me.DgvDO.Rows.RemoveAt(EraseIndex_DO)
            Me.ObjCD.SetNumber(Me.DgvDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOutSales_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvOutSales.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColfilterOutSales_DocNo.Index Then
                Me.buffer = Me.DgvOutSales(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_OutSales(ByVal vPrimary As Double, ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvOutSales.Rows.Count - 1
                RowDGV = Me.DgvOutSales.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColfilterOutSales_Primary.Index).Value = vPrimary Then
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

    Private Sub DgvOutSales_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvOutSales.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvOutSales.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Inventory_Module.clsTransaksiStockInOut

        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColfilterOutSales_DocNo.Index Then
                        .Cells(Me.ColfilterOutSales_DocNo.Index).Value = .Cells(Me.ColfilterOutSales_DocNo.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColfilterOutSales_DocNo.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColfilterOutSales_DocNo.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_OutSales(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColfilterOutSales_Primary.Index).Value = primary
                                .Cells(Me.ColfilterOutSales_DocNo.Index).Value = ObjSA.GetNoTransaksi(primary)
                                .Cells(Me.ColfilterOutSales_Date.Index).Value = Format(ObjSA.GetTanggalTransaksi(primary), "dd MMM yyyy")
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication transaction no, please input another transaction no.", MsgBoxStyle.Information, "Information")
                                    Me.DgvOutSales.Item(Me.ColfilterOutSales_DocNo.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SOS = e.RowIndex
                                        Me.bgwOutSales.RunWorkerAsync()
                                    Else
                                        Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvOutSales)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Inventory_Module.FrmCariTransaksiStockInOut_Sales
                            ij.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                            ij.txtCari.Text = Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryTransaksiStockInOut <> -1 Then
                                index = Me.GetIndexItem_OutSales(ij.PrimaryTransaksiStockInOut, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColfilterOutSales_Primary.Index).Value = ij.PrimaryTransaksiStockInOut
                                    .Cells(Me.ColfilterOutSales_DocNo.Index).Value = ObjSA.GetNoTransaksi(ij.PrimaryTransaksiStockInOut)
                                    .Cells(Me.ColfilterOutSales_Date.Index).Value = Format(ObjSA.GetTanggalTransaksi(ij.PrimaryTransaksiStockInOut), "dd MMM yyyy")
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication transaction no, please input another transaction no.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SOS = e.RowIndex
                                            Me.bgwOutSales.RunWorkerAsync()
                                        Else
                                            Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvOutSales)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColfilterOutSales_Primary.Index).Value = ij.PrimaryTransaksiStockInOut
                                        .Cells(Me.ColfilterOutSales_DocNo.Index).Value = ObjSA.GetNoTransaksi(ij.PrimaryTransaksiStockInOut)
                                        .Cells(Me.ColfilterOutSales_Date.Index).Value = Format(ObjSA.GetTanggalTransaksi(ij.PrimaryTransaksiStockInOut), "dd MMM yyyy")
                                    End If
                                End If
                            ElseIf ij.PrimaryTransaksiStockInOut = -1 Then
                                If Me.DgvOutSales.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SOS = e.RowIndex
                                        Me.bgwOutSales.RunWorkerAsync()
                                    Else
                                        Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvOutSales)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvOutSales)
                        Me.DgvOutSales.Item(Me.ColfilterOutSales_DocNo.Index, e.RowIndex).Value = Me.DgvOutSales.Item(Me.ColfilterOutSales_DocNo.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvOutSales.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColfilterOutSales_DocNo.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SOS = e.RowIndex
                                Me.bgwOutSales.RunWorkerAsync()
                            Else
                                Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvOutSales)
                            Exit Sub
                        Else
                            Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColfilterOutSales_DocNo.Index Then
                            Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvOutSales.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvOutSales_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvOutSales.DataError

    End Sub

    Private Sub DgvOutSales_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvOutSales.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvOutSales.Rows(Me.DgvOutSales.Rows.Count - 1)
            dr.Cells(Me.ColfilterOutSales_No.Index).Value = Me.DgvOutSales.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOutSales_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvOutSales.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvOutSales)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOutSales_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvOutSales.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvOutSales, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwOutSales_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwOutSales.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwOutSales_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwOutSales.RunWorkerCompleted
        Try
            Me.DgvOutSales.Rows.RemoveAt(EraseIndex_SOS)
            Me.ObjCD.SetNumber(Me.DgvOutSales)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class