Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanPenjualan
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim DS1 As New DseLaporanSalesAnalysis
    Dim DS2 As New DseLaporanPenjualan
    Dim ObjCD As New ClsCekData

    Dim TBL As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0DataTable = DS2.tbltrans_salesinvoice_item_view0
    Dim DTA As New DseLaporanPenjualanTableAdapters.tbltrans_salesinvoice_item_view0TableAdapter

    Dim TBLDetail As DseLaporanPenjualan.DetailDataTable = DS2.Detail
    Dim TBLRekap As DseLaporanPenjualan.RekapDataTable = DS2.Rekap
    Dim TBLRekapComparison As DseLaporanPenjualan.Rekap_ComparisonDataTable = DS2.Rekap_Comparison

    'komponen yang berkaitan dengan clb customer
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim TblCust As New Core_Module.dseCustomer.tblmaster_customer1DataTable
    Dim rowCust As Core_Module.dseCustomer.tblmaster_customer1Row

    'kompinen yang berkaitan clb product
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim TblProd As New Core_Module.dseMasterProduct.tblmaster_product2DataTable
    Dim rowProd As Core_Module.dseMasterProduct.tblmaster_product2Row

    Dim TBLReturn As New DseLaporanSalesAnalysis.TblReturnInvoiceDataTable
    Dim DTAReturn As New DseLaporanSalesAnalysisTableAdapters.TblReturnInvoiceTableAdapter
    Dim rowReturn As DseLaporanSalesAnalysis.TblReturnInvoiceRow

    Dim ListBrand As New List(Of String)
    Dim ListCust As New List(Of String)

    Dim ReportDetail As New CrpLaporanPenjualanDetail
    Dim ReportRekap As New CrpLaporanPenjualanRekap
    Dim ReportRekap_Comparison As New CrpLaporanPenjualanRekap_2PeriodCust

    Dim Buffer As String = ""
    Dim index As Integer = -1
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim EraseIndex_MS As Integer = -1
    Dim EraseIndex_C As Integer = -1
    Dim EraseIndex_SA As Integer = -1
    Dim EraseIndex_SG As Integer = -1
    Dim EraseIndex_CT As Integer = -1
    Dim ObjCT As New Core_Module.clsMasterCustomerType
    Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup

    Dim Currdecimal As Integer = 0
    Dim objCurr As New Core_Module.clsMasterCurrency

    Private Enum EnumFormat As Integer
        Rekap
        Rekap_2_Period_Comparison
        Detail
    End Enum

    Private Enum EnumGroupBy As Integer
        CodeNo
        CustCode
        Tanggal
        Invoice
    End Enum

    Private Sub IsikanDGV()
        Dim rowdgv As DataGridViewRow
        Dim Name1, Name2, Name3 As String

        Try
            Name1 = ""
            Name2 = ""
            Name3 = ""

            Me.dgvSortDetail.Rows.Clear()
            Me.dgvSortRekap.Rows.Clear()
            Me.dgvSortRekap2PeriodComp.Rows.Clear()

            If Me.cboFormat.SelectedIndex = 2 Then
                If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
                         Or Me.cboGroupBy.SelectedIndex = 4 Then 'Code No,Category,Sub Brand, Brand, Product Group
                    'group by Code No,Category,Sub Brand, Brand, Product Group
                    Name1 = "Date"
                    Name2 = "Invoice Number"
                    Name3 = "Customer Code"
                ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                    'group by customer code
                    Name1 = "Date"
                    Name2 = "Invoice Number"
                    Name3 = "Code No"
                ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                    'group by tanggal
                    Name1 = "Invoice Number"
                    Name2 = "Customer Code"
                    Name3 = "Code No"
                End If
            End If
         
            Me.dgvSortDetail.Rows.Add()
            rowdgv = Me.dgvSortDetail.Rows(Me.dgvSortDetail.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = Name1
                .Cells(Me.ColSortNumber.Index).Value = "1"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            Me.dgvSortDetail.Rows.Add()
            rowdgv = Me.dgvSortDetail.Rows(Me.dgvSortDetail.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = Name2
                .Cells(Me.ColSortNumber.Index).Value = "2"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            Me.dgvSortDetail.Rows.Add()
            rowdgv = Me.dgvSortDetail.Rows(Me.dgvSortDetail.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = Name3
                .Cells(Me.ColSortNumber.Index).Value = "3"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            Me.dgvSortDetail.Rows.Add()
            rowdgv = Me.dgvSortDetail.Rows(Me.dgvSortDetail.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = "Qty"
                .Cells(Me.ColSortNumber.Index).Value = "4"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            Me.dgvSortDetail.Rows.Add()
            rowdgv = Me.dgvSortDetail.Rows(Me.dgvSortDetail.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = "Amount"
                .Cells(Me.ColSortNumber.Index).Value = "5"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With

            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                'group by code no
                Name1 = "Customer Code"
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                'group by customer code
                Name1 = "Code No"
            End If
            Me.dgvSortRekap.Rows.Add()
            rowdgv = Me.dgvSortRekap.Rows(Me.dgvSortRekap.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = Name1
                .Cells(Me.ColSortNumber.Index).Value = "1"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            If Me.cboFormat.SelectedIndex <> 0 Then
            Else
                If Me.cboGroupBy.SelectedIndex <> 5 Then
                    Me.dgvSortRekap.Rows.Add()
                    rowdgv = Me.dgvSortRekap.Rows(Me.dgvSortRekap.Rows.Count - 1)
                    With rowdgv
                        .Cells(Me.ColName.Index).Value = "Qty"
                        .Cells(Me.ColSortNumber.Index).Value = "2"
                        .Cells(Me.ColSrtType.Index).Value = "Ascending"
                    End With
                    Me.dgvSortRekap.Rows.Add()
                    rowdgv = Me.dgvSortRekap.Rows(Me.dgvSortRekap.Rows.Count - 1)
                    With rowdgv
                        .Cells(Me.ColName.Index).Value = "Amount"
                        .Cells(Me.ColSortNumber.Index).Value = "3"
                        .Cells(Me.ColSrtType.Index).Value = "Ascending"
                    End With
                Else
                    'group by cust, tidak ada sort by qty, hanya amount
                    Me.dgvSortRekap.Rows.Add()
                    rowdgv = Me.dgvSortRekap.Rows(Me.dgvSortRekap.Rows.Count - 1)
                    With rowdgv
                        .Cells(Me.ColName.Index).Value = "Amount"
                        .Cells(Me.ColSortNumber.Index).Value = "2"
                        .Cells(Me.ColSrtType.Index).Value = "Ascending"
                    End With
                End If
            End If

            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
             Or Me.cboGroupBy.SelectedIndex = 4 Then
                'group by code no
                Name1 = "Customer Code"
            ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                'group by customer code
                Name1 = "Code No"
            End If
            Me.dgvSortRekap2PeriodComp.Rows.Add()
            rowdgv = Me.dgvSortRekap2PeriodComp.Rows(Me.dgvSortRekap2PeriodComp.Rows.Count - 1)
            With rowdgv
                .Cells(Me.ColName.Index).Value = Name1
                .Cells(Me.ColSortNumber.Index).Value = "1"
                .Cells(Me.ColSrtType.Index).Value = "Ascending"
            End With
            If Me.cboFormat.SelectedIndex <> 0 Then
                'jika formatnya 2 period comparison, tambahkan % diference
                Me.dgvSortRekap2PeriodComp.Rows.Add()
                rowdgv = Me.dgvSortRekap2PeriodComp.Rows(Me.dgvSortRekap2PeriodComp.Rows.Count - 1)
                With rowdgv
                    .Cells(Me.ColName.Index).Value = "% Difference"
                    .Cells(Me.ColSortNumber.Index).Value = "2"
                    .Cells(Me.ColSrtType.Index).Value = "Ascending"
                End With

                'group by cust, tidak ada sort by qty, hanya amount
                Me.dgvSortRekap2PeriodComp.Rows.Add()
                rowdgv = Me.dgvSortRekap2PeriodComp.Rows(Me.dgvSortRekap2PeriodComp.Rows.Count - 1)
                With rowdgv
                    .Cells(Me.ColName.Index).Value = "1-st Period Amount"
                    .Cells(Me.ColSortNumber.Index).Value = "3"
                    .Cells(Me.ColSrtType.Index).Value = "Ascending"
                End With
                'group by cust, tidak ada sort by qty, hanya amount
                Me.dgvSortRekap2PeriodComp.Rows.Add()
                rowdgv = Me.dgvSortRekap2PeriodComp.Rows(Me.dgvSortRekap2PeriodComp.Rows.Count - 1)
                With rowdgv
                    .Cells(Me.ColName.Index).Value = "2-nd Period Amount"
                    .Cells(Me.ColSortNumber.Index).Value = "4"
                    .Cells(Me.ColSrtType.Index).Value = "Ascending"
                End With
            End If

            'atur soal sorting
            Me.cboSort1.Items.Clear()
            If Me.cboFormat.SelectedIndex = 0 Or Me.cboFormat.SelectedIndex = 2 Then
                If Me.cboGroupBy.SelectedIndex <> 0 Then
                    Me.cboSort1.Items.Add("Ascending")
                    Me.cboSort1.Items.Add("Descending")
                    Me.cboSort1.Items.Add("Amount-Ascending")
                    Me.cboSort1.Items.Add("Amount-Descending")
                Else
                    Me.cboSort1.Items.Add("Ascending")
                    Me.cboSort1.Items.Add("Descending")
                    Me.cboSort1.Items.Add("Qty-Ascending")
                    Me.cboSort1.Items.Add("Qty-Descending")
                    Me.cboSort1.Items.Add("Amount-Ascending")
                    Me.cboSort1.Items.Add("Amount-Descending")
                End If
            Else
                Me.cboSort1.Items.Add("Ascending")
                Me.cboSort1.Items.Add("Descending")
                Me.cboSort1.Items.Add("1-st Period Amount-Ascending")
                Me.cboSort1.Items.Add("1-st Period Amount-Descending")
                Me.cboSort1.Items.Add("2-nd Period Amount-Ascending")
                Me.cboSort1.Items.Add("2-nd Period Amount-Descending")
                Me.cboSort1.Items.Add("% Difference-Ascending")
                Me.cboSort1.Items.Add("% Difference-Descending")
            End If
            Me.cboSort1.SelectedIndex = 0

            If Me.cboGroupBy.Text = "Customer Code" Or Me.cboGroupBy.Text = "Date" Then
                Me.cboitemDetail.Enabled = True
                Me.cboitemDetail.SelectedIndex = 0
                Me.ChkProductDesc.Visible = True
                Me.ChkProductDesc.Checked = True
            Else
                Me.cboitemDetail.Enabled = False
                Me.cboitemDetail.SelectedIndex = -1
                Me.ChkProductDesc.Visible = False
                Me.ChkProductDesc.Checked = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPenjualan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.cboFormat.SelectedIndex = 0
            Me.cboSort1.SelectedIndex = 0

            'pastikan crystalreport terlihat
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.ShowGroupTreeButton = False

            'pastikan chk customer diisikan
            ObjCust.IsiListBoxCustomer(Me.clbCustomer, TblCust)
            Me.chkCustomer.Checked = True
            'pastikan clb product diisikan
            ObjProd.IsiListBoxProduct(Me.clbProduct, TblProd)
            Me.chkProduct.Checked = True
            'Me.IsikanDGV()
            Me.Text = "Sales List"
            Me.dgvSortDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSortRekap.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSortRekap2PeriodComp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCustomerGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Currdecimal = Me.objCurr.GetDecimal_DefaultCurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                'isikan cbo group by
                Me.cboGroupBy.Items.Clear()
                Me.cboGroupBy.Items.Add("Code No")
                Me.cboGroupBy.Items.Add("Category")
                Me.cboGroupBy.Items.Add("Sub Brand")
                Me.cboGroupBy.Items.Add("Brand")
                Me.cboGroupBy.Items.Add("Product Group")
                Me.cboGroupBy.Items.Add("Customer Code")
                Me.cboGroupBy.SelectedIndex = 0
                Me.dgvSortRekap.BringToFront()
                Me.dgvSortDetail.SendToBack()
                Me.ChkProductDesc.Visible = False
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                'isikan cbo group by
                Me.cboGroupBy.Items.Clear()
                Me.cboGroupBy.Items.Add("Code No")
                Me.cboGroupBy.Items.Add("Category")
                Me.cboGroupBy.Items.Add("Sub Brand")
                Me.cboGroupBy.Items.Add("Brand")
                Me.cboGroupBy.Items.Add("Product Group")
                Me.cboGroupBy.Items.Add("Customer Code")
                Me.cboGroupBy.SelectedIndex = 0
                Me.dgvSortRekap2PeriodComp.BringToFront()
                Me.ChkProductDesc.Visible = False
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                Me.cboGroupBy.Items.Clear()
                Me.cboGroupBy.Items.Add("Code No")
                Me.cboGroupBy.Items.Add("Category")
                Me.cboGroupBy.Items.Add("Sub Brand")
                Me.cboGroupBy.Items.Add("Brand")
                Me.cboGroupBy.Items.Add("Product Group")
                Me.cboGroupBy.Items.Add("Customer Code")
                Me.cboGroupBy.Items.Add("Date")
                Me.cboGroupBy.SelectedIndex = 0
                Me.dgvSortDetail.BringToFront()
                Me.dgvSortRekap.SendToBack()
                Me.ChkProductDesc.Visible = True
            End If

            If Me.cboFormat.SelectedIndex = 1 Then
                Me.dtpFrom2.Enabled = True
                Me.dtpTo2.Enabled = True
            Else
                Me.dtpFrom2.Enabled = False
                Me.dtpTo2.Enabled = False
            End If
            Me.IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGroupBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedIndexChanged
        Try
            Me.IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomer.CheckedChanged
        Try
            If Me.chkCustomer.Checked = True Then
                Me.clbCustomer.Enabled = False
                For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                    Me.clbCustomer.SetItemCheckState(i, CheckState.Checked)
                Next
            Else
                For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                    Me.clbCustomer.SetItemCheckState(i, CheckState.Unchecked)
                Next
                Me.clbCustomer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkProduct.CheckedChanged
        Try
            If Me.chkProduct.Checked = True Then
                Me.clbProduct.Enabled = False
                For i As Integer = 0 To Me.clbProduct.Items.Count - 1
                    Me.clbProduct.SetItemCheckState(i, CheckState.Checked)
                Next
            Else
                For i As Integer = 0 To Me.clbProduct.Items.Count - 1
                    Me.clbProduct.SetItemCheckState(i, CheckState.Unchecked)
                Next
                Me.clbProduct.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahBolehCetak() As Boolean
        Try
            If Me.clbProduct.CheckedItems.Count <= 0 Then
                MsgBox("Please at least 1 product.", MsgBoxStyle.Information, "Information")
                Me.clbProduct.Focus()
                Return False
            End If
            If Me.clbCustomer.CheckedItems.Count <= 0 Then
                MsgBox("Please at least 1 customer.", MsgBoxStyle.Information, "Information")
                Me.clbCustomer.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row

        Dim i As Integer = 0

        Try
            For Each rowReturn In Me.TBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                Me.TBL.DefaultView.RowFilter = Filter

                i = 0
                While Me.TBL.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = Me.TBL.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        Me.TBL.Rows.Remove(row1)
                    End If
                End While
            Next

            Me.TBL.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim RowView As DataRowView
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpFrom2.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom2.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo2.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo2.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.ApakahBolehCetak = True Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                Me.TBL.Rows.Clear()
                Me.TBLDetail.Rows.Clear()
                Me.TBLRekap.Rows.Clear()
                Me.TBLRekapComparison.Rows.Clear()
                Me.DS2.Rekap1.Rows.Clear()
                Me.DS2.Detail1.Rows.Clear()


                DTA.Fill(TBL, Me.dtpTo.Value, Me.dtpFrom.Value)
                Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo.Value, Me.dtpFrom.Value)
                Me.KurangkanItemReturn()
                Me.GenerateFilteringBaru()

                'set kriteria filter
                Me.ListBrand.Clear()
                Me.ListCust.Clear()
                If Me.chkProduct.Checked = False Then
                    For Each RowView In Me.clbProduct.CheckedItems
                        rowProd = RowView.Row
                        Me.ListBrand.Add(rowProd.Product_Code)
                    Next
                End If
                If Me.chkCustomer.Checked = False Then
                    For Each RowView In Me.clbCustomer.CheckedItems
                        rowCust = RowView.Row
                        Me.ListCust.Add(rowCust.Customer_Code)
                    Next
                End If

                If Me.cboFormat.SelectedIndex = 2 Then
                    If TBL.Rows.Count = 0 Then
                        MsgBox("No sales data on the database.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.GenerateDetail()
                ElseIf Me.cboFormat.SelectedIndex = 0 Then
                    If TBL.Rows.Count = 0 Then
                        MsgBox("No sales data on the database.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.GenerateRekap()
                Else
                    Me.GenerateRekap2PeriodComparison()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub GenerateRekap2PeriodComparison()
        Dim row As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPenjualan.Rekap_Comparison1DataTable
        Dim rowDummy As DseLaporanPenjualan.Rekap_Comparison1Row
        Dim rowRekap_Comparison As DseLaporanPenjualan.Rekap_ComparisonRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str As String
        Dim Lanjut As Boolean = False
        Dim itemDetail As Integer = 0
        Dim Name2_Str As String = ""
        Try
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
                'ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                '    GroupBy = 2
            End If
            For Each row In Me.TBL.Rows
                Lanjut = True
                If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
                    Lanjut = True
                Else
                    If Me.ApakahMasukKriteria(row) <> True Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.cboitemDetail.Text <> "" Then
                        If Me.cboitemDetail.SelectedIndex = 0 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        End If
                        itemDetail = Me.cboitemDetail.SelectedIndex
                    Else
                        If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 5 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = 0
                        ElseIf Me.cboGroupBy.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        End If
                    End If

                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.NewRekap_Comparison1Row
                        rowDummy.Amount1 = row.Sales_Amount
                        rowDummy.Cust_Code = row.Customer_Code + " - " + row.Customer_Name
                        If itemDetail = 0 Then ' product
                            rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                        ElseIf itemDetail = 1 Then 'category
                            rowDummy.Product_Code = row.Category + " - " + row.Category_Desc
                        ElseIf itemDetail = 2 Then 'sub brand
                            rowDummy.Product_Code = row.Sub_Brand + " - " + row.Sub_Brand_Desc
                        ElseIf itemDetail = 3 Then 'brand
                            rowDummy.Product_Code = row.Brand + " - " + row.Brand_Desc
                        ElseIf itemDetail = 4 Then 'product group
                            rowDummy.Product_Code = row.Product_Group + " - " + row.Product_Group_Desc
                        End If
                        rowDummy.Qty1 = row.Total_Qty
                        rowDummy._Decimal = row._Decimal
                        rowDummy.Unit = row.Unit_Code
                        rowDummy.Product_Desc = row.Description
                        rowDummy.Amount2 = 0
                        rowDummy.Qty2 = 0
                        rowDummy.Persen_Sub_Total = 0
                        rowDummy.Persen_Total = 0
                        rowDummy.Persen_Amount = 0
                        TBLDummy.Rows.Add(rowDummy)
                    Else
                        rowDummy = TBLDummy.DefaultView(0).Row
                        rowDummy.Amount1 += row.Sales_Amount
                        rowDummy.Qty1 += row.Total_Qty
                    End If
                End If
            Next

            'isikan data periode 2
            Me.TBL.Rows.Clear()
            Me.TBLReturn.Rows.Clear()
            Me.DTA.Fill(Me.TBL, Me.dtpTo2.Value, Me.dtpFrom2.Value)
            Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo.Value, Me.dtpFrom.Value)
            Me.KurangkanItemReturn()
            Me.GenerateFilteringBaru()

            For Each rowView In Me.TBL.DefaultView
                Lanjut = True
                row = rowView.Row
                If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
                    Lanjut = True
                Else
                    If Me.ApakahMasukKriteria(row) <> True Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.cboitemDetail.Text <> "" Then
                        If Me.cboitemDetail.SelectedIndex = 0 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        End If
                        itemDetail = Me.cboitemDetail.SelectedIndex
                    Else
                        If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 5 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = 0
                        ElseIf Me.cboGroupBy.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        End If
                    End If

                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.NewRekap_Comparison1Row
                        rowDummy.Amount2 = row.Sales_Amount
                        rowDummy.Cust_Code = row.Customer_Code + " - " + row.Customer_Name
                        If itemDetail = 0 Then ' product
                            rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                        ElseIf itemDetail = 1 Then 'category
                            rowDummy.Product_Code = row.Category + " - " + row.Category_Desc
                        ElseIf itemDetail = 2 Then 'sub brand
                            rowDummy.Product_Code = row.Sub_Brand + " - " + row.Sub_Brand_Desc
                        ElseIf itemDetail = 3 Then 'brand
                            rowDummy.Product_Code = row.Brand + " - " + row.Brand_Desc
                        ElseIf itemDetail = 4 Then 'product group
                            rowDummy.Product_Code = row.Product_Group + " - " + row.Product_Group_Desc
                        End If

                        rowDummy.Qty2 = row.Total_Qty
                        rowDummy._Decimal = row._Decimal
                        rowDummy.Unit = row.Unit_Code
                        rowDummy.Product_Desc = row.Description
                        rowDummy.Amount1 = 0
                        rowDummy.Qty1 = 0
                        rowDummy.Persen_Sub_Total = 0
                        rowDummy.Persen_Total = 0
                        rowDummy.Persen_Amount = 0
                        TBLDummy.Rows.Add(rowDummy)
                    Else
                        rowDummy = TBLDummy.DefaultView(0).Row
                        rowDummy.Amount2 += row.Sales_Amount
                        rowDummy.Qty2 += row.Total_Qty
                    End If
                End If
            Next

            If TBLDummy.Rows.Count = 0 Then
                MsgBox("No sales data on the database.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'hitung total group
            Me.HitungSubTotalQtyAmount(TBLDummy)

            'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = Me.GenerateSortRekap_Comparison
            For Each rowView In TBLDummy.DefaultView
                rowDummy = rowView.Row
                rowRekap_Comparison = Me.TBLRekapComparison.NewRekap_ComparisonRow
                If GroupBy = EnumGroupBy.CodeNo Then
                    rowRekap_Comparison.Group_By = rowDummy.Product_Code
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    rowRekap_Comparison.Group_By = rowDummy.Cust_Code
                End If

                'name1
                Name1 = Me.GetName1_Rekap
                If Name1 = EnumGroupBy.CodeNo Then
                    rowRekap_Comparison.Name_1 = rowDummy.Product_Code
                ElseIf Name1 = EnumGroupBy.CustCode Then
                    rowRekap_Comparison.Name_1 = rowDummy.Cust_Code
                End If


                rowRekap_Comparison.Qty1 = rowDummy.Qty1
                rowRekap_Comparison.Amount1 = rowDummy.Amount1
                rowRekap_Comparison.Qty2 = rowDummy.Qty2
                rowRekap_Comparison.Amount2 = rowDummy.Amount2
                rowRekap_Comparison.Persen_Amount = rowDummy.Persen_Amount
                rowRekap_Comparison.Persen_Sub_Total = rowDummy.Persen_Sub_Total
                rowRekap_Comparison.Persen_Total = rowDummy.Persen_Total

                rowRekap_Comparison._Decimal = rowDummy._Decimal
                rowRekap_Comparison.Unit = rowDummy.Unit
                Me.TBLRekapComparison.Rows.Add(rowRekap_Comparison)
            Next

            Name1_Str = ""
            If Name1 = EnumGroupBy.CodeNo Then
                If itemDetail = 0 Then ' product
                    Name1_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name1_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name1_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name1_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name1_Str = "Product Group"
                End If
            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Customer Code"
                If itemDetail = 0 Then ' product
                    Name2_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name2_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name2_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name2_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name2_Str = "Product Group"
                End If
            End If


            ReportRekap_Comparison.SetDataSource(Me.DS2)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_ReportName.ParameterFieldName, "Sales List")
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate1.ParameterFieldName, Me.dtpFrom2.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate1.ParameterFieldName, Me.dtpTo2.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Name1.ParameterFieldName, Name1_Str)

            If GroupBy = EnumGroupBy.CodeNo Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, Name2_Str)
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Customer")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Date")
            End If
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Default_Precision.ParameterFieldName, Me.Currdecimal)

            Me.crv1.ReportSource = ReportRekap_Comparison
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()

            Me.TabControl1.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateRekap()
        Dim row As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPenjualan.Rekap1DataTable
        Dim rowDummy As DseLaporanPenjualan.Rekap1Row
        Dim rowRekap As DseLaporanPenjualan.RekapRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str As String
        Dim Lanjut As Boolean = False

        Dim itemDetail As Integer = 0
        Dim Name2_Str As String = ""
        Try
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
                'ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                '    GroupBy = 2
            End If
            For Each row In Me.TBL.Rows
                Lanjut = True
                If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
                    Lanjut = True
                Else
                    If Me.ApakahMasukKriteria(row) <> True Then
                        Lanjut = False
                    End If
                End If

                If Lanjut = True Then
                    If Me.cboitemDetail.Text <> "" Then
                        If Me.cboitemDetail.SelectedIndex = 0 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        ElseIf Me.cboitemDetail.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                        End If
                        itemDetail = Me.cboitemDetail.SelectedIndex
                    Else
                        If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 5 Then 'Code No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = 0
                        ElseIf Me.cboGroupBy.SelectedIndex = 1 Then 'Category No
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 2 Then 'Sub Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 3 Then 'Brand
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        ElseIf Me.cboGroupBy.SelectedIndex = 4 Then 'Product Group
                            TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + _
                            "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            itemDetail = Me.cboGroupBy.SelectedIndex
                        End If
                    End If
                    
                   

                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.NewRekap1Row
                        rowDummy.Amount = row.Sales_Amount
                        rowDummy.Cust_Code = row.Customer_Code + " - " + row.Customer_Name
                        If itemDetail = 0 Then ' product
                            rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                        ElseIf itemDetail = 1 Then 'category
                            rowDummy.Product_Code = row.Category + " - " + row.Category_Desc
                        ElseIf itemDetail = 2 Then 'sub brand
                            rowDummy.Product_Code = row.Sub_Brand + " - " + row.Sub_Brand_Desc
                        ElseIf itemDetail = 3 Then 'brand
                            rowDummy.Product_Code = row.Brand + " - " + row.Brand_Desc
                        ElseIf itemDetail = 4 Then 'product group
                            rowDummy.Product_Code = row.Product_Group + " - " + row.Product_Group_Desc
                        End If

                        rowDummy.Qty = row.Total_Qty
                        rowDummy._Decimal = row._Decimal
                        rowDummy.Unit = row.Unit_Code
                        rowDummy.Product_Desc = row.Description
                        TBLDummy.Rows.Add(rowDummy)
                    Else
                        rowDummy = TBLDummy.DefaultView(0).Row
                        rowDummy.Amount += row.Sales_Amount
                        rowDummy.Qty += row.Total_Qty
                    End If
                End If
            Next

            'hitung total group
            If Me.cboSort1.SelectedIndex <> 0 And Me.cboSort1.SelectedIndex <> 1 Then
                Me.HitungSubTotalQtyAmount(TBLDummy)
            End If

            'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = Me.GenerateSortRekap
            For Each rowView In TBLDummy.DefaultView
                rowDummy = rowView.Row
                rowRekap = Me.TBLRekap.NewRekapRow
                If GroupBy = EnumGroupBy.CodeNo Then
                    rowRekap.Group_By = rowDummy.Product_Code
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    rowRekap.Group_By = rowDummy.Cust_Code
                End If

                'name1
                Name1 = Me.GetName1_Rekap
                If Name1 = EnumGroupBy.CodeNo Then
                    rowRekap.Name_1 = rowDummy.Product_Code
                ElseIf Name1 = EnumGroupBy.CustCode Then
                    rowRekap.Name_1 = rowDummy.Cust_Code
                End If


                rowRekap.Qty = rowDummy.Qty
                rowRekap.Amount = rowDummy.Amount
                rowRekap._Decimal = rowDummy._Decimal
                rowRekap.Unit = rowDummy.Unit
                Me.TBLRekap.Rows.Add(rowRekap)
            Next

            Name1_Str = ""
            If Name1 = EnumGroupBy.CodeNo Then
                If itemDetail = 0 Then ' product
                    Name1_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name1_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name1_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name1_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name1_Str = "Product Group"
                End If
            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Customer Code"
                If itemDetail = 0 Then ' product
                    Name2_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name2_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name2_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name2_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name2_Str = "Product Group"
                End If
            End If



            ReportRekap.SetDataSource(Me.DS2)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Sales List")
            ReportRekap.SetParameterValue(ReportRekap.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_Name1.ParameterFieldName, Name1_Str)
            If GroupBy = EnumGroupBy.CodeNo Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, Name2_Str)
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Customer")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Date")
            End If
            ReportRekap.SetParameterValue(ReportRekap.Parameter_Default_Precision.ParameterFieldName, Me.Currdecimal)

            Me.crv1.ReportSource = ReportRekap
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()

            Me.TabControl1.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    


    Private Function ApakahMasukKriteria(ByVal row1 As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row) As Boolean
        Dim ketemu As Boolean = False
        Dim RowView As DataRowView

        Try
            'cek Prod
            If Me.chkProduct.Checked = False Then
                ketemu = False
                For Each RowView In Me.clbProduct.CheckedItems
                    rowProd = RowView.Row
                    If rowProd.Product_Code = row1.Product_Code Then
                        ketemu = True
                    End If
                Next
                If ketemu = False Then
                    Return False
                End If
            End If
            'cek cust code
            If Me.chkCustomer.Checked = False Then
                ketemu = False
                For Each RowView In Me.clbCustomer.CheckedItems
                    rowCust = RowView.Row
                    If rowCust.Customer_Code = row1.Customer_Code Then
                        ketemu = True
                    End If
                Next
                If ketemu = False Then
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GenerateSortDetail2() As String
        Dim SortString As String = ""
        Dim GroupBy As EnumGroupBy
        Dim Sort1, SortType1 As String
        Dim ListSort As New List(Of String)
        Dim ListSortType As New List(Of String)
        Dim Sort, SortType As String

        Try
            Sort1 = ""
            SortType1 = ""

            'pertama sort berdasarkan grouping
            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
           Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
            Else
                GroupBy = 2
            End If

            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Sort1 = "[Cust Code]"
                Else
                    Sort1 = "Tanggal"
                End If
                If Me.cboSort1.SelectedIndex = 0 Then
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 1 Then
                    SortType1 = "DESC"
                End If
            Else
                If GroupBy = EnumGroupBy.CodeNo Then
                    If Me.cboSort1.SelectedIndex = 2 Then
                        Sort1 = "[Total Qty Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 3 Then
                        Sort1 = "[Total Qty Group By]"
                        SortType1 = "DESC"
                    ElseIf Me.cboSort1.SelectedIndex = 4 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 5 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "DESC"
                    End If
                Else
                    If Me.cboSort1.SelectedIndex = 2 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 3 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "DESC"
                    End If
                End If
            End If

            For i As Integer = 1 To 5
                Sort = Me.GetColNameBySortNumber(i)
                SortType = Me.GetSortTypeBySortNumber(i)

                ListSort.Add(Sort)
                ListSortType.Add(SortType)
            Next

            'tambahkan sort berdasarkan dgv
            Return Sort1 + " " + SortType1 + ", " + ListSort.Item(0) + " " + ListSortType.Item(0) + ", " + ListSort.Item(1) + " " + ListSortType.Item(1) + ", " + _
            ListSort.Item(2) + " " + ListSortType.Item(2) + ", " + ListSort.Item(3) + " " + ListSortType.Item(3) + ", " + ListSort.Item(4) + " " + ListSortType.Item(4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListSort = Nothing
            ListSortType = Nothing
        End Try
    End Function

    Private Function GenerateSortRekap() As String
        Dim SortString As String = ""
        Dim GroupBy As EnumGroupBy
        Dim Sort1, SortType1 As String
        Dim ListSort As New List(Of String)
        Dim ListSortType As New List(Of String)
        Dim Sort, SortType As String

        Try
            Sort1 = ""
            SortType1 = ""

            'pertama sort berdasarkan grouping
            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            Else
                GroupBy = 1
            End If
            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Sort1 = "[Cust Code]"
                End If
                If Me.cboSort1.SelectedIndex = 0 Then
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 1 Then
                    SortType1 = "DESC"
                End If
            Else
                If GroupBy = EnumGroupBy.CodeNo Then
                    If Me.cboSort1.SelectedIndex = 2 Then
                        Sort1 = "[Total Qty Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 3 Then
                        Sort1 = "[Total Qty Group By]"
                        SortType1 = "DESC"
                    ElseIf Me.cboSort1.SelectedIndex = 4 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 5 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "DESC"
                    End If
                Else
                    If Me.cboSort1.SelectedIndex = 2 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "ASC"
                    ElseIf Me.cboSort1.SelectedIndex = 3 Then
                        Sort1 = "[Total Amount Group By]"
                        SortType1 = "DESC"
                    End If
                End If
            End If

            'tambahkan sort berdasarkan dgv
            If GroupBy = EnumGroupBy.CodeNo Then
                For i As Integer = 1 To 3
                    Sort = Me.GetColNameBySortNumberRekap(i)
                    SortType = Me.GetSortTypeBySortNumberRekap(i)

                    ListSort.Add(Sort)
                    ListSortType.Add(SortType)
                Next

                Return Sort1 + " " + SortType1 + ", " + ListSort.Item(0) + " " + ListSortType.Item(0) _
                + ", " + ListSort.Item(1) + " " + ListSortType.Item(1) _
                + ", " + ListSort.Item(2) + " " + ListSortType.Item(2)
            Else
                For i As Integer = 1 To 2
                    Sort = Me.GetColNameBySortNumberRekap(i)
                    SortType = Me.GetSortTypeBySortNumberRekap(i)

                    ListSort.Add(Sort)
                    ListSortType.Add(SortType)
                Next

                Return Sort1 + " " + SortType1 + ", " + ListSort.Item(0) + " " + ListSortType.Item(0) _
                + ", " + ListSort.Item(1) + " " + ListSortType.Item(1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListSort = Nothing
            ListSortType = Nothing
        End Try
    End Function

    Private Function GenerateSortRekap_Comparison() As String
        Dim SortString As String = ""
        Dim GroupBy As EnumGroupBy
        Dim Sort1, SortType1 As String
        Dim ListSort As New List(Of String)
        Dim ListSortType As New List(Of String)
        Dim Sort, SortType As String

        Try
            Sort1 = ""
            SortType1 = ""

            'pertama sort berdasarkan grouping
            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            Else
                GroupBy = 1
            End If
            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Sort1 = "[Cust Code]"
                End If
                If Me.cboSort1.SelectedIndex = 0 Then
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 1 Then
                    SortType1 = "DESC"
                End If
            Else
                If Me.cboSort1.SelectedIndex = 2 Then
                    Sort1 = "[Total Amount 1]"
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 3 Then
                    Sort1 = "[Total Amount 1]"
                    SortType1 = "DESC"
                ElseIf Me.cboSort1.SelectedIndex = 4 Then
                    Sort1 = "[Total Amount 2]"
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 5 Then
                    Sort1 = "[Total Amount 2]"
                    SortType1 = "DESC"
                ElseIf Me.cboSort1.SelectedIndex = 6 Then
                    Sort1 = "[Persen Sub Total]"
                    SortType1 = "ASC"
                ElseIf Me.cboSort1.SelectedIndex = 7 Then
                    Sort1 = "[Persen Sub Total]"
                    SortType1 = "DESC"
                End If
            End If


            'tambahkan sort berdasarkan dgv
            For i As Integer = 1 To 4
                Sort = Me.GetColNameBySortNumberRekap_Comparison(i)
                SortType = Me.GetSortTypeBySortNumberRekap_Comparison(i)

                ListSort.Add(Sort)
                ListSortType.Add(SortType)
            Next

            Return Sort1 + " " + SortType1 + ", " + ListSort.Item(0) + " " + ListSortType.Item(0) _
            + ", " + ListSort.Item(1) + " " + ListSortType.Item(1) _
            + ", " + ListSort.Item(2) + " " + ListSortType.Item(2) _
            + ", " + ListSort.Item(3) + " " + ListSortType.Item(3)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListSort = Nothing
            ListSortType = Nothing
        End Try
    End Function

    Private Function GetColNameBySortNumberRekap_Comparison(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortRekap2PeriodComp.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "% Difference" Then
                        Return "[Persen Amount]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "1-st Period Amount" Then
                        Return "Amount1"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "2-nd Period Amount" Then
                        Return "Amount2"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Or rowDGV.Cells(Me.ColName.Index).Value = "Category" Or rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" _
                    Or rowDGV.Cells(Me.ColName.Index).Value = "Brand" Or rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "Product Code"
                    Else
                        Return rowDGV.Cells(Me.ColName.Index).Value.ToString
                    End If
                End If
            Next

            Return "[Product Code]"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetSortTypeBySortNumberRekap_Comparison(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortRekap2PeriodComp.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColSrtType.Index).Value = "Ascending" Then
                        Return "ASC"
                    Else
                        Return "DESC"
                    End If
                End If
            Next

            Return "ASC"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "ASC"
        End Try
    End Function

    Private Function GetColNameBySortNumberRekap(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return "[Invoice Number]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return "Tanggal"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Or rowDGV.Cells(Me.ColName.Index).Value = "Category" Or rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" _
                    Or rowDGV.Cells(Me.ColName.Index).Value = "Brand" Or rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "Product Code"
                    Else
                        Return rowDGV.Cells(Me.ColName.Index).Value.ToString
                    End If
                End If
            Next

            Return "[Product Code]"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetSortTypeBySortNumberRekap(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColSrtType.Index).Value = "Ascending" Then
                        Return "ASC"
                    Else
                        Return "DESC"
                    End If
                End If
            Next

            Return "ASC"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "ASC"
        End Try
    End Function

    Private Function GetColNameBySortNumber(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return "[Invoice Number]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return "Tanggal"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Or rowDGV.Cells(Me.ColName.Index).Value = "Category" Or rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" _
                    Or rowDGV.Cells(Me.ColName.Index).Value = "Brand" Or rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "Product Code"
                    Else
                        Return rowDGV.Cells(Me.ColName.Index).Value.ToString
                    End If
                End If
            Next

            Return "Product Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetSortTypeBySortNumber(ByVal vSortNumber As Integer) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = vSortNumber Then
                    If rowDGV.Cells(Me.ColSrtType.Index).Value = "Ascending" Then
                        Return "ASC"
                    Else
                        Return "DESC"
                    End If
                End If
            Next

            Return "ASC"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "ASC"
        End Try
    End Function

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPenjualan.Detail1DataTable)
        Dim rowDummy As DseLaporanPenjualan.Detail1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Qty, Amount As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
                        Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
            Else
                GroupBy = 2
            End If
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Tanggal Then
                    If ListFilter.Contains(Format(CDate(rowDummy.Tanggal), "dd MMM yyyy")) = False Then
                        ListFilter.Add(Format(CDate(rowDummy.Tanggal), "dd MMM yyyy"))
                    End If
                End If
            Next

            For Each Filter In ListFilter
                If GroupBy = EnumGroupBy.Tanggal Then
                    Filter = "[Tanggal] = #" + CDate(Filter).ToShortDateString + "#"
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                End If

                TBLDummy.DefaultView.RowFilter = Filter
                Qty = 0
                Amount = 0
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    Qty += rowDummy.Qty
                    Amount += rowDummy.Amount
                Next
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    rowDummy.Total_Qty_Group_By = Qty
                    rowDummy.Total_Amount_Group_By = Amount
                Next
            Next

            TBLDummy.DefaultView.RowFilter = FilterLama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListFilter = Nothing
        End Try
    End Sub

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPenjualan.Rekap1DataTable)
        Dim rowDummy As DseLaporanPenjualan.Rekap1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Qty, Amount As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
                'ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                '    GroupBy = 2
            End If
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                End If
            Next

            For Each Filter In ListFilter
                If GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                End If

                TBLDummy.DefaultView.RowFilter = Filter
                Qty = 0
                Amount = 0
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    Qty += rowDummy.Qty
                    Amount += rowDummy.Amount
                Next
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    rowDummy.Total_Qty_Group_By = Qty
                    rowDummy.Total_Amount_Group_By = Amount
                Next
            Next

            TBLDummy.DefaultView.RowFilter = FilterLama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListFilter = Nothing
        End Try
    End Sub

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPenjualan.Rekap_Comparison1DataTable)
        Dim rowDummy As DseLaporanPenjualan.Rekap_Comparison1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Amount1, Amount2, Tot1, Tot2 As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            'GroupBy = Me.cboGroupBy.SelectedIndex
            'GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
            Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
                'ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                '    GroupBy = 2
            End If
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                End If
            Next

            Tot1 = 0
            Tot2 = 0
            For Each Filter In ListFilter
                If GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                End If

                TBLDummy.DefaultView.RowFilter = Filter
                Amount1 = 0
                Amount2 = 0
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    Amount1 += rowDummy.Amount1
                    Tot1 += rowDummy.Amount1
                    Amount2 += rowDummy.Amount2
                    Tot2 += rowDummy.Amount2
                Next
                For Each rowView In TBLDummy.DefaultView
                    rowDummy = rowView.Row
                    rowDummy.Total_Amount_1 = Amount1
                    rowDummy.Total_Amount_2 = Amount2
                    If rowDummy.Total_Amount_1 <> 0 Then
                        rowDummy.Persen_Sub_Total = Math.Round(((rowDummy.Total_Amount_2 - _
                        rowDummy.Total_Amount_1) / rowDummy.Total_Amount_1) * 100, 2, _
                        MidpointRounding.AwayFromZero)
                    Else
                        rowDummy.Persen_Sub_Total = 0
                    End If
                Next
            Next
            For Each rowDummy In TBLDummy.Rows
                If Tot1 <> 0 Then
                    rowDummy.Persen_Total = Math.Round(((Tot2 - Tot1) / Tot1) * 100, 2, MidpointRounding.AwayFromZero)
                Else
                    rowDummy.Persen_Total = 0
                End If
                If rowDummy.Amount1 <> 0 Then
                    rowDummy.Persen_Amount = Math.Round(((rowDummy.Amount2 - rowDummy.Amount1) _
                    / rowDummy.Amount1) * 100, 2, MidpointRounding.AwayFromZero)
                Else
                    rowDummy.Persen_Amount = 0
                End If
            Next
            TBLDummy.DefaultView.RowFilter = FilterLama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListFilter = Nothing
        End Try
    End Sub

    Private Sub GenerateDetail()
        Dim row As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPenjualan.Detail1DataTable
        Dim rowDummy As DseLaporanPenjualan.Detail1Row
        Dim rowDetail As DseLaporanPenjualan.DetailRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1, Name2, Name3 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str, Name2_Str, Name3_Str As String
        Dim Lanjut As Boolean = False
        Dim itemDetail As Integer = 0
        Dim Name4_Str As String
        Try
            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 _
           Or Me.cboGroupBy.SelectedIndex = 4 Then
                GroupBy = 0
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                GroupBy = 1
            ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                GroupBy = 2
            End If
            For Each row In Me.TBL.Rows
                Lanjut = True
                If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
                    Lanjut = True
                Else
                    If Me.ApakahMasukKriteria(row) <> True Then
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    If Me.ChkProductDesc.Checked = True Then
                        If Me.cboitemDetail.Text <> "" Then
                            If Me.cboitemDetail.SelectedIndex = 0 Then 'Code No
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 1 Then 'Category No
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                               "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 2 Then 'Sub Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                               "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 3 Then 'Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                               "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 4 Then 'product Group
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                               "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                            End If
                            itemDetail = Me.cboitemDetail.SelectedIndex
                        Else
                            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 5 Or Me.cboGroupBy.SelectedIndex = 6 Then 'Code No
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Description) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                                itemDetail = 0
                            ElseIf Me.cboGroupBy.SelectedIndex = 1 Then 'Category No
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 2 Then 'Sub Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 3 Then 'Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 4 Then 'Product Group
                                TBLDummy.DefaultView.RowFilter = "[Product Desc] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group_Desc) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "' AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            End If
                        End If
                        
                    Else
                        If Me.cboitemDetail.Text <> "" Then
                            If Me.cboitemDetail.SelectedIndex = 0 Then 'Code No
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 1 Then 'Category No
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 2 Then 'Sub Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 3 Then 'Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                            ElseIf Me.cboitemDetail.SelectedIndex = 4 Then 'product Group
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                            End If
                            itemDetail = Me.cboitemDetail.SelectedIndex
                        Else
                            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 5 Or Me.cboGroupBy.SelectedIndex = 6 Then 'Code No
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'"
                                itemDetail = 0
                            ElseIf Me.cboGroupBy.SelectedIndex = 1 Then 'Category No
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Category) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 2 Then 'Sub Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 3 Then 'Brand
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Brand) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            ElseIf Me.cboGroupBy.SelectedIndex = 4 Then 'Product Group
                                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Product_Group) + "' AND Tanggal =#" + row.Invoice_Date.ToShortDateString + "# AND " + _
                                "[Invoice Number] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.No_Transaksi) + "' AND [Cust Code] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row.Customer_Name) + "'AND [Unit] = '" & row.Unit_Code & "'"
                                itemDetail = Me.cboGroupBy.SelectedIndex
                            End If
                        End If
                        
                    End If

                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.NewDetail1Row
                        rowDummy.Amount = row.Sales_Amount
                        rowDummy.Cust_Code = row.Customer_Name
                        rowDummy.Invoice_Number = row.No_Transaksi

                        If Me.ChkProductDesc.Checked = True Then
                            If itemDetail = 0 Then ' product
                                If GroupBy = EnumGroupBy.CustCode Or GroupBy = EnumGroupBy.Tanggal Then
                                    rowDummy.Product_Code = row.Description
                                    rowDummy.Product_Desc = row.Description
                                Else
                                    rowDummy.Product_Code = row.Product_Code
                                    rowDummy.Product_Desc = row.Description
                                End If
                               
                            ElseIf itemDetail = 1 Then 'category
                                If GroupBy = EnumGroupBy.CustCode Or GroupBy = EnumGroupBy.Tanggal Then
                                    rowDummy.Product_Code = row.Category_Desc
                                    rowDummy.Product_Desc = row.Category_Desc
                                Else
                                    rowDummy.Product_Code = row.Category
                                    rowDummy.Product_Desc = row.Category_Desc
                                End If
                               
                            ElseIf itemDetail = 2 Then 'sub brand
                                If GroupBy = EnumGroupBy.CustCode Or GroupBy = EnumGroupBy.Tanggal Then
                                    rowDummy.Product_Code = row.Sub_Brand_Desc
                                    rowDummy.Product_Desc = row.Sub_Brand_Desc
                                Else
                                    rowDummy.Product_Code = row.Sub_Brand
                                    rowDummy.Product_Desc = row.Sub_Brand_Desc
                                End If
                            ElseIf itemDetail = 3 Then 'brand
                                If GroupBy = EnumGroupBy.CustCode Or GroupBy = EnumGroupBy.Tanggal Then
                                    rowDummy.Product_Code = row.Brand_Desc
                                    rowDummy.Product_Desc = row.Brand_Desc
                                Else
                                    rowDummy.Product_Code = row.Brand
                                    rowDummy.Product_Desc = row.Brand_Desc
                                End If
                            ElseIf itemDetail = 4 Then 'product group
                                If GroupBy = EnumGroupBy.CustCode Or GroupBy = EnumGroupBy.Tanggal Then
                                    rowDummy.Product_Code = row.Product_Group_Desc
                                    rowDummy.Product_Desc = row.Product_Group_Desc
                                Else
                                    rowDummy.Product_Code = row.Product_Group
                                    rowDummy.Product_Desc = row.Product_Group_Desc
                                End If
                            End If
                        Else
                            If itemDetail = 0 Then ' product
                                rowDummy.Product_Code = row.Product_Code
                                rowDummy.Product_Desc = row.Description
                            ElseIf itemDetail = 1 Then 'category
                                rowDummy.Product_Code = row.Category
                                rowDummy.Product_Desc = row.Category_Desc
                            ElseIf itemDetail = 2 Then 'sub brand
                                rowDummy.Product_Code = row.Sub_Brand
                                rowDummy.Product_Desc = row.Sub_Brand_Desc
                            ElseIf itemDetail = 3 Then 'brand
                                rowDummy.Product_Code = row.Brand
                                rowDummy.Product_Desc = row.Brand_Desc
                            ElseIf itemDetail = 4 Then 'product group
                                rowDummy.Product_Code = row.Product_Group
                                rowDummy.Product_Desc = row.Product_Group_Desc
                            End If
                        End If
                        rowDummy.Qty = row.Total_Qty
                        rowDummy.Tanggal = Format(row.Invoice_Date, "dd MMM yyyy")
                        rowDummy._Decimal = row._Decimal
                        rowDummy.Unit = row.Unit_Code

                        TBLDummy.Rows.Add(rowDummy)
                    Else
                        rowDummy = TBLDummy.DefaultView(0).Row
                        rowDummy.Amount += row.Sales_Amount
                        rowDummy.Qty += row.Total_Qty
                    End If
                End If
            Next

            'hitung total group
            If Me.cboSort1.SelectedIndex <> 0 And Me.cboSort1.SelectedIndex <> 1 Then
                Me.HitungSubTotalQtyAmount(TBLDummy)
            End If

            'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = Me.GenerateSortDetail2
            For Each rowView In TBLDummy.DefaultView
                rowDummy = rowView.Row
                rowDetail = Me.TBLDetail.NewDetailRow
                If GroupBy = EnumGroupBy.CodeNo Then
                    rowDetail.Group_By = rowDummy.Product_Code + " - " + rowDummy.Product_Desc
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    rowDetail.Group_By = rowDummy.Cust_Code
                ElseIf GroupBy = EnumGroupBy.Tanggal Then
                    rowDetail.Group_By = rowDummy.Tanggal
                End If

                'name1
                Name1 = Me.GetName1
                If Name1 = EnumGroupBy.CodeNo Then
                    rowDetail.Name_1 = rowDummy.Product_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Code No")
                ElseIf Name1 = EnumGroupBy.CustCode Then
                    rowDetail.Name_1 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Customer Code")
                ElseIf Name1 = EnumGroupBy.Invoice Then
                    rowDetail.Name_1 = rowDummy.Invoice_Number
                    IndexMinimum = Me.GetSortNumberMinimum("Invoice Number")
                ElseIf Name1 = EnumGroupBy.Tanggal Then
                    rowDetail.Name_1 = rowDummy.Tanggal
                    IndexMinimum = Me.GetSortNumberMinimum("Date")
                End If

                'name2
                Name2 = Me.GetName2_3(IndexMinimum)
                If Name2 = EnumGroupBy.CodeNo Then
                    rowDetail.Name_2 = rowDummy.Product_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Code No")
                ElseIf Name2 = EnumGroupBy.CustCode Then
                    rowDetail.Name_2 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Customer Code")
                ElseIf Name2 = EnumGroupBy.Invoice Then
                    rowDetail.Name_2 = rowDummy.Invoice_Number
                    IndexMinimum = Me.GetSortNumberMinimum("Invoice Number")
                ElseIf Name2 = EnumGroupBy.Tanggal Then
                    rowDetail.Name_2 = rowDummy.Tanggal
                    IndexMinimum = Me.GetSortNumberMinimum("Date")
                End If

                'name 3
                Name3 = Me.GetName2_3(IndexMinimum)
                If Name3 = EnumGroupBy.CodeNo Then
                    rowDetail.Name_3 = rowDummy.Product_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Code No")
                ElseIf Name3 = EnumGroupBy.CustCode Then
                    rowDetail.Name_3 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Customer Code")
                ElseIf Name3 = EnumGroupBy.Invoice Then
                    rowDetail.Name_3 = rowDummy.Invoice_Number
                    IndexMinimum = Me.GetSortNumberMinimum("Invoice Number")
                ElseIf Name3 = EnumGroupBy.Tanggal Then
                    rowDetail.Name_3 = rowDummy.Tanggal
                    IndexMinimum = Me.GetSortNumberMinimum("Date")
                End If

                rowDetail.Qty = rowDummy.Qty
                rowDetail.Amount = rowDummy.Amount
                rowDetail._Decimal = rowDummy._Decimal
                rowDetail.Unit = rowDummy.Unit
                Me.TBLDetail.Rows.Add(rowDetail)
            Next

            Name1_Str = ""
            Name2_Str = ""
            Name3_Str = ""
            Name4_Str = ""
            If Name1 = EnumGroupBy.CodeNo Then
                If itemDetail = 0 Then ' product
                    Name1_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name1_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name1_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name1_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name1_Str = "Product Group"
                End If

            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Customer Code"
                If itemDetail = 0 Then ' product
                    Name4_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name4_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name4_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name4_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name4_Str = "Product Group"
                End If
            ElseIf Name1 = EnumGroupBy.Invoice Then
                Name1_Str = "Invoice"
            ElseIf Name1 = EnumGroupBy.Tanggal Then
                Name1_Str = "Date"
            End If
            If Name2 = EnumGroupBy.CodeNo Then

                If itemDetail = 0 Then ' product
                    Name2_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name2_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name2_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name2_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name2_Str = "Product Group"
                End If
            ElseIf Name2 = EnumGroupBy.CustCode Then
                Name2_Str = "Customer Code"
                If itemDetail = 0 Then ' product
                    Name4_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name4_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name4_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name4_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name4_Str = "Product Group"
                End If
            ElseIf Name2 = EnumGroupBy.Invoice Then
                Name2_Str = "Invoice"
            ElseIf Name2 = EnumGroupBy.Tanggal Then
                Name2_Str = "Date"
            End If
            If Name3 = EnumGroupBy.CodeNo Then
                If itemDetail = 0 Then ' product
                    Name3_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name3_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name3_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name3_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name3_Str = "Product Group"
                End If
            ElseIf Name3 = EnumGroupBy.CustCode Then
                Name3_Str = "Customer Code"
                If itemDetail = 0 Then ' product
                    Name4_Str = "Code No"
                ElseIf itemDetail = 1 Then 'category
                    Name4_Str = "Category"
                ElseIf itemDetail = 2 Then 'sub brand
                    Name4_Str = "Sub Brand"
                ElseIf itemDetail = 3 Then 'brand
                    Name4_Str = "Brand"
                ElseIf itemDetail = 4 Then 'product group
                    Name4_Str = "Product Group"
                End If
            ElseIf Name3 = EnumGroupBy.Invoice Then
                Name3_Str = "Invoice"
            ElseIf Name3 = EnumGroupBy.Tanggal Then
                Name3_Str = "Date"
            End If

            ReportDetail.SetDataSource(Me.DS2)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_ReportName.ParameterFieldName, "Sales List")
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name1.ParameterFieldName, Name1_Str)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name2.ParameterFieldName, Name2_Str)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name3.ParameterFieldName, Name3_Str)
            If GroupBy = EnumGroupBy.CodeNo Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, Name4_Str)
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Customer")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Date")
            End If
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Default_Precision.ParameterFieldName, Me.Currdecimal)

            Me.crv1.ReportSource = ReportDetail
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()

            Me.TabControl1.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetName1() As EnumGroupBy
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 1 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetName1_Rekap() As EnumGroupBy
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 1 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetSortNumberMinimum(ByVal ColName As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColName.Index).Value = ColName Then
                    Return rowDGV.Cells(Me.ColSortNumber.Index).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetName2_3(ByVal SortNumberMinimum As Integer) As EnumGroupBy
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = SortNumberMinimum + 1 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = SortNumberMinimum + 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = SortNumberMinimum + 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Customer Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Invoice Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next

            'cari yang paling pertama keluar

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvSortDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSortDetail.CellBeginEdit
        Try
            'yang bisa diedit hanya masalah urutan dan jenis sort
            If e.ColumnIndex <> Me.ColSortNumber.Index And e.ColumnIndex <> Me.ColSrtType.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.dgvSortDetail.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSortDetail.CellEndEdit
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColSortNumber.Index Then
                For Each rowDGV In Me.dgvSortDetail.Rows
                    If rowDGV.Index <> e.RowIndex Then
                        If rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.dgvSortDetail.Item(e.ColumnIndex, e.RowIndex).Value Then
                            rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.Buffer
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortRekap_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSortRekap.CellBeginEdit
        Try
            'yang bisa diedit hanya masalah urutan dan jenis sort
            If e.ColumnIndex <> Me.ColSortNumber.Index And e.ColumnIndex <> Me.ColSrtType.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.dgvSortRekap.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortRekap_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSortRekap.CellEndEdit
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColSortNumber.Index Then
                For Each rowDGV In Me.dgvSortRekap.Rows
                    If rowDGV.Index <> e.RowIndex Then
                        If rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.dgvSortRekap.Item(e.ColumnIndex, e.RowIndex).Value Then
                            rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.Buffer
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortRekap2PeriodComp_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSortRekap2PeriodComp.CellBeginEdit
        Try
            'yang bisa diedit hanya masalah urutan dan jenis sort
            If e.ColumnIndex <> Me.ColSortNumber.Index And e.ColumnIndex <> Me.ColSrtType.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.dgvSortRekap2PeriodComp.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortRekap2PeriodComp_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSortRekap2PeriodComp.CellEndEdit
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColSortNumber.Index Then
                For Each rowDGV In Me.dgvSortRekap2PeriodComp.Rows
                    If rowDGV.Index <> e.RowIndex Then
                        If rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.dgvSortRekap2PeriodComp.Item(e.ColumnIndex, e.RowIndex).Value Then
                            rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.Buffer
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub dgvProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProduct.CellBeginEdit
        Try
            If Me.colProduct_GroupLevel.Index Or Me.colProduct_Code.Index Then
                Me.Buffer = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProduct.CellEndEdit
        Try
            If IsNothing(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If String.IsNullOrEmpty(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'If IsNothing(Me.Buffer) = True Then
                    '    Me.index = e.RowIndex
                    '    Me.BgwProduct.RunWorkerAsync()
                    '    Exit Sub
                    'End If
                    If Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.colProduct_GroupLevel.Index Then
                        cellEndEditGroupLevel(sender, e)
                    End If
                    If e.ColumnIndex = Me.colProduct_Code.Index Then
                        cellEndEditCode(sender, e)
                    End If
                Else
                    If IsNothing(Me.Buffer) = False Then
                        If String.IsNullOrEmpty(Buffer) = False Then
                            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub cellEndEditCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 3)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 3
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
                Case "Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 2
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
                Case "Sub Brand"
                    If Me.objGroup.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objGroup.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value, 1)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindGroup.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindGroup.ProductGroupLevel = 1
                        frmFindGroup.ShowDialog()
                        If frmFindGroup.PrimaryProductGroup > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
                Case "Category"
                    If Me.objCategory.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objCategory.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindCategory.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindCategory.ShowDialog()
                        If frmFindCategory.PrimaryCategory > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
                Case "Product"
                    If Me.objProduct.ApakahSudahAda(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = objProduct.GetPrimary(Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        frmFindProduct.txtCari.Text = Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        frmFindProduct.ShowDialog()
                        If frmFindProduct.PrimaryProduct > 0 Then
                            Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                            Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                            Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Else
                            If IsNothing(Me.Buffer) = False Then
                                If String.IsNullOrEmpty(Buffer) = False Then
                                    Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                            Me.index = e.RowIndex
                            Me.BgwProduct.RunWorkerAsync()
                            Exit Sub
                        End If
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            Else
                'If IsNothing(Me.Buffer) = False Then
                '    If String.IsNullOrEmpty(Buffer) = False Then
                '        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                '        Exit Sub
                '    End If
                'End If
                Me.index = e.RowIndex
                Me.BgwProduct.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function getIndexProduct(ByVal vRowException As Integer, ByVal vGroupLevel As String, ByVal vPrimary As Double) As Integer
        Try
            For Each dr As DataGridViewRow In Me.dgvProduct.Rows
                With dr
                    If Me.dgvProduct.Rows.IndexOf(dr) <> vRowException Then
                        If dr.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                            If dr.Cells(Me.colProduct_GroupLevel.Index).Value = vGroupLevel And dr.Cells(Me.colProduct_Primary.Index).Value = vPrimary Then
                                Return Me.dgvProduct.Rows.IndexOf(dr)
                            End If
                        End If
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub cellEndEditGroupLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frmFindGroup As New Core_Module.FrmCariMasterProductGroup
        Dim frmFindCategory As New Core_Module.frmCariMasterCategory
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            Select Case Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value
                Case "Group"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 3
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                Case "Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 2
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                Case "Sub Brand"
                    frmFindGroup.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindGroup.ProductGroupLevel = 1
                    frmFindGroup.ShowDialog()
                    If frmFindGroup.PrimaryProductGroup > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindGroup.PrimaryProductGroup
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objGroup.GetGroupCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objGroup.GetGroupDesc(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Me.index = e.RowIndex
                        Me.BgwProduct.RunWorkerAsync()
                        Exit Sub
                    End If
                Case "Category"
                    frmFindCategory.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindCategory.ShowDialog()
                    If frmFindCategory.PrimaryCategory > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindCategory.PrimaryCategory
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objCategory.GetCategoryCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objCategory.GetBrandDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Case "Product"
                    frmFindProduct.txtCari.Text = Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value
                    frmFindProduct.ShowDialog()
                    If frmFindProduct.PrimaryProduct > 0 Then
                        Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value = frmFindProduct.PrimaryProduct
                        Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value = objProduct.GetProductCode(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                        Me.dgvProduct.Item(Me.colProduct_Desc.Index, e.RowIndex).Value = objProduct.GetDescription(Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value)
                    Else
                        Me.dgvProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
            End Select
            If Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value <> -1 Then
                If Me.getIndexProduct(e.RowIndex, Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value, Me.dgvProduct.Item(Me.colProduct_Primary.Index, e.RowIndex).Value) <> -1 Then
                    MsgBox(Me.dgvProduct.Item(Me.colProduct_GroupLevel.Index, e.RowIndex).Value & " " & Me.dgvProduct.Item(Me.colProduct_Code.Index, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                    Me.index = e.RowIndex
                    Me.BgwProduct.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwProduct.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwProduct_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwProduct.RunWorkerCompleted
        Try
            Me.dgvProduct.Rows.RemoveAt(Me.index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProduct.DataError

    End Sub

    Private Sub dgvProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProduct.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
            dr.Cells(Me.colProduct_No.Index).Value = Me.dgvProduct.Rows.Count
            dr.Cells(Me.colProduct_GroupLevel.Index).Value = "Group"
            dr.Cells(Me.colProduct_Desc.Index).Value = ""
            dr.Cells(Me.colProduct_Primary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProduct.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProduct.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub DgvCustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColC_Code.Index Then
                Me.Buffer = Me.DgvCustomer(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomer.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomer.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjRef As New Core_Module.clsMasterCustomer
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColC_Code.Index Then
                        .Cells(Me.ColC_Code.Index).Value = .Cells(Me.ColC_Code.Index).Value.ToString.ToUpper
                        If ObjRef.ApakahSudahAda(.Cells(Me.ColC_Code.Index).Value) = True Then
                            primary = ObjRef.GetPrimary(.Cells(Me.ColC_Code.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_C(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColC_Primary.Index).Value = primary
                                .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(primary)
                                .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.bgwCustomer.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomer
                            ij.txtCari.Text = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomer <> -1 Then
                                index = Me.GetIndexItem_C(ij.PrimaryCustomer, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                    .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                    .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_C = e.RowIndex
                                            Me.bgwCustomer.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                        .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                        .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomer = -1 Then
                                If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.bgwCustomer.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                        Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColC_Code.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_C = e.RowIndex
                                Me.bgwCustomer.RunWorkerAsync()
                            Else
                                Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomer)
                            Exit Sub
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColC_Code.Index Then
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
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
    Private Function GetIndexItem_C(ByVal vPrimary As Double, _
      ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomer.Rows.Count - 1
                RowDGV = Me.DgvCustomer.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColC_Primary.Index).Value = vPrimary Then
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
    Private Sub DgvCustomer_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvCustomer.DataError

    End Sub

    Private Sub DgvCustomer_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvCustomer.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvCustomer.Rows(Me.DgvCustomer.Rows.Count - 1)
            dr.Cells(Me.ColC_No.Index).Value = Me.DgvCustomer.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomer.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomer.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvCustomer, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustomer.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustomer.RunWorkerCompleted
        Try
            Me.DgvCustomer.Rows.RemoveAt(EraseIndex_C)
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMarketSegmentation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                Me.Buffer = Me.DgvMarketSegmentation(e.ColumnIndex, e.RowIndex).Value
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_MS = e.RowIndex
                                            Me.bgwCustSegmen.RunWorkerAsync()
                                        Else
                                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()

                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_MS = e.RowIndex
                                Me.bgwCustSegmen.RunWorkerAsync()
                            Else
                                Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                            Exit Sub
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColMS_MarketSeg.Index Then
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                Me.Buffer = Me.DgvSalesArea(e.ColumnIndex, e.RowIndex).Value
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_SA = e.RowIndex
                                            Me.bgwSalesArea.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_SA = e.RowIndex
                                Me.bgwSalesArea.RunWorkerAsync()
                            Else
                                Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesArea)
                            Exit Sub
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSA_SalesArea.Index Then
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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

    Private Sub DgvCustomerType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomerType.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCT_CustType.Index Then
                Me.Buffer = Me.DgvCustomerType(e.ColumnIndex, e.RowIndex).Value
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_CT = e.RowIndex
                                            Me.bgwCustType.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_CT = e.RowIndex
                                Me.bgwCustType.RunWorkerAsync()
                            Else
                                Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomerType)
                            Exit Sub
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCT_CustType.Index Then
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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

    Private Sub dgvSalesGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomerGroup.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                Me.Buffer = Me.dgvCustomerGroup(e.ColumnIndex, e.RowIndex).Value
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_SG = e.RowIndex
                                            Me.bgwCustGroup.RunWorkerAsync()
                                        Else
                                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_SG = e.RowIndex
                                Me.bgwCustGroup.RunWorkerAsync()
                            Else
                                Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                            Exit Sub
                        Else
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSG_SalesGroup.Index Then
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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

    Private Sub cboitemDetail_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboitemDetail.SelectedIndexChanged
        Try
            UpdateDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub UpdateDGV()
        Dim rowDGV As DataGridViewRow
        Try
            If Me.cboFormat.SelectedIndex <> 2 Then
                If Me.dgvSortRekap.Rows.Count > 0 And Me.cboitemDetail.Text <> "" Then
                    rowDGV = Me.dgvSortRekap.Rows(0)
                    rowDGV.Cells(Me.ColName.Index).Value = Me.cboitemDetail.Text
                End If
                If Me.dgvSortRekap2PeriodComp.Rows.Count > 0 And Me.cboitemDetail.Text <> "" Then
                    rowDGV = Me.dgvSortRekap2PeriodComp.Rows(0)
                    rowDGV.Cells(Me.ColName1.Index).Value = Me.cboitemDetail.Text
                End If
            Else
                If Me.dgvSortDetail.Rows.Count > 0 And Me.cboitemDetail.Text <> "" Then
                    rowDGV = Me.dgvSortDetail.Rows(2)
                    rowDGV.Cells(Me.ColName2.Index).Value = Me.cboitemDetail.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateFilteringBaru()
        Dim FilterProduct As String = ""
        Dim FilterSalesArea As String = ""
        Dim FilterSegment As String = ""
        Dim FilterCustomer As String = ""
        Dim FilterCustomerType As String = ""
        Dim FilterCustomerGroup As String = ""
        Dim FilterGroup As String = ""
        Dim rowDGV As New DataGridViewRow
        Dim Filter As String = ""
        Dim row As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row
        Dim TBLCopyan As New DseLaporanPenjualan.tbltrans_salesinvoice_item_view0DataTable
        Dim rowCopyan As DseLaporanPenjualan.tbltrans_salesinvoice_item_view0Row
        Try
            If Me.DgvMarketSegmentation.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvMarketSegmentation.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                        If FilterSegment = "" Then
                            FilterSegment = "[Primary Customer Segmentasi] = " & rowDGV.Cells(Me.ColMS_Primary.Index).Value
                        Else
                            FilterSegment &= " Or [Primary Customer Segmentasi] = " & rowDGV.Cells(Me.ColMS_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvSalesArea.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvSalesArea.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                        If FilterSalesArea = "" Then
                            FilterSalesArea = "[Primary Sales Area] = " & rowDGV.Cells(Me.ColSA_Primary.Index).Value
                        Else
                            FilterSalesArea &= " Or [Primary Sales Area] = " & rowDGV.Cells(Me.ColSA_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomer.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvCustomer.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColC_Code.Index).Value) = False Then
                        If FilterCustomer = "" Then
                            FilterCustomer = "[Primary Customer] = " & rowDGV.Cells(Me.ColC_Primary.Index).Value
                        Else
                            FilterCustomer &= " Or [Primary Customer] = " & rowDGV.Cells(Me.ColC_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomerType.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvCustomerType.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCT_CustType.Index).Value) = False Then
                        If FilterCustomerType = "" Then
                            FilterCustomerType = "[Primary Customer Type] = " & rowDGV.Cells(Me.ColCT_Primary.Index).Value
                        Else
                            FilterCustomerType &= " Or [Primary Customer Type] = " & rowDGV.Cells(Me.ColCT_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.dgvCustomerGroup.Rows.Count > 0 Then
                For Each rowDGV In Me.dgvCustomerGroup.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSG_SalesGroup.Index).Value) = False Then
                        If FilterCustomerGroup = "" Then
                            FilterCustomerGroup = "[Primary Customer Type] = " & rowDGV.Cells(Me.ColSG_Primary.Index).Value
                        Else
                            FilterCustomerGroup &= " Or [Primary Customer Type] = " & rowDGV.Cells(Me.ColSG_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            For Each rowDGV In Me.dgvProduct.Rows
                If rowDGV.Cells(Me.colProduct_Primary.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.colProduct_GroupLevel.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Product Group] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        Else
                            FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Product Group] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.colProduct_Primary.Index).Value.ToString
                        End If
                    End If
                End If
            Next

            If FilterSegment <> "" Then
                Filter = FilterSegment
            End If
            If FilterSalesArea <> "" Then
                If Filter = "" Then
                    Filter = FilterSalesArea
                Else
                    Filter = "(" & Filter & ") AND (" & FilterSalesArea & ") "
                End If
            End If
            If FilterCustomer <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomer
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomer & ") "
                End If
            End If
            If FilterCustomerType <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomerType
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomerType & ") "
                End If
            End If

            If FilterCustomerGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomerGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomerGroup & ") "
                End If
            End If

            

            If FilterGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterGroup & ") "
                End If
            End If

            If Filter <> "" Then
                Me.TBL.DefaultView.RowFilter = Filter
                For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                    rowCopyan = Me.TBL.DefaultView(i).Row
                    TBLCopyan.ImportRow(rowCopyan)
                Next

                Me.TBL.Rows.Clear()
                For Each row In TBLCopyan.Rows
                    TBL.ImportRow(row)
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

    Private Sub DgvMarketSegmentation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvMarketSegmentation.Sorted
        Try
            ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class