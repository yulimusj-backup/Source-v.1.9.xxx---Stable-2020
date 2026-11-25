Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanPembelian
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim DS2 As New DseLaporanPembelian

    Public TBL As DseLaporanPembelian.tbltrans_aop_view2DataTable = DS2.tbltrans_aop_view2
    Dim DTA As New DseLaporanPembelianTableAdapters.tbltrans_aop_view2TableAdapter

    'aul
    Public TBLPO As DseLaporanPembelian.tbltrans_po_viewDataTable = DS2.tbltrans_po_view
    Dim DTAPO As New DseLaporanPembelianTableAdapters.tbltrans_po_viewTableAdapter
    Dim rowPO As DseLaporanPembelian.tbltrans_po_viewRow
    '---

    Dim TBLDetail As DseLaporanPembelian.DetailDataTable = DS2.Detail
    Dim TBLRekap As DseLaporanPembelian.RekapDataTable = DS2.Rekap
    Dim TBLRekapComparison As DseLaporanPembelian.Rekap_ComparisonDataTable = DS2.Rekap_Comparison

    'komponen yang berkaitan dengan clb customer
    Dim ObjSupp As New Core_Module.clsMasterSupplier
    Dim TblSupp As New Core_Module.dseMasterSupplier.tblmaster_supplier2DataTable
    Dim rowSupp As Core_Module.dseMasterSupplier.tblmaster_supplier2Row

    'kompinen yang berkaitan clb product
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim TblProd As New Core_Module.dseMasterProduct.tblmaster_product2DataTable
    Dim rowProd As Core_Module.dseMasterProduct.tblmaster_product2Row

    Dim TBLReturn As DseLaporanPembelian.DataTable2DataTable = Me.DS2.DataTable2
    Dim DTAReturn As New DseLaporanPembelianTableAdapters.DataTable2TableAdapter
    Dim rowReturn As DseLaporanPembelian.DataTable2Row

    Dim ListBrand As New List(Of String)
    Dim ListCust As New List(Of String)

    Dim ReportDetail As New CrpLaporanPenjualanDetail
    Dim ReportRekap As New CrpLaporanPenjualanRekap
    Dim ReportRekap_Comparison As New CrpLaporanPenjualanRekap_2PeriodCust
    Dim ObjCD As New ClsCekData

    Dim Buffer As String = ""
    Dim ObjLicense As New Core_Module.ClsLicense

    'ahdar
    Dim ObjSupplier As New clsMasterSupplier
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim EraseRowIndex As Integer = 0
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim DummyIndex As Integer = -1
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim DefaultCurrency As Decimal = ObjCurr.GetDecimal_DefaultCurrency
    Dim TblCur As New DseLaporanPembelian.tblcurrDataTable
    Dim RowCur As DseLaporanPembelian.tblcurrRow

    Private Enum EnumFormat As Integer
        Rekap
        Rekap_2_Period_Comparison
        Detail
    End Enum

    Private Enum EnumGroupBy As Integer
        CodeNo
        Category
        SubBrand
        Brand
        ProductGroup
        CustCode
        Tanggal
        Invoice
    End Enum

    Private Enum EnumDataBy As Integer
        AOP
        PO
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

            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 Or Me.cboGroupBy.SelectedIndex = 4 Then
                'group by code no
                Name1 = "Date"
                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name2 = "AOP Number"
                Else
                    Name2 = "PO Number"
                End If

                Name3 = "Supplier Code"
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                'group by customer code
                Name1 = "Date"
                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name2 = "AOP Number"
                Else
                    Name2 = "PO Number"
                End If
                If Me.CboItemDetail.SelectedIndex = 0 Then
                    Name3 = "Code No"
                ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                    Name3 = "Category"
                ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                    Name3 = "Sub Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                    Name3 = "Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                    Name3 = "Product Group"
                End If
            ElseIf Me.cboGroupBy.SelectedIndex = 6 Then
                'group by tanggal
                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name1 = "AOP Number"
                Else
                    Name1 = "PO Number"
                End If
                Name2 = "Supplier Code"

                If Me.CboItemDetail.SelectedIndex = 0 Then
                    Name3 = "Code No"
                ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                    Name3 = "Category"
                ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                    Name3 = "Sub Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                    Name3 = "Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                    Name3 = "Product Group"
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

            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 Or Me.cboGroupBy.SelectedIndex = 4 Then
                'group by code no
                Name1 = "Supplier Code"
            ElseIf Me.cboGroupBy.SelectedIndex = 5 Then
                'group by customer code
                If Me.CboItemDetail.SelectedIndex = 0 Then
                    Name1 = "Code No"
                ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                    Name1 = "Category"
                ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                    Name1 = "Sub Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                    Name1 = "Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                    Name1 = "Product Group"
                End If
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
                If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 Or Me.cboGroupBy.SelectedIndex = 4 Then
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

            If Me.cboGroupBy.SelectedIndex = 0 Or Me.cboGroupBy.SelectedIndex = 1 Or Me.cboGroupBy.SelectedIndex = 2 Or Me.cboGroupBy.SelectedIndex = 3 Or Me.cboGroupBy.SelectedIndex = 4 Then
                'group by code no
                Name1 = "Supplier Code"
            ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                'group by customer code
                If Me.CboItemDetail.SelectedIndex = 0 Then
                    Name1 = "Code No"
                ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                    Name1 = "Category"
                ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                    Name1 = "Sub Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                    Name1 = "Brand"
                ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                    Name1 = "Product Group"
                End If
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

            If Me.CboItemDetail.SelectedIndex = 0 Then
                Me.ChkProductDesc.Text = "Show Product Description"
            ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                Me.ChkProductDesc.Text = "Show Category Description"
            ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                Me.ChkProductDesc.Text = "Show Sub Brand Description"
            ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                Me.ChkProductDesc.Text = "Show Brand Description"
            ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                Me.ChkProductDesc.Text = "Show Product Group Description"
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
            'ObjSupp.IsiCheckListBox(Me.clbCustomer, TblSupp)
            'Me.chkCustomer.Checked = True
            'pastikan clb product diisikan
            'ObjProd.IsiListBoxProduct(Me.clbProduct, TblProd)
            'Me.chkProduct.Checked = True
            'Me.IsikanDGV()
            Me.Text = "Purchase List"
            'aul
            Me.CboDataFrom.SelectedIndex = 0
            Me.DgvSupplier.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSortDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSortRekap.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSortRekap2PeriodComp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            Me.ChkProductDesc.Checked = False
            If Me.cboFormat.SelectedIndex = 0 Then
                'isikan cbo group by
                Me.cboGroupBy.Items.Clear()
                Me.cboGroupBy.Items.Add("Code No")
                Me.cboGroupBy.Items.Add("Category")
                Me.cboGroupBy.Items.Add("Sub Brand")
                Me.cboGroupBy.Items.Add("Brand")
                Me.cboGroupBy.Items.Add("Product Group")
                Me.cboGroupBy.Items.Add("Supplier Code")
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
                Me.cboGroupBy.Items.Add("Supplier Code")
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
                Me.cboGroupBy.Items.Add("Supplier Code")
                Me.cboGroupBy.Items.Add("Date")
                Me.cboGroupBy.SelectedIndex = 0
                Me.dgvSortDetail.BringToFront()
                Me.dgvSortRekap.SendToBack()
                If Me.cboGroupBy.SelectedIndex = 5 Or Me.cboGroupBy.SelectedIndex = 6 Then
                    Me.ChkProductDesc.Visible = True
                Else
                    Me.ChkProductDesc.Visible = False
                End If
            End If

            If Me.cboFormat.SelectedIndex = 1 Then
                Me.dtpFrom2.Enabled = True
                Me.dtpTo2.Enabled = True
            Else
                Me.dtpFrom2.Enabled = False
                Me.dtpTo2.Enabled = False
            End If

            If Me.cboGroupBy.Text = "Supplier Code" Then
                Me.CboItemDetail.Enabled = True
                Me.CboItemDetail.SelectedIndex = 0
            Else
                Me.CboItemDetail.Enabled = False
                Me.CboItemDetail.SelectedIndex = -1
            End If
            Me.IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGroupBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedIndexChanged
        Try
            Me.ChkProductDesc.Checked = False
            If Me.cboGroupBy.SelectedIndex = 5 Or Me.cboGroupBy.SelectedIndex = 6 Then
                Me.ChkProductDesc.Visible = True
            Else
                Me.ChkProductDesc.Visible = False
            End If
            If Me.cboGroupBy.Text = "Supplier Code" Or Me.cboGroupBy.Text = "Date" Then
                Me.CboItemDetail.Enabled = True
                Me.CboItemDetail.SelectedIndex = 0
            Else
                Me.CboItemDetail.Enabled = False
                Me.CboItemDetail.SelectedIndex = -1
            End If

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
            'If Me.clbProduct.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 product.", MsgBoxStyle.Information, "Information")
            '    Me.clbProduct.Focus()
            '    Return False
            'End If
            'If Me.clbCustomer.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 customer.", MsgBoxStyle.Information, "Information")
            '    Me.clbCustomer.Focus()
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanPembelian.tbltrans_aop_view2Row

        Dim i As Integer = 0

        Try
            For Each rowReturn In Me.TBLReturn.Rows
                Filter = "[Primary AOP] =" + rowReturn.Primary_AOP.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary PO] =" + rowReturn.Primary_PO.ToString + _
                " AND [Primary GRN] =" + rowReturn.Primary_GRN.ToString
                Me.TBL.DefaultView.RowFilter = Filter

                i = 0
                While Me.TBL.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = TryCast(Me.TBL.DefaultView(i).Row, DseLaporanPembelian.tbltrans_aop_view2Row)
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

    Public Sub CekMonthly(ByVal StartDate As Date, ByVal EndDate As Date)
        Dim rowAOP As DseLaporanPembelian.tbltrans_aop_view2Row
        Try
            Me.TBL.Rows.Clear()
            Me.TBLDetail.Rows.Clear()
            Me.TBLRekap.Rows.Clear()
            Me.TBLRekapComparison.Rows.Clear()
            Me.DS2.Rekap1.Rows.Clear()
            Me.DS2.Detail1.Rows.Clear()

            'aul
            Me.TBLPO.Rows.Clear()
            Dim dataBy As EnumDataBy = Me.CboDataFrom.SelectedIndex
            If dataBy = EnumDataBy.AOP Then
                DTA.FillBy_TglSampaiDari(TBL, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            ElseIf dataBy = EnumDataBy.PO Then
                DTAPO.FillBy_TglSampaiDari(TBLPO, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                For Each rowPO In TBLPO
                    With rowPO
                        rowAOP = TBL.Newtbltrans_aop_view2Row
                        rowAOP.Primary_AOP = .Primary_Transaksi
                        rowAOP.AOP_Date = .Tanggal_Transaksi
                        rowAOP.Primary_GRN = -1
                        rowAOP.Primary_PO = .Primary_Transaksi
                        rowAOP.Supplier_Code = .Supplier_Code
                        rowAOP.Supplier_Name = .Supplier_Name
                        rowAOP.Primary_Product = .Primary_Product
                        rowAOP.Product_Code = .Product_Code
                        rowAOP.Description = .Description
                        rowAOP._Decimal = ._Decimal
                        rowAOP.Total_Qty = .Qty
                        rowAOP.Amount = .Amount
                        rowAOP.No_Transaksi = .No_Transaksi
                        rowAOP.Unit_Code = .Unit_Code
                        TBL.Rows.Add(rowAOP)
                    End With
                Next
            End If
            '---
            Me.DTAReturn.Fill(Me.TBLReturn, EndDate, StartDate)
            Me.KurangkanItemReturn()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim rowView As DataRowView
        Dim rowDGV As New DataGridViewRow
        Dim rowAOP As DseLaporanPembelian.tbltrans_aop_view2Row
        Dim EndLicenseDate As Date = #1/1/1950#
        Dim ListPrimaryCur As New List(Of Double)
        Dim FrmOptionRate As New Core_Module.FrmOptionRate
        Dim Retval As New DseLaporanPembelian.tbltrans_po_viewDataTable
        Dim R As DseLaporanPembelian.tbltrans_po_viewRow

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

                'aul
                Me.TBLPO.Rows.Clear()
                Dim dataBy As EnumDataBy = Me.CboDataFrom.SelectedIndex
                If dataBy = EnumDataBy.AOP Then
                    DTA.FillBy_TglSampaiDari(TBL, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                ElseIf dataBy = EnumDataBy.PO Then
                    DTAPO.FillBy_TglSampaiDari(TBLPO, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                    For Each rowPO In TBLPO
                        With rowPO
                            If ListPrimaryCur.Contains(.Primary_Currency) = False And .Primary_Currency <> 1 Then
                                ListPrimaryCur.Add(.Primary_Currency)
                            End If
                        End With
                    Next

                    If Me.cboFormat.SelectedIndex = 1 Then
                        'tambah list jika format compare
                        DTAPO.FillBy_TglSampaiDari(Retval, Me.dtpTo2.Value.Date, Me.dtpFrom2.Value.Date)
                        For Each R In Retval
                            With R
                                If ListPrimaryCur.Contains(.Primary_Currency) = False And .Primary_Currency <> 1 Then
                                    ListPrimaryCur.Add(.Primary_Currency)
                                End If
                            End With
                        Next
                    End If

                    Me.TblCur.Rows.Clear()
                    If ListPrimaryCur.Count > 0 Then
                        FrmOptionRate.ListPrimaryCurr = ListPrimaryCur
                        If FrmOptionRate.ShowDialog = Windows.Forms.DialogResult.OK Then
                            For Each rowDGV In FrmOptionRate.DgvRate.Rows
                                If String.IsNullOrEmpty(rowDGV.Cells(FrmOptionRate.ColCurrency.Index).Value) = False Then
                                    TBLPO.DefaultView.RowFilter = "[Primary Currency] =" & rowDGV.Cells(FrmOptionRate.ColPrimaryCurr.Index).Value
                                    For i As Integer = 0 To TBLPO.DefaultView.Count - 1
                                        rowPO = TBLPO.DefaultView(i).Row
                                        rowPO.Amount = rowPO.Amount * CDec(rowDGV.Cells(FrmOptionRate.ColRate.Index).Value)
                                    Next

                                    'jika format compare add to tbl
                                    If Me.cboFormat.SelectedIndex = 1 Then
                                        RowCur = TblCur.NewtblcurrRow
                                        With RowCur
                                            .primary_currency = rowDGV.Cells(FrmOptionRate.ColPrimaryCurr.Index).Value
                                            .Currency_Code = rowDGV.Cells(FrmOptionRate.ColCurrency.Index).Value
                                            .rate = rowDGV.Cells(FrmOptionRate.ColRate.Index).Value
                                            .Default_Precision = 0 'kolom tidak dipake jadi set 0 saja
                                        End With
                                        TblCur.AddtblcurrRow(RowCur)
                                    End If
                                End If
                            Next
                        End If
                    End If

                    For Each rowPO In TBLPO
                        With rowPO
                            rowAOP = TBL.Newtbltrans_aop_view2Row
                            rowAOP.Primary_AOP = .Primary_Transaksi
                            rowAOP.AOP_Date = .Tanggal_Transaksi
                            rowAOP.Primary_GRN = -1
                            rowAOP.Primary_PO = .Primary_Transaksi
                            rowAOP.Supplier_Code = .Supplier_Code
                            rowAOP.Supplier_Name = .Supplier_Name
                            rowAOP.Primary_Product = .Primary_Product
                            rowAOP.Product_Code = .Product_Code
                            rowAOP.Description = .Description
                            rowAOP._Decimal = ._Decimal
                            rowAOP.Total_Qty = .Qty
                            rowAOP.Amount = .Amount
                            rowAOP.No_Transaksi = .No_Transaksi
                            rowAOP.Unit_Code = .Unit_Code
                            rowAOP.Primary_Supplier = .Primary_Supplier
                            rowAOP.Primary_Category = .Primary_Category
                            rowAOP.Category_Code = .Category_Code
                            rowAOP.Category_Desc = .Category_Desc
                            rowAOP.Primary_Sub_Brand = .Primary_Sub_Brand
                            rowAOP.Sub_Brand_Code = .Sub_Brand_Code
                            rowAOP.Sub_Brand_Desc = .Sub_Brand_Desc
                            rowAOP.Primary_Brand = .Primary_Brand
                            rowAOP.Brand_Code = .Brand_Code
                            rowAOP.Brand_Desc = .Brand_Desc
                            rowAOP.Primary_Product_Group = .Primary_Product_Group
                            rowAOP.Group_Code = .Group_Code
                            rowAOP.Group_Desc = .Group_Desc
                            rowAOP.Primary_Supplier = .Primary_Supplier
                            rowAOP.Supplier_Code = .Supplier_Code
                            rowAOP.Supplier_Name = .Supplier_Name
                            TBL.Rows.Add(rowAOP)
                        End With
                    Next
                End If
                '---

                Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpFrom.Value, Me.dtpTo.Value)
                Me.KurangkanItemReturn()

                'set kriteria filter
                'Me.ListBrand.Clear()
                'Me.ListCust.Clear()

                'If Me.dgvItemProductGroup.Rows.Count > 0 Then
                '    For Each rowDGV In Me.dgvItemProductGroup.Rows
                '        If String.IsNullOrEmpty(rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value) = False Then
                '            Me.ListBrand.Add(rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value)
                '        End If
                '    Next
                'End If

                'If Me.DgvSupplier.Rows.Count > 0 Then
                '    For Each rowDGV In Me.DgvSupplier.Rows
                '        If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSupp_Code.Index).Value) = False Then
                '            Me.ListCust.Add(rowDGV.Cells(Me.ColSupp_Primary.Index).Value)
                '        End If
                '    Next
                'End If

                'filter
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.FilterByItemProductAndSupplier()
                If Me.cboFormat.SelectedIndex = 2 Then
                    If TBL.Rows.Count = 0 Then
                        MsgBox("No purchase data on the database.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.GenerateDetail()
                ElseIf Me.cboFormat.SelectedIndex = 0 Then
                    If TBL.Rows.Count = 0 Then
                        MsgBox("No purchase data on the database.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.GenerateRekap()
                Else
                    Me.GenerateRekap2PeriodComparison()
                End If
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Function FilterByItemProductAndSupplier() As Boolean
        Dim FilterProduct As String = ""
        Dim FilterSupplier As String = ""
        Dim FilterGroup As String = ""
        Dim rowDGV As New DataGridViewRow
        Dim Filter As String = ""
        Dim rowAOP As DseLaporanPembelian.tbltrans_aop_view2Row
        Dim TBLCopyan As New DseLaporanPembelian.tbltrans_aop_view2DataTable
        Dim rowAOPCopyan As DseLaporanPembelian.tbltrans_aop_view2Row
        Try
            If Me.DgvSupplier.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvSupplier.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSupp_Code.Index).Value) = False Then
                        If FilterSupplier = "" Then
                            FilterSupplier = "[Primary Supplier] = " & rowDGV.Cells(Me.ColSupp_Primary.Index).Value
                        Else
                            FilterSupplier &= " Or [Primary Supplier] = " & rowDGV.Cells(Me.ColSupp_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If FilterSupplier <> "" Then
                Filter = FilterSupplier
            End If

            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Product Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    End If
                End If
            Next

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
                    rowAOP = Me.TBL.DefaultView(i).Row
                    TBLCopyan.ImportRow(rowAOP)
                Next

                Me.TBL.Rows.Clear()
                For Each rowAOPCopyan In TBLCopyan.Rows
                    TBL.ImportRow(rowAOPCopyan)
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

#Region "Coding Lama GenerateRekap2PeriodComparison"
    'Private Sub GenerateRekap2PeriodComparison()
    '    Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
    '    Dim rowView As DataRowView
    '    Dim TBLDummy As New DseLaporanPembelian.Rekap_Comparison1DataTable
    '    Dim rowDummy As DseLaporanPembelian.Rekap_Comparison1Row
    '    Dim rowRekap_Comparison As DseLaporanPembelian.Rekap_ComparisonRow
    '    Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
    '    Dim Name1 As EnumGroupBy
    '    Dim IndexMinimum As Integer = 0
    '    Dim Name1_Str As String
    '    Dim Lanjut As Boolean = False

    '    Try
    '        For Each row In Me.TBL.Rows
    '            Lanjut = True

    '            If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
    '                Lanjut = True
    '            Else
    '                If Me.ApakahMasukKriteria(row) <> True Then
    '                    Lanjut = False
    '                End If
    '            End If
    '            If Lanjut = True Then
    '                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
    '                "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "'"

    '                If TBLDummy.DefaultView.Count = 0 Then
    '                    rowDummy = TBLDummy.NewRekap_Comparison1Row
    '                    rowDummy.Amount1 = row.Amount
    '                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
    '                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
    '                    rowDummy.Qty1 = row.Total_Qty
    '                    rowDummy._Decimal = row._Decimal
    '                    rowDummy.Unit = row.Unit_Code
    '                    rowDummy.Product_Desc = row.Description
    '                    rowDummy.Amount2 = 0
    '                    rowDummy.Qty2 = 0
    '                    rowDummy.Persen_Amount = 0
    '                    rowDummy.Persen_Sub_Total = 0
    '                    rowDummy.Persen_Total = 0
    '                    TBLDummy.Rows.Add(rowDummy)
    '                Else
    '                    rowDummy = TBLDummy.DefaultView(0).Row
    '                    rowDummy.Amount1 += row.Amount
    '                    rowDummy.Qty1 += row.Total_Qty
    '                End If
    '            End If
    '        Next

    '        'isikan data periode 2
    '        Me.TBL.Rows.Clear()
    '        Me.TBLReturn.Rows.Clear()

    '        'aul
    '        Me.TBLPO.Rows.Clear()
    '        Dim dataBy As EnumDataBy = Me.CboDataFrom.SelectedIndex
    '        If dataBy = EnumDataBy.AOP Then
    '            DTA.FillBy_TglSampaiDari(TBL, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        ElseIf dataBy = EnumDataBy.PO Then
    '            DTAPO.FillBy_TglSampaiDari(TBLPO, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '            For Each rowPO In TBLPO
    '                With rowPO
    '                    row = TBL.Newtbltrans_aop_view2Row
    '                    row.Primary_AOP = .Primary_Transaksi
    '                    row.AOP_Date = .Tanggal_Transaksi
    '                    row.Primary_GRN = -1
    '                    row.Primary_PO = .Primary_Transaksi
    '                    row.Supplier_Code = .Supplier_Code
    '                    row.Supplier_Name = .Supplier_Name
    '                    row.Primary_Product = .Primary_Product
    '                    row.Product_Code = .Product_Code
    '                    row.Description = .Description
    '                    row._Decimal = ._Decimal
    '                    row.Total_Qty = .Qty
    '                    row.Amount = .Amount
    '                    row.No_Transaksi = .No_Transaksi
    '                    row.Unit_Code = .Unit_Code
    '                    TBL.Rows.Add(row)
    '                End With
    '            Next
    '        End If
    '        '---
    '        Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpFrom.Value, Me.dtpTo.Value)
    '        Me.KurangkanItemReturn()

    '        For Each rowView In Me.TBL.DefaultView
    '            Lanjut = True
    '            row = rowView.Row
    '            If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
    '                Lanjut = True
    '            Else
    '                If Me.ApakahMasukKriteria(row) <> True Then
    '                    Lanjut = False
    '                End If
    '            End If
    '            If Lanjut = True Then
    '                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
    '                "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "'"

    '                If TBLDummy.DefaultView.Count = 0 Then
    '                    rowDummy = TBLDummy.NewRekap_Comparison1Row
    '                    rowDummy.Amount2 = row.Amount
    '                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
    '                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
    '                    rowDummy.Qty2 = row.Total_Qty
    '                    rowDummy._Decimal = row._Decimal
    '                    rowDummy.Unit = row.Unit_Code
    '                    rowDummy.Product_Desc = row.Description
    '                    rowDummy.Amount1 = 0
    '                    rowDummy.Qty1 = 0
    '                    rowDummy.Persen_Amount = 0
    '                    rowDummy.Persen_Sub_Total = 0
    '                    rowDummy.Persen_Total = 0
    '                    TBLDummy.Rows.Add(rowDummy)
    '                Else
    '                    rowDummy = TBLDummy.DefaultView(0).Row
    '                    rowDummy.Amount2 += row.Amount
    '                    rowDummy.Qty2 += row.Total_Qty
    '                End If
    '            End If
    '        Next

    '        If TBLDummy.Rows.Count = 0 Then
    '            MsgBox("No purchase data on the database.", MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        'hitung total group
    '        Me.HitungSubTotalQtyAmount(TBLDummy)

    '        'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
    '        TBLDummy.DefaultView.RowFilter = ""
    '        TBLDummy.DefaultView.Sort = Me.GenerateSortRekap_Comparison
    '        For Each rowView In TBLDummy.DefaultView
    '            rowDummy = rowView.Row
    '            rowRekap_Comparison = Me.TBLRekapComparison.NewRekap_ComparisonRow
    '            If GroupBy = EnumGroupBy.CodeNo Then
    '                rowRekap_Comparison.Group_By = rowDummy.Product_Code
    '            ElseIf GroupBy = EnumGroupBy.CustCode Then
    '                rowRekap_Comparison.Group_By = rowDummy.Cust_Code
    '            End If

    '            'name1
    '            Name1 = Me.GetName1_Rekap
    '            If Name1 = EnumGroupBy.CodeNo Then
    '                rowRekap_Comparison.Name_1 = rowDummy.Product_Code
    '            ElseIf Name1 = EnumGroupBy.CustCode Then
    '                rowRekap_Comparison.Name_1 = rowDummy.Cust_Code
    '            End If


    '            rowRekap_Comparison.Qty1 = rowDummy.Qty1
    '            rowRekap_Comparison.Amount1 = rowDummy.Amount1
    '            rowRekap_Comparison.Qty2 = rowDummy.Qty2
    '            rowRekap_Comparison.Amount2 = rowDummy.Amount2
    '            rowRekap_Comparison.Persen_Amount = rowDummy.Persen_Amount
    '            rowRekap_Comparison.Persen_Sub_Total = rowDummy.Persen_Sub_Total
    '            rowRekap_Comparison.Persen_Total = rowDummy.Persen_Total

    '            rowRekap_Comparison._Decimal = rowDummy._Decimal
    '            rowRekap_Comparison.Unit = rowDummy.Unit
    '            Me.TBLRekapComparison.Rows.Add(rowRekap_Comparison)
    '        Next

    '        Name1_Str = ""
    '        If Name1 = EnumGroupBy.CodeNo Then
    '            Name1_Str = "Code No"
    '        ElseIf Name1 = EnumGroupBy.CustCode Then
    '            Name1_Str = "Supplier Code"
    '        End If


    '        ReportRekap_Comparison.SetDataSource(Me.DS2)
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Purchase List")
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate1.ParameterFieldName, Me.dtpFrom2.Value)
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate1.ParameterFieldName, Me.dtpTo2.Value)
    '        ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Name1.ParameterFieldName, Name1_Str)

    '        If GroupBy = EnumGroupBy.CodeNo Then
    '            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Code No")
    '        ElseIf GroupBy = EnumGroupBy.CustCode Then
    '            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Supplier")
    '        ElseIf GroupBy = EnumGroupBy.Tanggal Then
    '            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Date")
    '        End If

    '        Me.crv1.ReportSource = ReportRekap_Comparison
    '        Me.crv1.DisplayGroupTree = False
    '        Me.crv1.DisplayStatusBar = True
    '        Me.crv1.DisplayToolbar = True
    '        Me.crv1.ShowGroupTreeButton = False
    '        Me.crv1.ShowExportButton = True
    '        Me.crv1.ShowRefreshButton = False
    '        Me.crv1.Show()

    '        Me.TabControl1.SelectedIndex = 1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
#End Region

#Region "Coding Baru"
    Private Sub GenerateRekap2PeriodComparison()
        Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPembelian.Rekap_Comparison1DataTable
        Dim rowDummy As DseLaporanPembelian.Rekap_Comparison1Row
        Dim rowRekap_Comparison As DseLaporanPembelian.Rekap_ComparisonRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str As String
        Dim Lanjut As Boolean = False

        Try
            For Each row In Me.TBL.Rows
                TBLDummy.DefaultView.RowFilter = Me.GetFilterGroupBy(GroupBy, row)
                If TBLDummy.DefaultView.Count = 0 Then
                    rowDummy = TBLDummy.NewRekap_Comparison1Row
                    rowDummy.Amount1 = row.Amount
                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                    rowDummy.Category_Code = row.Category_Code + " - " + row.Category_Desc
                    rowDummy.Category_Desc = row.Category_Desc
                    rowDummy.Sub_Brand_Code = row.Sub_Brand_Code + " - " + row.Sub_Brand_Desc
                    rowDummy.Sub_Brand_Desc = row.Sub_Brand_Desc
                    rowDummy.Brand_Code = row.Brand_Code + " - " + row.Brand_Desc
                    rowDummy.Brand_Desc = row.Brand_Desc
                    rowDummy.Group_Code = row.Group_Code + " - " + row.Group_Desc
                    rowDummy.Group_Desc = row.Group_Desc
                    rowDummy.Qty1 = row.Total_Qty
                    rowDummy._Decimal = row._Decimal
                    rowDummy.Unit = row.Unit_Code
                    rowDummy.Product_Desc = row.Description
                    rowDummy.Amount2 = 0
                    rowDummy.Qty2 = 0
                    rowDummy.Persen_Amount = 0
                    rowDummy.Persen_Sub_Total = 0
                    rowDummy.Persen_Total = 0
                    TBLDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = TBLDummy.DefaultView(0).Row
                    rowDummy.Amount1 += row.Amount
                    rowDummy.Qty1 += row.Total_Qty
                End If
            Next

            'isikan data periode 2
            Me.TBL.Rows.Clear()
            Me.TBLReturn.Rows.Clear()

            Me.TBLPO.Rows.Clear()
            Dim dataBy As EnumDataBy = Me.CboDataFrom.SelectedIndex
            If dataBy = EnumDataBy.AOP Then
                DTA.FillBy_TglSampaiDari(TBL, Me.dtpTo2.Value.Date, Me.dtpFrom2.Value.Date)
            ElseIf dataBy = EnumDataBy.PO Then
                DTAPO.FillBy_TglSampaiDari(TBLPO, Me.dtpTo2.Value.Date, Me.dtpFrom2.Value.Date)
                'convert to idr if non idr
                For Each RowCur In Me.TblCur.Rows
                    TBLPO.DefaultView.RowFilter = "[Primary Currency] =" & RowCur.primary_currency
                    For i As Integer = 0 To TBLPO.DefaultView.Count - 1
                        rowPO = TBLPO.DefaultView(i).Row
                        rowPO.Amount = rowPO.Amount * RowCur.rate
                    Next
                Next
                'finish convert

                For Each rowPO In TBLPO
                    With rowPO
                        row = TBL.Newtbltrans_aop_view2Row
                        row.Primary_AOP = .Primary_Transaksi
                        row.AOP_Date = .Tanggal_Transaksi
                        row.Primary_GRN = -1
                        row.Primary_PO = .Primary_Transaksi
                        row.Supplier_Code = .Supplier_Code
                        row.Supplier_Name = .Supplier_Name
                        row.Primary_Product = .Primary_Product
                        row.Product_Code = .Product_Code
                        row.Primary_Category = .Primary_Category
                        row.Category_Code = .Category_Code
                        row.Category_Desc = .Category_Desc
                        row.Primary_Sub_Brand = .Primary_Sub_Brand
                        row.Sub_Brand_Code = .Sub_Brand_Code
                        row.Sub_Brand_Desc = .Sub_Brand_Desc
                        row.Primary_Brand = .Primary_Sub_Brand
                        row.Brand_Code = .Brand_Code
                        row.Brand_Desc = .Brand_Desc
                        row.Primary_Product_Group = .Primary_Product_Group
                        row.Group_Code = .Group_Code
                        row.Group_Desc = .Group_Desc
                        row.Primary_Supplier = .Primary_Supplier
                        row.Supplier_Code = .Supplier_Code
                        row.Supplier_Name = .Supplier_Name
                        row.Description = .Description
                        row._Decimal = ._Decimal
                        row.Total_Qty = .Qty
                        row.Amount = .Amount
                        row.No_Transaksi = .No_Transaksi
                        row.Unit_Code = .Unit_Code
                        TBL.Rows.Add(row)
                    End With
                Next
            End If
            '---
            Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpFrom2.Value, Me.dtpTo2.Value)
            Me.KurangkanItemReturn()
            Me.FilterByItemProductAndSupplier()
            For Each rowView In Me.TBL.DefaultView
                row = rowView.Row
                TBLDummy.DefaultView.RowFilter = Me.GetFilterGroupBy(GroupBy, row)
                If TBLDummy.DefaultView.Count = 0 Then
                    rowDummy = TBLDummy.NewRekap_Comparison1Row
                    rowDummy.Amount2 = row.Amount
                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                    rowDummy.Category_Code = row.Category_Code + " - " + row.Category_Desc
                    rowDummy.Category_Desc = row.Category_Desc
                    rowDummy.Sub_Brand_Code = row.Sub_Brand_Code + " - " + row.Sub_Brand_Desc
                    rowDummy.Sub_Brand_Desc = row.Sub_Brand_Desc
                    rowDummy.Brand_Code = row.Brand_Code + " - " + row.Brand_Desc
                    rowDummy.Brand_Desc = row.Brand_Desc
                    rowDummy.Group_Code = row.Group_Code + " - " + row.Group_Desc
                    rowDummy.Group_Desc = row.Group_Desc
                    rowDummy.Qty2 = row.Total_Qty
                    rowDummy._Decimal = row._Decimal
                    rowDummy.Unit = row.Unit_Code
                    rowDummy.Product_Desc = row.Description
                    rowDummy.Amount1 = 0
                    rowDummy.Qty1 = 0
                    rowDummy.Persen_Amount = 0
                    rowDummy.Persen_Sub_Total = 0
                    rowDummy.Persen_Total = 0
                    TBLDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = TBLDummy.DefaultView(0).Row
                    rowDummy.Amount2 += row.Amount
                    rowDummy.Qty2 += row.Total_Qty
                End If
            Next

            If TBLDummy.Rows.Count = 0 Then
                MsgBox("No purchase data on the database.", MsgBoxStyle.Information, "Information")
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
                ElseIf GroupBy = EnumGroupBy.Category Then
                    rowRekap_Comparison.Group_By = rowDummy.Category_Code
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    rowRekap_Comparison.Group_By = rowDummy.Sub_Brand_Code
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    rowRekap_Comparison.Group_By = rowDummy.Brand_Code
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    rowRekap_Comparison.Group_By = rowDummy.Group_Code
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    rowRekap_Comparison.Group_By = rowDummy.Cust_Code
                End If

                'name1
                Name1 = Me.GetName1_Rekap
                If Name1 = EnumGroupBy.CodeNo Then
                    rowRekap_Comparison.Name_1 = rowDummy.Product_Code
                ElseIf Name1 = EnumGroupBy.Category Then
                    rowRekap_Comparison.Name_1 = rowDummy.Category_Code
                ElseIf Name1 = EnumGroupBy.SubBrand Then
                    rowRekap_Comparison.Name_1 = rowDummy.Sub_Brand_Code
                ElseIf Name1 = EnumGroupBy.Brand Then
                    rowRekap_Comparison.Name_1 = rowDummy.Brand_Code
                ElseIf Name1 = EnumGroupBy.ProductGroup Then
                    rowRekap_Comparison.Name_1 = rowDummy.Group_Code
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
                Name1_Str = "Code No"
            ElseIf Name1 = EnumGroupBy.Category Then
                Name1_Str = "Category"
            ElseIf Name1 = EnumGroupBy.SubBrand Then
                Name1_Str = "Sub Brand"
            ElseIf Name1 = EnumGroupBy.Brand Then
                Name1_Str = "Brand"
            ElseIf Name1 = EnumGroupBy.ProductGroup Then
                Name1_Str = "Product Group"
            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Supplier Code"
            End If

            ReportRekap_Comparison.SetDataSource(Me.DS2)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Default_Precision.ParameterFieldName, Me.DefaultCurrency)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_ReportName.ParameterFieldName, "Purchase List")
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_StartDate1.ParameterFieldName, Me.dtpFrom2.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_EndDate1.ParameterFieldName, Me.dtpTo2.Value)
            ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Name1.ParameterFieldName, Name1_Str)

            If GroupBy = EnumGroupBy.CodeNo Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Code No")
            ElseIf GroupBy = EnumGroupBy.Category Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Category")
            ElseIf GroupBy = EnumGroupBy.SubBrand Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Sub Brand")
            ElseIf GroupBy = EnumGroupBy.Brand Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Brand")
            ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Product Group")
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Supplier")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportRekap_Comparison.SetParameterValue(ReportRekap_Comparison.Parameter_Group_By.ParameterFieldName, "Date")
            End If

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
#End Region

#Region "Coding Lama GenerateRekap"
    'Private Sub GenerateRekap()
    '    Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
    '    Dim rowView As DataRowView
    '    Dim TBLDummy As New DseLaporanPembelian.Rekap1DataTable
    '    Dim rowDummy As DseLaporanPembelian.Rekap1Row
    '    Dim rowRekap As DseLaporanPembelian.RekapRow
    '    Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
    '    Dim Name1 As EnumGroupBy
    '    Dim IndexMinimum As Integer = 0
    '    Dim Name1_Str As String
    '    Dim Lanjut As Boolean = False

    '    Try
    '        For Each row In Me.TBL.Rows
    '            Lanjut = True

    '            If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
    '                Lanjut = True
    '            Else
    '                If Me.ApakahMasukKriteria(row) <> True Then
    '                    Lanjut = False
    '                End If
    '            End If
    '            If Lanjut = True Then
    '                TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
    '                "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "'"

    '                If TBLDummy.DefaultView.Count = 0 Then
    '                    rowDummy = TBLDummy.NewRekap1Row
    '                    rowDummy.Amount = row.Amount
    '                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
    '                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
    '                    rowDummy.Qty = row.Total_Qty
    '                    rowDummy._Decimal = row._Decimal
    '                    rowDummy.Unit = row.Unit_Code
    '                    rowDummy.Product_Desc = row.Description
    '                    TBLDummy.Rows.Add(rowDummy)
    '                Else
    '                    rowDummy = TBLDummy.DefaultView(0).Row
    '                    rowDummy.Amount += row.Amount
    '                    rowDummy.Qty += row.Total_Qty
    '                End If
    '            End If
    '        Next

    '        'hitung total group
    '        If Me.cboSort1.SelectedIndex <> 0 And Me.cboSort1.SelectedIndex <> 1 Then
    '            Me.HitungSubTotalQtyAmount(TBLDummy)
    '        End If

    '        'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
    '        TBLDummy.DefaultView.RowFilter = ""
    '        TBLDummy.DefaultView.Sort = Me.GenerateSortRekap
    '        For Each rowView In TBLDummy.DefaultView
    '            rowDummy = rowView.Row
    '            rowRekap = Me.TBLRekap.NewRekapRow
    '            If GroupBy = EnumGroupBy.CodeNo Then
    '                rowRekap.Group_By = rowDummy.Product_Code
    '            ElseIf GroupBy = EnumGroupBy.CustCode Then
    '                rowRekap.Group_By = rowDummy.Cust_Code
    '            End If

    '            'name1
    '            Name1 = Me.GetName1_Rekap
    '            If Name1 = EnumGroupBy.CodeNo Then
    '                rowRekap.Name_1 = rowDummy.Product_Code
    '            ElseIf Name1 = EnumGroupBy.CustCode Then
    '                rowRekap.Name_1 = rowDummy.Cust_Code
    '            End If


    '            rowRekap.Qty = rowDummy.Qty
    '            rowRekap.Amount = rowDummy.Amount
    '            rowRekap._Decimal = rowDummy._Decimal
    '            rowRekap.Unit = rowDummy.Unit
    '            Me.TBLRekap.Rows.Add(rowRekap)
    '        Next

    '        Name1_Str = ""
    '        If Name1 = EnumGroupBy.CodeNo Then
    '            Name1_Str = "Code No"
    '        ElseIf Name1 = EnumGroupBy.CustCode Then
    '            Name1_Str = "Supplier Code"
    '        End If


    '        ReportRekap.SetDataSource(Me.DS2)
    '        ReportRekap.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Purchase List")
    '        ReportRekap.SetParameterValue(ReportRekap.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
    '        ReportRekap.SetParameterValue(ReportRekap.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
    '        ReportRekap.SetParameterValue(ReportRekap.Parameter_Name1.ParameterFieldName, Name1_Str)
    '        If GroupBy = EnumGroupBy.CodeNo Then
    '            ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Code No")
    '        ElseIf GroupBy = EnumGroupBy.CustCode Then
    '            ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Supplier")
    '        ElseIf GroupBy = EnumGroupBy.Tanggal Then
    '            ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Date")
    '        End If

    '        Me.crv1.ReportSource = ReportRekap
    '        Me.crv1.DisplayGroupTree = False
    '        Me.crv1.DisplayStatusBar = True
    '        Me.crv1.DisplayToolbar = True
    '        Me.crv1.ShowGroupTreeButton = False
    '        Me.crv1.ShowExportButton = True
    '        Me.crv1.ShowRefreshButton = False
    '        Me.crv1.Show()

    '        Me.TabControl1.SelectedIndex = 1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
#End Region

#Region "Coding Baru"
    Private Sub GenerateRekap()
        Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPembelian.Rekap1DataTable
        Dim rowDummy As DseLaporanPembelian.Rekap1Row
        Dim rowRekap As DseLaporanPembelian.RekapRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str As String
        Dim Lanjut As Boolean = False

        Try
            For Each row In Me.TBL.Rows
                TBLDummy.DefaultView.RowFilter = Me.GetFilterGroupBy(GroupBy, row)
                If TBLDummy.DefaultView.Count = 0 Then
                    rowDummy = TBLDummy.NewRekap1Row
                    rowDummy.Amount = row.Amount
                    rowDummy.Cust_Code = row.Supplier_Code + " - " + row.Supplier_Name
                    rowDummy.Product_Code = row.Product_Code + " - " + row.Description
                    rowDummy.Qty = row.Total_Qty
                    rowDummy._Decimal = row._Decimal
                    rowDummy.Unit = row.Unit_Code
                    rowDummy.Product_Desc = row.Description
                    rowDummy.Category_Code = row.Category_Code + " - " + row.Category_Desc
                    rowDummy.Category_Desc = row.Category_Desc
                    rowDummy.Sub_Brand_Code = row.Sub_Brand_Code + " - " + row.Sub_Brand_Desc
                    rowDummy.Sub_Brand_Desc = row.Sub_Brand_Desc
                    rowDummy.Brand_Code = row.Brand_Code + " - " + row.Brand_Desc
                    rowDummy.Brand_Desc = row.Brand_Desc
                    rowDummy.Group_Code = row.Group_Code + " - " + row.Group_Desc
                    rowDummy.Group_Desc = row.Group_Desc
                    TBLDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = TBLDummy.DefaultView(0).Row
                    rowDummy.Amount += row.Amount
                    rowDummy.Qty += row.Total_Qty
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
                ElseIf GroupBy = EnumGroupBy.Category Then
                    rowRekap.Group_By = rowDummy.Category_Code
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    rowRekap.Group_By = rowDummy.Sub_Brand_Code
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    rowRekap.Group_By = rowDummy.Brand_Code
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    rowRekap.Group_By = rowDummy.Group_Code
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    rowRekap.Group_By = rowDummy.Cust_Code
                End If

                'name1
                Name1 = Me.GetName1_Rekap
                If Name1 = EnumGroupBy.CodeNo Then
                    rowRekap.Name_1 = rowDummy.Product_Code
                ElseIf Name1 = EnumGroupBy.Category Then
                    rowRekap.Name_1 = rowDummy.Category_Code
                ElseIf Name1 = EnumGroupBy.SubBrand Then
                    rowRekap.Name_1 = rowDummy.Sub_Brand_Code
                ElseIf Name1 = EnumGroupBy.Brand Then
                    rowRekap.Name_1 = rowDummy.Brand_Code
                ElseIf Name1 = EnumGroupBy.ProductGroup Then
                    rowRekap.Name_1 = rowDummy.Group_Code
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
                Name1_Str = "Code No"
            ElseIf Name1 = EnumGroupBy.Category Then
                Name1_Str = "Category"
            ElseIf Name1 = EnumGroupBy.SubBrand Then
                Name1_Str = "Sub Brand"
            ElseIf Name1 = EnumGroupBy.Brand Then
                Name1_Str = "Brand"
            ElseIf Name1 = EnumGroupBy.ProductGroup Then
                Name1_Str = "Product Group"
            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Supplier Code"
            End If

            ReportRekap.SetDataSource(Me.DS2)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_Default_Precision.ParameterFieldName, Me.DefaultCurrency)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Purchase List")
            ReportRekap.SetParameterValue(ReportRekap.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_Name1.ParameterFieldName, Name1_Str)
            If GroupBy = EnumGroupBy.CodeNo Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Code No")
            ElseIf GroupBy = EnumGroupBy.Category Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Category")
            ElseIf GroupBy = EnumGroupBy.SubBrand Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Sub Brand")
            ElseIf GroupBy = EnumGroupBy.Brand Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Brand")
            ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Product Group")
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Supplier")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportRekap.SetParameterValue(ReportRekap.Parameter_Group_By.ParameterFieldName, "Date")
            End If

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
#End Region

    Private Function ApakahMasukKriteria(ByVal row1 As DseLaporanPembelian.tbltrans_aop_view2Row) As Boolean
        Dim ketemu As Boolean = False
        Dim rowView As DataRowView

        Try
            'cek Prod
            If Me.chkProduct.Checked = False Then
                ketemu = False
                For Each rowView In Me.clbProduct.CheckedItems
                    rowProd = rowView.Row
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
                For Each rowView In Me.clbCustomer.CheckedItems
                    rowSupp = rowView.Row
                    If rowSupp.Supplier_Code = row1.Supplier_Code Then
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
            GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Sort1 = "[Cust Code]"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Sort1 = "[Category Code]"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Sort1 = "[Sub Brand Code]"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Sort1 = "[Brand Code]"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Sort1 = "[Group Code]"
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
            GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Sort1 = "[Category Code]"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Sort1 = "[Sub Brand Code]"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Sort1 = "[Brand Code]"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Sort1 = "[Group Code]"
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
            If GroupBy = EnumGroupBy.CodeNo Or GroupBy = EnumGroupBy.Category Or GroupBy = EnumGroupBy.SubBrand Or GroupBy = EnumGroupBy.Brand Or GroupBy = EnumGroupBy.ProductGroup Then
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
            GroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboSort1.SelectedIndex = 0 Or Me.cboSort1.SelectedIndex = 1 Then
                If GroupBy = EnumGroupBy.CodeNo Then
                    Sort1 = "[Product Code]"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Sort1 = "[Category Code]"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Sort1 = "[Sub Brand Code]"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Sort1 = "[Brand Code]"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Sort1 = "[Group Code]"
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
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "1-st Period Amount" Then
                        Return "Amount1"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "2-nd Period Amount" Then
                        Return "Amount2"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return "[Product Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return "[Category Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return "[Sub Brand Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return "[Brand Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "[Group Code]"
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
                    If rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Then
                        Return "[Invoice Number]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return "Tanggal"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return "Product Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return "Category Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return "Sub Brand Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return "Brand Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "Group Code"
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
                    If rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
                        Return "[Invoice Number]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return "[Cust Code]"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return "Tanggal"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return "Product Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return "Category Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return "Sub Brand Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return "Brand Code"
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return "Group Code"
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

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPembelian.Detail1DataTable)
        Dim rowDummy As DseLaporanPembelian.Detail1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Qty, Amount As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            GroupBy = Me.cboGroupBy.SelectedIndex
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Category Then
                    If ListFilter.Contains(rowDummy.Category_Code) = False Then
                        ListFilter.Add(rowDummy.Category_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    If ListFilter.Contains(rowDummy.Sub_Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Sub_Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    If ListFilter.Contains(rowDummy.Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    If ListFilter.Contains(rowDummy.Group_Code) = False Then
                        ListFilter.Add(rowDummy.Group_Code)
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
                    Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
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

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPembelian.Rekap1DataTable)
        Dim rowDummy As DseLaporanPembelian.Rekap1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Qty, Amount As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            GroupBy = Me.cboGroupBy.SelectedIndex
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Category Then
                    If ListFilter.Contains(rowDummy.Category_Code) = False Then
                        ListFilter.Add(rowDummy.Category_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    If ListFilter.Contains(rowDummy.Sub_Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Sub_Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    If ListFilter.Contains(rowDummy.Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    If ListFilter.Contains(rowDummy.Group_Code) = False Then
                        ListFilter.Add(rowDummy.Group_Code)
                    End If
                End If
            Next

            For Each Filter In ListFilter
                If GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
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

    Private Sub HitungSubTotalQtyAmount(ByVal TBLDummy As DseLaporanPembelian.Rekap_Comparison1DataTable)
        Dim rowDummy As DseLaporanPembelian.Rekap_Comparison1Row
        Dim GroupBy As EnumGroupBy = EnumGroupBy.CodeNo
        Dim ListFilter As New List(Of String)
        Dim Filter As String = ""
        Dim Amount1, Amount2, Tot1, Tot2 As Decimal
        Dim rowView As DataRowView
        Dim FilterLama As String = ""

        Try
            FilterLama = TBLDummy.DefaultView.RowFilter

            GroupBy = Me.cboGroupBy.SelectedIndex
            For Each rowDummy In TBLDummy.Rows
                If GroupBy = EnumGroupBy.CustCode Then
                    If ListFilter.Contains(rowDummy.Cust_Code) = False Then
                        ListFilter.Add(rowDummy.Cust_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.CodeNo Then
                    If ListFilter.Contains(rowDummy.Product_Code) = False Then
                        ListFilter.Add(rowDummy.Product_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Category Then
                    If ListFilter.Contains(rowDummy.Category_Code) = False Then
                        ListFilter.Add(rowDummy.Category_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    If ListFilter.Contains(rowDummy.Sub_Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Sub_Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    If ListFilter.Contains(rowDummy.Brand_Code) = False Then
                        ListFilter.Add(rowDummy.Brand_Code)
                    End If
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    If ListFilter.Contains(rowDummy.Group_Code) = False Then
                        ListFilter.Add(rowDummy.Group_Code)
                    End If
                End If
            Next

            Tot1 = 0
            Tot2 = 0
            For Each Filter In ListFilter
                If GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Category Then
                    Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
                ElseIf GroupBy = EnumGroupBy.CustCode Then
                    Filter = "[Cust Code] = '" + ObjCD.FormatSpecialChar_DataFilter(Filter) + "'"
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

#Region "Coding Lama GenerateDetail"
    'Private Sub GenerateDetail()
    '    Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
    '    Dim rowView As DataRowView
    '    Dim TBLDummy As New DseLaporanPembelian.Detail1DataTable
    '    Dim rowDummy As DseLaporanPembelian.Detail1Row
    '    Dim rowDetail As DseLaporanPembelian.DetailRow
    '    Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
    '    Dim Name1, Name2, Name3 As EnumGroupBy
    '    Dim IndexMinimum As Integer = 0
    '    Dim Name1_Str, Name2_Str, Name3_Str As String
    '    Dim Lanjut As Boolean = False

    '    Try
    '        For Each row In Me.TBL.Rows
    '            Lanjut = True
    '            If Me.chkCustomer.Checked = True And Me.chkProduct.Checked = True Then
    '                Lanjut = True
    '            Else
    '                If Me.ApakahMasukKriteria(row) <> True Then
    '                    Lanjut = False
    '                End If
    '            End If
    '            If Lanjut = True Then
    '                If Me.cboGroupBy.SelectedIndex = 1 Then
    '                    If Me.ChkProductDesc.Checked = True Then
    '                        TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Description) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
    '                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "'"
    '                    Else
    '                        TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
    '                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "'"
    '                    End If
    '                Else
    '                    TBLDummy.DefaultView.RowFilter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
    '                    "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "'"
    '                End If

    '                If TBLDummy.DefaultView.Count = 0 Then
    '                    rowDummy = TBLDummy.NewDetail1Row
    '                    rowDummy.Amount = row.Amount
    '                    rowDummy.Cust_Code = row.Supplier_Name
    '                    rowDummy.Invoice_Number = row.No_Transaksi
    '                    If Me.cboGroupBy.SelectedIndex = 1 Then
    '                        If Me.ChkProductDesc.Checked = True Then
    '                            rowDummy.Product_Code = row.Description
    '                        Else
    '                            rowDummy.Product_Code = row.Product_Code
    '                        End If
    '                    Else
    '                        rowDummy.Product_Code = row.Product_Code
    '                    End If
    '                    rowDummy.Qty = row.Total_Qty
    '                    rowDummy.Tanggal = Format(row.AOP_Date, "dd MMM yyyy")
    '                    rowDummy._Decimal = row._Decimal
    '                    rowDummy.Unit = row.Unit_Code
    '                    rowDummy.Product_Desc = row.Description
    '                    TBLDummy.Rows.Add(rowDummy)
    '                Else
    '                    rowDummy = TBLDummy.DefaultView(0).Row
    '                    rowDummy.Amount += row.Amount
    '                    rowDummy.Qty += row.Total_Qty
    '                End If
    '            End If
    '        Next

    '        'hitung total group
    '        If Me.cboSort1.SelectedIndex <> 0 And Me.cboSort1.SelectedIndex <> 1 Then
    '            Me.HitungSubTotalQtyAmount(TBLDummy)
    '        End If

    '        'pindahkan dari tbldummy ke tbl yang akan disetorkan ke dalam laporan bergantung pada sorting
    '        TBLDummy.DefaultView.RowFilter = ""
    '        TBLDummy.DefaultView.Sort = Me.GenerateSortDetail2
    '        For Each rowView In TBLDummy.DefaultView
    '            rowDummy = rowView.Row
    '            rowDetail = Me.TBLDetail.NewDetailRow
    '            If GroupBy = EnumGroupBy.CodeNo Then
    '                rowDetail.Group_By = rowDummy.Product_Code + " - " + rowDummy.Product_Desc
    '            ElseIf GroupBy = EnumGroupBy.CustCode Then
    '                rowDetail.Group_By = rowDummy.Cust_Code
    '            ElseIf GroupBy = EnumGroupBy.Tanggal Then
    '                rowDetail.Group_By = rowDummy.Tanggal
    '            End If

    '            'name1
    '            Name1 = Me.GetName1
    '            If Name1 = EnumGroupBy.CodeNo Then
    '                rowDetail.Name_1 = rowDummy.Product_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Code No")
    '            ElseIf Name1 = EnumGroupBy.CustCode Then
    '                rowDetail.Name_1 = rowDummy.Cust_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
    '            ElseIf Name1 = EnumGroupBy.Invoice Then
    '                rowDetail.Name_1 = rowDummy.Invoice_Number
    '                IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
    '            ElseIf Name1 = EnumGroupBy.Tanggal Then
    '                rowDetail.Name_1 = rowDummy.Tanggal
    '                IndexMinimum = Me.GetSortNumberMinimum("Date")
    '            End If

    '            'name2
    '            Name2 = Me.GetName2_3(IndexMinimum)
    '            If Name2 = EnumGroupBy.CodeNo Then
    '                rowDetail.Name_2 = rowDummy.Product_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Code No")
    '            ElseIf Name2 = EnumGroupBy.CustCode Then
    '                rowDetail.Name_2 = rowDummy.Cust_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
    '            ElseIf Name2 = EnumGroupBy.Invoice Then
    '                rowDetail.Name_2 = rowDummy.Invoice_Number
    '                IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
    '            ElseIf Name2 = EnumGroupBy.Tanggal Then
    '                rowDetail.Name_2 = rowDummy.Tanggal
    '                IndexMinimum = Me.GetSortNumberMinimum("Date")
    '            End If

    '            'name 3
    '            Name3 = Me.GetName2_3(IndexMinimum)
    '            If Name3 = EnumGroupBy.CodeNo Then
    '                rowDetail.Name_3 = rowDummy.Product_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Code No")
    '            ElseIf Name3 = EnumGroupBy.CustCode Then
    '                rowDetail.Name_3 = rowDummy.Cust_Code
    '                IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
    '            ElseIf Name3 = EnumGroupBy.Invoice Then
    '                rowDetail.Name_3 = rowDummy.Invoice_Number
    '                IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
    '            ElseIf Name3 = EnumGroupBy.Tanggal Then
    '                rowDetail.Name_3 = rowDummy.Tanggal
    '                IndexMinimum = Me.GetSortNumberMinimum("Date")
    '            End If

    '            rowDetail.Qty = rowDummy.Qty
    '            rowDetail.Amount = rowDummy.Amount
    '            rowDetail._Decimal = rowDummy._Decimal
    '            rowDetail.Unit = rowDummy.Unit
    '            Me.TBLDetail.Rows.Add(rowDetail)
    '        Next

    '        Name1_Str = ""
    '        Name2_Str = ""
    '        Name3_Str = ""
    '        If Name1 = EnumGroupBy.CodeNo Then
    '            Name1_Str = "Code No"
    '        ElseIf Name1 = EnumGroupBy.CustCode Then
    '            Name1_Str = "Supplier"
    '        ElseIf Name1 = EnumGroupBy.Invoice Then
    '            Name1_Str = "Invoice"
    '        ElseIf Name1 = EnumGroupBy.Tanggal Then
    '            Name1_Str = "Date"
    '        End If
    '        If Name2 = EnumGroupBy.CodeNo Then
    '            Name2_Str = "Code No"
    '        ElseIf Name2 = EnumGroupBy.CustCode Then
    '            Name2_Str = "Supplier"
    '        ElseIf Name2 = EnumGroupBy.Invoice Then
    '            Name2_Str = "Invoice"
    '        ElseIf Name2 = EnumGroupBy.Tanggal Then
    '            Name2_Str = "Date"
    '        End If
    '        If Name3 = EnumGroupBy.CodeNo Then
    '            Name3_Str = "Code No"
    '        ElseIf Name3 = EnumGroupBy.CustCode Then
    '            Name3_Str = "Supplier"
    '        ElseIf Name3 = EnumGroupBy.Invoice Then
    '            Name3_Str = "Invoice"
    '        ElseIf Name3 = EnumGroupBy.Tanggal Then
    '            Name3_Str = "Date"
    '        End If

    '        ReportDetail.SetDataSource(Me.DS2)
    '        ReportDetail.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Purchase List")
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_Name1.ParameterFieldName, Name1_Str)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_Name2.ParameterFieldName, Name2_Str)
    '        ReportDetail.SetParameterValue(ReportDetail.Parameter_Name3.ParameterFieldName, Name3_Str)
    '        If GroupBy = EnumGroupBy.CodeNo Then
    '            ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Code No")
    '        ElseIf GroupBy = EnumGroupBy.CustCode Then
    '            ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Supplier")
    '        ElseIf GroupBy = EnumGroupBy.Tanggal Then
    '            ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Date")
    '        End If

    '        Me.crv1.ReportSource = ReportDetail
    '        Me.crv1.DisplayGroupTree = False
    '        Me.crv1.DisplayStatusBar = True
    '        Me.crv1.DisplayToolbar = True
    '        Me.crv1.ShowGroupTreeButton = False
    '        Me.crv1.ShowExportButton = True
    '        Me.crv1.ShowRefreshButton = False
    '        Me.crv1.Show()

    '        Me.TabControl1.SelectedIndex = 1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
#End Region

    Private Sub GenerateDetail()
        Dim row As DseLaporanPembelian.tbltrans_aop_view2Row
        Dim rowView As DataRowView
        Dim TBLDummy As New DseLaporanPembelian.Detail1DataTable
        Dim rowDummy As DseLaporanPembelian.Detail1Row
        Dim rowDetail As DseLaporanPembelian.DetailRow
        Dim GroupBy As EnumGroupBy = Me.cboGroupBy.SelectedIndex
        Dim Name1, Name2, Name3 As EnumGroupBy
        Dim IndexMinimum As Integer = 0
        Dim Name1_Str, Name2_Str, Name3_Str As String
        Dim Lanjut As Boolean = False

        Try
            For Each row In Me.TBL.Rows
                TBLDummy.DefaultView.RowFilter = Me.GetFilterGroupBy(GroupBy, row)
                If TBLDummy.DefaultView.Count = 0 Then
                    rowDummy = TBLDummy.NewDetail1Row
                    rowDummy.Amount = row.Amount
                    rowDummy.Cust_Code = row.Supplier_Code
                    rowDummy.Invoice_Number = row.No_Transaksi
                  
                    If Me.cboGroupBy.SelectedIndex = 5 Or Me.cboGroupBy.SelectedIndex = 6 Then
                        If Me.ChkProductDesc.Checked = True Then
                            rowDummy.Product_Code = row.Description
                            rowDummy.Category_Code = row.Category_Desc
                            rowDummy.Sub_Brand_Code = row.Sub_Brand_Desc
                            rowDummy.Brand_Code = row.Brand_Desc
                            rowDummy.Group_Code = row.Group_Desc
                        Else
                            rowDummy.Product_Code = row.Product_Code
                            rowDummy.Category_Code = row.Category_Code
                            rowDummy.Sub_Brand_Code = row.Sub_Brand_Code
                            rowDummy.Brand_Code = row.Brand_Code
                            rowDummy.Group_Code = row.Group_Code
                        End If
                    Else
                        rowDummy.Product_Code = row.Product_Code
                        rowDummy.Category_Code = row.Category_Code
                        rowDummy.Sub_Brand_Code = row.Sub_Brand_Code
                        rowDummy.Brand_Code = row.Brand_Code
                        rowDummy.Group_Code = row.Group_Code
                    End If
                    rowDummy.Qty = row.Total_Qty
                    rowDummy.Tanggal = Format(row.AOP_Date, "dd MMM yyyy")
                    rowDummy._Decimal = row._Decimal
                    rowDummy.Unit = row.Unit_Code
                    rowDummy.Product_Desc = row.Description
                    rowDummy.Category_Desc = row.Category_Desc
                    rowDummy.Sub_Brand_Desc = row.Sub_Brand_Desc
                    rowDummy.Brand_Desc = row.Brand_Desc
                    rowDummy.Group_Desc = row.Group_Desc
                    TBLDummy.Rows.Add(rowDummy)
                Else
                    rowDummy = TBLDummy.DefaultView(0).Row
                    rowDummy.Amount += row.Amount
                    rowDummy.Qty += row.Total_Qty
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
                ElseIf GroupBy = EnumGroupBy.Category Then
                    rowDetail.Group_By = rowDummy.Category_Code + " - " + rowDummy.Category_Desc
                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    rowDetail.Group_By = rowDummy.Sub_Brand_Code + " - " + rowDummy.Sub_Brand_Desc
                ElseIf GroupBy = EnumGroupBy.Brand Then
                    rowDetail.Group_By = rowDummy.Brand_Code + " - " + rowDummy.Brand_Desc
                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    rowDetail.Group_By = rowDummy.Group_Code + " - " + rowDummy.Group_Desc
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
                ElseIf Name1 = EnumGroupBy.Category Then
                    rowDetail.Name_1 = rowDummy.Category_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Category Code")
                ElseIf Name1 = EnumGroupBy.SubBrand Then
                    rowDetail.Name_1 = rowDummy.Sub_Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Sub Brand Code")
                ElseIf Name1 = EnumGroupBy.Brand Then
                    rowDetail.Name_1 = rowDummy.Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Brand Code")
                ElseIf Name1 = EnumGroupBy.ProductGroup Then
                    rowDetail.Name_1 = rowDummy.Group_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Group Code")
                ElseIf Name1 = EnumGroupBy.CustCode Then
                    rowDetail.Name_1 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
                ElseIf Name1 = EnumGroupBy.Invoice Then
                    rowDetail.Name_1 = rowDummy.Invoice_Number
                    If Me.CboDataFrom.SelectedIndex = 0 Then
                        IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
                    Else
                        IndexMinimum = Me.GetSortNumberMinimum("PO Number")
                    End If

                ElseIf Name1 = EnumGroupBy.Tanggal Then
                    rowDetail.Name_1 = rowDummy.Tanggal
                    IndexMinimum = Me.GetSortNumberMinimum("Date")
                End If

                'name2
                Name2 = Me.GetName2_3(IndexMinimum)
                If Name2 = EnumGroupBy.CodeNo Then
                    rowDetail.Name_2 = rowDummy.Product_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Code No")
                ElseIf Name2 = EnumGroupBy.Category Then
                    rowDetail.Name_2 = rowDummy.Category_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Category Code")
                ElseIf Name2 = EnumGroupBy.SubBrand Then
                    rowDetail.Name_2 = rowDummy.Sub_Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Sub Brand Code")
                ElseIf Name2 = EnumGroupBy.Brand Then
                    rowDetail.Name_2 = rowDummy.Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Brand Code")
                ElseIf Name2 = EnumGroupBy.ProductGroup Then
                    rowDetail.Name_2 = rowDummy.Group_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Group Code")
                ElseIf Name2 = EnumGroupBy.CustCode Then
                    rowDetail.Name_2 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
                ElseIf Name2 = EnumGroupBy.Invoice Then
                    rowDetail.Name_2 = rowDummy.Invoice_Number
                    If Me.CboDataFrom.SelectedIndex = 0 Then
                        IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
                    Else
                        IndexMinimum = Me.GetSortNumberMinimum("PO Number")
                    End If
                ElseIf Name2 = EnumGroupBy.Tanggal Then
                    rowDetail.Name_2 = rowDummy.Tanggal
                    IndexMinimum = Me.GetSortNumberMinimum("Date")
                End If

                'name 3
                Name3 = Me.GetName2_3(IndexMinimum)
                If Name3 = EnumGroupBy.CodeNo Then
                    rowDetail.Name_3 = rowDummy.Product_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Code No")
                ElseIf Name3 = EnumGroupBy.Category Then
                    rowDetail.Name_3 = rowDummy.Category_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Category Code")
                ElseIf Name3 = EnumGroupBy.SubBrand Then
                    rowDetail.Name_3 = rowDummy.Sub_Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Sub Brand Code")
                ElseIf Name3 = EnumGroupBy.Brand Then
                    rowDetail.Name_3 = rowDummy.Brand_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Brand Code")
                ElseIf Name3 = EnumGroupBy.ProductGroup Then
                    rowDetail.Name_3 = rowDummy.Group_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Group Code")
                ElseIf Name3 = EnumGroupBy.CustCode Then
                    rowDetail.Name_3 = rowDummy.Cust_Code
                    IndexMinimum = Me.GetSortNumberMinimum("Supplier Code")
                ElseIf Name3 = EnumGroupBy.Invoice Then
                    rowDetail.Name_3 = rowDummy.Invoice_Number
                    If Me.CboDataFrom.SelectedIndex = 0 Then
                        IndexMinimum = Me.GetSortNumberMinimum("AOP Number")
                    Else
                        IndexMinimum = Me.GetSortNumberMinimum("PO Number")
                    End If
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
            If Name1 = EnumGroupBy.CodeNo Then
                Name1_Str = "Code No"
            ElseIf Name1 = EnumGroupBy.Category Then
                Name1_Str = "Category Code"
            ElseIf Name1 = EnumGroupBy.SubBrand Then
                Name1_Str = "Sub Brand Code"
            ElseIf Name1 = EnumGroupBy.Brand Then
                Name1_Str = "Brand Code"
            ElseIf Name1 = EnumGroupBy.ProductGroup Then
                Name1_Str = "Group Code"
            ElseIf Name1 = EnumGroupBy.CustCode Then
                Name1_Str = "Supplier"
            ElseIf Name1 = EnumGroupBy.Invoice Then
                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name1_Str = "AOP Number"
                Else
                    Name1_Str = "PO Number"
                End If

            ElseIf Name1 = EnumGroupBy.Tanggal Then
                Name1_Str = "Date"
            End If

            If Name2 = EnumGroupBy.CodeNo Then
                Name2_Str = "Code No"
            ElseIf Name2 = EnumGroupBy.Category Then
                Name2_Str = "Category Code"
            ElseIf Name2 = EnumGroupBy.SubBrand Then
                Name2_Str = "Sub Brand Code"
            ElseIf Name2 = EnumGroupBy.Brand Then
                Name2_Str = "Brand Code"
            ElseIf Name2 = EnumGroupBy.ProductGroup Then
                Name2_Str = "Group Code"
            ElseIf Name2 = EnumGroupBy.CustCode Then
                Name2_Str = "Supplier"
            ElseIf Name2 = EnumGroupBy.Invoice Then
                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name2_Str = "AOP Number"
                Else
                    Name2_Str = "PO Number"
                End If
            ElseIf Name2 = EnumGroupBy.Tanggal Then
                Name2_Str = "Date"
            End If

            If Name3 = EnumGroupBy.CodeNo Then
                Name3_Str = "Code No"
            ElseIf Name3 = EnumGroupBy.Category Then
                Name3_Str = "Category"
            ElseIf Name3 = EnumGroupBy.SubBrand Then
                Name3_Str = "Sub Brand"
            ElseIf Name3 = EnumGroupBy.Brand Then
                Name3_Str = "Brand"
            ElseIf Name3 = EnumGroupBy.ProductGroup Then
                Name3_Str = "Product Group"
            ElseIf Name3 = EnumGroupBy.CustCode Then
                Name3_Str = "Supplier"
            ElseIf Name3 = EnumGroupBy.Invoice Then

                If Me.CboDataFrom.SelectedIndex = 0 Then
                    Name3_Str = "AOP Number"
                Else
                    Name3_Str = "PO Number"
                End If
            ElseIf Name3 = EnumGroupBy.Tanggal Then
                Name3_Str = "Date"
            End If

            ReportDetail.SetDataSource(Me.DS2)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Default_Precision.ParameterFieldName, Me.DefaultCurrency)
            ReportDetail.SetParameterValue(ReportRekap.Parameter_ReportName.ParameterFieldName, "Purchase List")
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, Me.dtpFrom.Value)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, Me.dtpTo.Value)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name1.ParameterFieldName, Name1_Str)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name2.ParameterFieldName, Name2_Str)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Name3.ParameterFieldName, Name3_Str)
            If GroupBy = EnumGroupBy.CodeNo Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Code No")
            ElseIf GroupBy = EnumGroupBy.Category Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Category")
            ElseIf GroupBy = EnumGroupBy.SubBrand Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Sub Brand")
            ElseIf GroupBy = EnumGroupBy.Brand Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Brand")
            ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Product Group")
            ElseIf GroupBy = EnumGroupBy.CustCode Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Supplier")
            ElseIf GroupBy = EnumGroupBy.Tanggal Then
                ReportDetail.SetParameterValue(ReportDetail.Parameter_Group_By.ParameterFieldName, "Date")
            End If

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
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next

            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next

            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
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
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortRekap.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
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
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = SortNumberMinimum + 2 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
                        Return EnumGroupBy.Invoice
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvSortDetail.Rows
                If rowDGV.Cells(Me.ColSortNumber.Index).Value = SortNumberMinimum + 3 Then
                    If rowDGV.Cells(Me.ColName.Index).Value = "Supplier Code" Then
                        Return EnumGroupBy.CustCode
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Date" Then
                        Return EnumGroupBy.Tanggal
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Code No" Then
                        Return EnumGroupBy.CodeNo
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Category" Then
                        Return EnumGroupBy.Category
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Sub Brand" Then
                        Return EnumGroupBy.SubBrand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Brand" Then
                        Return EnumGroupBy.Brand
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "Product Group" Then
                        Return EnumGroupBy.ProductGroup
                    ElseIf rowDGV.Cells(Me.ColName.Index).Value = "AOP Number" Or rowDGV.Cells(Me.ColName.Index).Value = "PO Number" Then
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

    'Private Sub dgvSortRekap_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs)
    '    Try
    '        'yang bisa diedit hanya masalah urutan dan jenis sort
    '        If e.ColumnIndex <> Me.ColSortNumber.Index And e.ColumnIndex <> Me.ColSrtType.Index Then
    '            e.Cancel = True
    '        Else
    '            Me.Buffer = Me.dgvSortRekap.Item(e.ColumnIndex, e.RowIndex).Value
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub dgvSortRekap_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '    Dim rowDGV As DataGridViewRow

    '    Try
    '        If e.ColumnIndex = Me.ColSortNumber.Index Then
    '            For Each rowDGV In Me.dgvSortRekap.Rows
    '                If rowDGV.Index <> e.RowIndex Then
    '                    If rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.dgvSortRekap.Item(e.ColumnIndex, e.RowIndex).Value Then
    '                        rowDGV.Cells(Me.ColSortNumber.Index).Value = Me.Buffer
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

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

#Region "Filteran Dgv"
    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.Buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError

    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemProductGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

    Private Sub DgvSupplier_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSupplier.CellBeginEdit
        Try
            If Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                    Me.Buffer = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                    Me.Buffer = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSupplier.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSupplier.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Supplier As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSupp_Code.Index Then
                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjSupplier.ApakahSudahAda(Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Supplier = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(Supplier, e.RowIndex)
                            If index < 0 Then
                                Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value)
                                Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value
                                Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwSupplier.RunWorkerAsync()
                                    Else
                                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSupplier
                            ij.TxtCari.Text = Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimarySupplier <> -1 Then
                                index = Me.GetIndexItem(ij.PrimarySupplier, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                                    Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(ij.PrimarySupplier)
                                    Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwSupplier.RunWorkerAsync()
                                        Else
                                            Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvSupplier.Item(Me.ColSupp_Code.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                                        Me.DgvSupplier.Item(Me.ColSupp_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(ij.PrimarySupplier)
                                        Me.DgvSupplier.Item(Me.ColSupp_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                                    End If
                                End If
                            ElseIf ij.PrimarySupplier = -1 Then
                                If e.ColumnIndex = Me.ColSupp_Code.Index Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        If Me.DgvSupplier.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwSupplier.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                Else
                                    Me.DgvSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSupplier)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwSupplier.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeSupplier As String, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSupplier.Rows.Count - 2
                RowDGV = Me.DgvSupplier.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSupp_Code.Index).Value = vKodeSupplier Then
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

    Private Sub BgwSupplier_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwSupplier.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwSupplier_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwSupplier.RunWorkerCompleted
        Try
            Me.DgvSupplier.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvSupplier.DataError

    End Sub

    Private Sub DgvSupplier_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSupplier.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvSupplier.Rows(Me.DgvSupplier.RowCount - 1)
            With rowDGV
                ' .Cells(Me.ColN.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColSupp_Code.Index).Value = ""
                .Cells(Me.ColSupp_Name.Index).Value = ""
                .Cells(Me.ColSupp_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSupplier.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSupplier_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSupplier.Sorted
        Me.ObjCD.SetNumber(Me.DgvSupplier)
    End Sub

    Private Sub DgvSupplier_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvSupplier.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboItemDetail_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboItemDetail.SelectedIndexChanged
        Try
            IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSortRekap_CellBeginEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvSortRekap.CellBeginEdit
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

    Private Function GetFilterGroupBy(ByVal GroupBy As EnumGroupBy, ByVal row As DseLaporanPembelian.tbltrans_aop_view2Row) As String
        Dim Filter As String = ""
        Try
            If Me.cboFormat.SelectedIndex = 2 Then
                If Me.ChkProductDesc.Checked = True Then
                    If GroupBy = EnumGroupBy.CodeNo Then
                        Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Description) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.Category Then
                        Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                       "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.SubBrand Then
                        Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                      "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.Brand Then
                        Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                        Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                       "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    Else
                        'item detail kalo pilih supplier
                        If Me.CboItemDetail.SelectedIndex = 0 Then
                            Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Description) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                  "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                        ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                            Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                        ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                            Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                        ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                            Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                  "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                        ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                            Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Desc) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        End If
                    End If

                Else
                    If GroupBy = EnumGroupBy.CodeNo Then
                        Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.Category Then
                        Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                       "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.SubBrand Then
                        Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                      "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.Brand Then
                        Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                        "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                        Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                       "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"

                    Else
                        'item detail kalo pilih supplier
                        If Me.CboItemDetail.SelectedIndex = 0 Then
                            Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                            Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                 "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                            Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                               "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                            Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                   "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                            Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Code) + "' AND Tanggal =#" + row.AOP_Date.ToShortDateString + "# AND " + _
                                                 "[Invoice Number] ='" + row.No_Transaksi + "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + "' AND [Unit]='" + row.Unit_Code + "'"
                        End If
                    End If
                End If
            Else
                If GroupBy = EnumGroupBy.CodeNo Then
                    Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                    "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                ElseIf GroupBy = EnumGroupBy.Category Then
                    Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                    "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                ElseIf GroupBy = EnumGroupBy.SubBrand Then
                    Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                    "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                ElseIf GroupBy = EnumGroupBy.Brand Then
                    Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                    "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                ElseIf GroupBy = EnumGroupBy.ProductGroup Then
                    Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Group_Desc) + _
                    "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"
                Else
                    'item detail kalo pilih supplier
                    If Me.CboItemDetail.SelectedIndex = 0 Then
                        Filter = "[Product Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Product_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Description) + _
                        "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf Me.CboItemDetail.SelectedIndex = 1 Then
                        Filter = "[Category Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Category_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Category_Desc) + _
                        "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf Me.CboItemDetail.SelectedIndex = 2 Then
                        Filter = "[Sub Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Sub_Brand_Desc) + _
                        "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf Me.CboItemDetail.SelectedIndex = 3 Then
                        Filter = "[Brand Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Brand_Desc) + _
                        "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                    ElseIf Me.CboItemDetail.SelectedIndex = 4 Then
                        Filter = "[Group Code] = '" + ObjCD.FormatSpecialChar_DataFilter(row.Group_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Group_Desc) + _
                        "' AND [Cust Code] ='" + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Code) + " - " + ObjCD.FormatSpecialChar_DataFilter(row.Supplier_Name) + "' AND [Unit]='" + row.Unit_Code + "'"

                    End If
                End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CboDataFrom_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboDataFrom.SelectedIndexChanged
        Try
            Me.IsikanDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class