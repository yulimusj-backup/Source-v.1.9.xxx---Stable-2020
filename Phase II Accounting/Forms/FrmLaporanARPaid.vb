Option Explicit On
Option Strict Off

Imports AccModule, Core_Module
Public Class FrmLaporanARPaid
    Dim PrimaryMinInv, PrimaryMaxInv, PrimaryMinSR, PrimaryMaxSR, PrimaryMinHistoricalSales, PrimaryMaxHistoricalSales As Double
    Dim PrimaryMinHistoricalSalesReturn, PrimaryMaxHistoricalSalesReturn, PrimaryMinCreditNote, PrimaryMaxCreditNote, PrimaryMinDebitNote, PrimaryMaxDebitNote As Double
    Dim PrimaryMinSO, PrimaryMaxSO As Double
    Dim TipeTransaksiAsal As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
    Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    Dim TipeJurnalAR As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
    Dim TipeJurnalAP As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
    Dim TipeStockInSR As clsTransaksiStockInOut1.EnumTipeStockInOut = clsTransaksiStockInOut1.EnumTipeStockInOut.StockIn_SalesReturn
    'Debit Kredit Note
    Dim TipeTransaksiDebitNote As ClsTransaksiCreditNote.EnumTipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE
    Dim TipeTransaksiKreditNote As ClsTransaksiCreditNote.EnumTipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE
    Dim ObjLicense As New Core_Module.ClsLicense

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
    Dim TipeTransaksiAsalKartu As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis

    Dim ObjCust As New clsMasterCustomer
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjProd As New clsMasterProduct
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSupp As New clsMasterSupplier
    Dim ObjPO As New ClsTransaksiPO

    '*********ALL DATA*********
    Dim DTA_Pay As New DseLaporanARPaidTableAdapters.Item_PaymentTableAdapter
    Dim TBL_Pay As New DseLaporanARPaid.Item_PaymentDataTable
    Dim Row_Pay As DseLaporanARPaid.Item_PaymentRow
    Dim DTADNote As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
    Dim TBLDNote As New DseLaporanARPaid.tbltrans_credit_note1DataTable
    Dim RowDNote As DseLaporanARPaid.tbltrans_credit_note1Row
    Dim RowDNoteDummy As DseLaporanARPaid.tbltrans_credit_note1Row
    Dim DTACNote As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
    Dim TBLCNote As New DseLaporanARPaid.tbltrans_credit_note1DataTable
    Dim RowCNote As DseLaporanARPaid.tbltrans_credit_note1Row
    Dim RowCNoteDummy As DseLaporanARPaid.tbltrans_credit_note1Row
    'Laporan Customer PAID (AR)
    Dim DTASaldo_AR As New DseLaporanARPaidTableAdapters.TBLSaldoTableAdapter
    Dim TBLSaldo_AR As New DseLaporanARPaid.TBLSaldoDataTable
    Dim RowSaldo_AR As DseLaporanARPaid.TBLSaldoRow
    Dim DTAInv As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoiceTableAdapter
    Dim TBLInv As New DseLaporanARPaid.tbltrans_salesinvoiceDataTable
    Dim RowInv As DseLaporanARPaid.tbltrans_salesinvoiceRow
    Dim RowInvDummy As DseLaporanARPaid.tbltrans_salesinvoiceRow
    Dim DTAInvItem As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoice_itemTableAdapter
    Dim TBLInvItem As New DseLaporanARPaid.tbltrans_salesinvoice_itemDataTable
    Dim RowInvItem As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
    Dim DTA_SR As New DseLaporanARPaidTableAdapters.tbltransaksi_sales_return1TableAdapter
    Dim TBL_SR As New DseLaporanARPaid.tbltransaksi_sales_return1DataTable
    Dim Row_SR As DseLaporanARPaid.tbltransaksi_sales_return1Row
    Dim Row_SRDummy As DseLaporanARPaid.tbltransaksi_sales_return1Row
    Dim DTA_SRItem As New DseLaporanARPaidTableAdapters.tbltrans_salesreturn_itemTableAdapter
    Dim TBL_SRItem As New DseLaporanARPaid.tbltrans_salesreturn_itemDataTable
    Dim Row_SRItem As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
    Dim DTA_HisSales As New DseLaporanARPaidTableAdapters.tbltrans_historicalsalesTableAdapter
    Dim TBL_HisSales As New DseLaporanARPaid.tbltrans_historicalsalesDataTable
    Dim Row_HisSales As DseLaporanARPaid.tbltrans_historicalsalesRow
    Dim DTA_HisSR As New DseLaporanARPaidTableAdapters.tbltrans_historical_sales_returnTableAdapter
    Dim TBL_HisSR As New DseLaporanARPaid.tbltrans_historical_sales_returnDataTable
    Dim Row_HisSR As DseLaporanARPaid.tbltrans_historical_sales_returnRow
    Dim DTA_SO As New DseLaporanARPaidTableAdapters.tbltrans_soTableAdapter
    Dim TBL_SO As New DseLaporanARPaid.tbltrans_soDataTable
    Dim Row_SO As DseLaporanARPaid.tbltrans_soRow
    'aul add
    Dim DTA_HisSalesDP, DTA_HisPurchDP As New DseLaporanARPaidTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
    Dim TBL_HisSalesDP, TBL_HisPurchDP As New DseLaporanARPaid.tbltrans_historicalbalancemanual_dpDataTable
    Dim Row_HisSalesDP, Row_HisPurchDP As DseLaporanARPaid.tbltrans_historicalbalancemanual_dpRow
    'Laporan Supplier PAID (AP)
    Dim DTASaldo_AP As New DseLaporanARPaidTableAdapters.TBLSaldo1TableAdapter
    Dim TBLSaldo_AP As New DseLaporanARPaid.TBLSaldo1DataTable
    Dim RowSaldo_AP As DseLaporanARPaid.TBLSaldo1Row
    Dim DTA_AOP As New DseLaporanARPaidTableAdapters.tbltrans_aopTableAdapter
    Dim TBL_AOP As New DseLaporanARPaid.tbltrans_aopDataTable
    Dim Row_AOP As DseLaporanARPaid.tbltrans_aopRow
    Dim Row_AOPDummy As DseLaporanARPaid.tbltrans_aopRow
    Dim DTA_AOP_Item As New DseLaporanARPaidTableAdapters.tbltrans_aop_itemTableAdapter
    Dim TBL_AOP_Item As New DseLaporanARPaid.tbltrans_aop_itemDataTable
    Dim Row_AOP_Item As DseLaporanARPaid.tbltrans_aop_itemRow
    Dim DTA_AOP_PR As New DseLaporanARPaidTableAdapters.tbltrans_purchasereturnTableAdapter
    Dim TBL_AOP_PR As New DseLaporanARPaid.tbltrans_purchasereturnDataTable
    Dim Row_AOP_PR As DseLaporanARPaid.tbltrans_purchasereturnRow
    Dim Row_AOP_PRDummy As DseLaporanARPaid.tbltrans_purchasereturnRow
    Dim DTA_AOP_PR_Item As New DseLaporanARPaidTableAdapters.tbltrans_purchasereturn_itemTableAdapter
    Dim TBL_AOP_PR_Item As New DseLaporanARPaid.tbltrans_purchasereturn_itemDataTable
    Dim row_AOP_PR_Item As DseLaporanARPaid.tbltrans_purchasereturn_itemRow
    Dim DTA_HisPurc As New DseLaporanARPaidTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
    Dim TBL_HisPurc As New DseLaporanARPaid.tbltrans_historicalpurchasebalancemanualDataTable
    Dim Row_HisPurc As DseLaporanARPaid.tbltrans_historicalpurchasebalancemanualRow
    Dim DTA_HisPR As New DseLaporanARPaidTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
    Dim TBL_HisPR As New DseLaporanARPaid.tbltrans_historicalpurchasereturnbalancemanualDataTable
    Dim Row_HisPR As DseLaporanARPaid.tbltrans_historicalpurchasereturnbalancemanualRow
    Dim DTA_PO As New DseLaporanARPaidTableAdapters.tbltrans_poTableAdapter
    Dim TBL_PO As New DseLaporanARPaid.tbltrans_poDataTable
    Dim Row_PO As DseLaporanARPaid.tbltrans_poRow
    '*****TBL_Balance Sales*****
    Dim DTABal_SalesAutomatic As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesAutomatic As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesAutomatic As DseLaporanARPaid.tbltrans_balance_salesRow
    Dim DTABal_SalesAutomaticSR As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesAutomaticSR As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesAutomaticSR As DseLaporanARPaid.tbltrans_balance_salesRow
    Dim DTABal_SalesManual As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesManual As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesManual As DseLaporanARPaid.tbltrans_balance_salesRow
    Dim DTABal_SalesManualSR As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesManualSR As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesManualSR As DseLaporanARPaid.tbltrans_balance_salesRow
    Dim DTABal_SalesDNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesDNote As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesDNote As DseLaporanARPaid.tbltrans_balance_salesRow
    Dim DTABal_SalesCNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    Dim TBLBal_SalesCNote As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    Dim RowBal_SalesCNote As DseLaporanARPaid.tbltrans_balance_salesRow
    '*****TBL Balance Purchase*****
    Dim DTABal_PurcAutomatic As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcAutomatic As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcAutomatic As DseLaporanARPaid.tbltrans_balance_purchaseRow
    Dim DTABal_PurcAutomaticPR As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcAutomaticPR As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcAutomaticPR As DseLaporanARPaid.tbltrans_balance_purchaseRow
    Dim DTABal_PurcManual As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcManual As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcManual As DseLaporanARPaid.tbltrans_balance_purchaseRow
    Dim DTABal_PurcManualPR As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcManualPR As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcManualPR As DseLaporanARPaid.tbltrans_balance_purchaseRow
    Dim DTABal_PurcDNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcDNote As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcDNote As DseLaporanARPaid.tbltrans_balance_purchaseRow
    Dim DTABal_PurcCNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    Dim TBLBal_PurcCNote As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    Dim RowBal_PurcCNote As DseLaporanARPaid.tbltrans_balance_purchaseRow

    Private Sub GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.GotFocus, TxtCodeNo.GotFocus, _
    txtCurrencyCode.GotFocus, TxtCodeNo.GotFocus, TxtSONumber.GotFocus, TxtSuppCode.GotFocus, TxtPONumber.GotFocus, txtArAccountCode.GotFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        Try
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.LostFocus, TxtCodeNo.LostFocus, _
    txtCurrencyCode.LostFocus, TxtCodeNo.LostFocus, TxtSONumber.LostFocus, TxtSuppCode.LostFocus, TxtPONumber.LostFocus, txtArAccountCode.LostFocus
        Dim txt As TextBox = TryCast(sender, TextBox)
        Try
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanObjek()
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Me.CboPayment.Items.Add("Receivables Unfinished")
                Me.CboPayment.Items.Add("All Data")
                Me.Text = "AR Balance"
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                Me.CboPayment.Items.Add("Payables Unfinished")
                Me.CboPayment.Items.Add("All Data")
                Me.Text = "AP Balance"
            End If
            Me.DtpAsOf.Value = Now
            Me.CboPayment.SelectedIndex = 1
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.txtPrimaryCust.Text = "-1"
            Me.txtCustCode.Clear()
            Me.txtCustDesc.Clear()
            Me.txtPrimaryArAccount.Text = "-1"
            Me.txtArAccountCode.Clear()
            Me.txtArAccountDesc.Clear()
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtCurrencyCode.Clear()
            Me.txtCurrencyDesc.Clear()
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.TxtCodeNo.Clear()
            Me.TxtDesc.Clear()
            Me.TxtPrimarySO.Text = "-1"
            Me.TxtSONumber.Clear()
            Me.TxtCustomerPONumber.Clear()
            Me.TxtPrimarySupplier.Text = "-1"
            Me.TxtSuppCode.Clear()
            Me.TxtSupplierName.Clear()
            Me.TxtPrimaryPO.Text = "-1"
            Me.TxtPONumber.Clear()
            Me.VisibleObjek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub VisibleObjek()
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Me.LblSuppNo.Visible = False
                Me.TxtSuppCode.Visible = False
                Me.btnFindSupplier.Visible = False
                Me.TxtSupplierName.Visible = False
                Me.LblPONo.Visible = False
                Me.TxtPONumber.Visible = False
                Me.btnFindPO.Visible = False
            Else
                Me.LblCustCode.Visible = False
                Me.btnFindCust.Visible = False
                Me.txtCustCode.Visible = False
                Me.txtCustDesc.Visible = False
                Me.LblSONo.Visible = False
                Me.btnFindSO.Visible = False
                Me.TxtSONumber.Visible = False
                Me.LblCustPONo.Visible = False
                Me.TxtCustomerPONumber.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function AbaikanHistorical(ByVal vTipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu) As Boolean
        Try
            If vTipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If Me.TxtPrimaryCodeNo.Text = "-1" And Me.TxtPrimarySO.Text = "-1" Then
                    Return False
                Else
                    Return True
                End If
            ElseIf vTipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TxtPrimaryCodeNo.Text = "-1" And Me.TxtPrimaryPO.Text = "-1" Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetFilterHeader() As String
        Dim Retval As String = ""
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If Me.txtPrimaryCust.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Customer] = " & Me.txtPrimaryCust.Text.ToString
                    Else
                        Retval = "[Primary Customer] = " & Me.txtPrimaryCust.Text.ToString
                    End If
                End If
                If Me.txtPrimaryArAccount.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [primary akun ar] = " & Me.txtPrimaryArAccount.Text.ToString
                    Else
                        Retval = "[primary akun ar] = " & Me.txtPrimaryArAccount.Text.ToString
                    End If
                End If
                If Me.txtPrimaryCurrency.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Currency] = " & Me.txtPrimaryCurrency.Text.ToString
                    Else
                        Retval = "[Primary Currency] = " & Me.txtPrimaryCurrency.Text.ToString
                    End If
                End If
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TxtPrimarySupplier.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Supplier] = " & Me.TxtPrimarySupplier.Text.ToString
                    Else
                        Retval = "[Primary Supplier] = " & Me.TxtPrimarySupplier.Text.ToString
                    End If
                End If
                If Me.txtPrimaryArAccount.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [primary akun ap] = " & Me.txtPrimaryArAccount.Text.ToString
                    Else
                        Retval = "[primary akun ap] = " & Me.txtPrimaryArAccount.Text.ToString
                    End If
                End If
                If Me.txtPrimaryCurrency.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Currency] = " & Me.txtPrimaryCurrency.Text.ToString
                    Else
                        Retval = "[Primary Currency] = " & Me.txtPrimaryCurrency.Text.ToString
                    End If
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetFilterItem() As String
        Dim Retval As String = ""
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Product] = " & Me.TxtPrimaryCodeNo.Text.ToString
                    Else
                        Retval = "[Primary Product] = " & Me.TxtPrimaryCodeNo.Text.ToString
                    End If
                End If
                If Me.TxtPrimarySO.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary SO] = " & Me.TxtPrimarySO.Text.ToString
                    Else
                        Retval = "[Primary SO] = " & Me.TxtPrimarySO.Text.ToString
                    End If
                End If
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary Product] = " & Me.TxtPrimaryCodeNo.Text.ToString
                    Else
                        Retval = "[Primary Product] = " & Me.TxtPrimaryCodeNo.Text.ToString
                    End If
                End If
                If Me.TxtPrimaryPO.Text <> "-1" Then
                    If String.IsNullOrEmpty(Retval) = False Then
                        Retval += " AND [Primary PO] = " & Me.TxtPrimaryPO.Text.ToString
                    Else
                        Retval = "[Primary PO] = " & Me.TxtPrimaryPO.Text.ToString
                    End If
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahAbaikanKriteriaFilterData() As Boolean
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If Me.txtPrimaryCust.Text = "-1" And Me.txtPrimaryArAccount.Text = "-1" _
                And Me.TxtPrimaryCodeNo.Text = "-1" And Me.txtPrimaryCurrency.Text = "-1" _
                And Me.TxtPrimarySO.Text = "-1" Then
                    Return True
                Else
                    Return False
                End If
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TxtPrimarySupplier.Text = "-1" And Me.txtPrimaryArAccount.Text = "-1" _
                And Me.TxtPrimaryCodeNo.Text = "-1" And Me.txtPrimaryCurrency.Text = "-1" _
                And Me.TxtPrimaryPO.Text = "-1" Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GetPrimaryMinMax_SalesOrder(ByVal vTBLInvItem As DseLaporanARPaid.tbltrans_salesinvoice_itemDataTable)
        Dim vRowInvItem As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
        Try
            'Get Primary Min Max SO
            vTBLInvItem.DefaultView.RowFilter = Nothing
            vTBLInvItem.DefaultView.Sort = Nothing
            vTBLInvItem.DefaultView.Sort = "[Primary SO]"
            If vTBLInvItem.DefaultView.Count > 0 Then
                vRowInvItem = vTBLInvItem.DefaultView(0).Row
                Me.PrimaryMinSO = vRowInvItem.Primary_SO
                vRowInvItem = vTBLInvItem.DefaultView(vTBLInvItem.DefaultView.Count - 1).Row
                Me.PrimaryMaxSO = vRowInvItem.Primary_SO
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetPrimaryMinMax_PurchaseOrder(ByVal vTBLAOPItem As DseLaporanARPaid.tbltrans_aop_itemDataTable)
        Dim vRowAOPItem As DseLaporanARPaid.tbltrans_aop_itemRow
        Try
            'Get Primary Min Max SO
            vTBLAOPItem.DefaultView.RowFilter = Nothing
            vTBLAOPItem.DefaultView.Sort = Nothing
            vTBLAOPItem.DefaultView.Sort = "[Primary PO]"
            If vTBLAOPItem.DefaultView.Count > 0 Then
                vRowAOPItem = vTBLAOPItem.DefaultView(0).Row
                Me.PrimaryMinSO = vRowAOPItem.Primary_PO
                vRowAOPItem = vTBLAOPItem.DefaultView(vTBLAOPItem.DefaultView.Count - 1).Row
                Me.PrimaryMaxSO = vRowAOPItem.Primary_PO
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FilterDataByPrimaryCurrency(ByVal TBLInv As DseLaporanARPaid.tbltrans_salesinvoiceDataTable, _
    ByVal TBL_SR As DseLaporanARPaid.tbltransaksi_sales_return1DataTable, _
    ByVal TBLCNote As DseLaporanARPaid.tbltrans_credit_note1DataTable, _
    ByVal TBLDNote As DseLaporanARPaid.tbltrans_credit_note1DataTable, _
    ByVal TBL_AOP As DseLaporanARPaid.tbltrans_aopDataTable, _
    ByVal TBL_AOP_PR As DseLaporanARPaid.tbltrans_purchasereturnDataTable)
        Dim ij As Integer = 0
        Dim RowInv As DseLaporanARPaid.tbltrans_salesinvoiceRow
        Dim Row_SR As DseLaporanARPaid.tbltransaksi_sales_return1Row
        Dim Row_AOP As DseLaporanARPaid.tbltrans_aopRow
        Dim Row_AOP_PR As DseLaporanARPaid.tbltrans_purchasereturnRow
        Try
            If Me.txtPrimaryCurrency.Text <> "-1" Then
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                    '*****Kartu AR*****
                    If Me.ApakahAbaikanKriteriaFilterData() = False Then
                        'TBL Invoice 
                        ij = 0
                        While ij <= TBLInv.Rows.Count - 1
                            RowInv = TBLInv.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> RowInv.Primary_Currency Then
                                TBLInv.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        'TBL Sales Return
                        ij = 0
                        While ij <= TBL_SR.Rows.Count - 1
                            Row_SR = TBL_SR.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> Row_SR.Primary_Currency Then
                                TBL_SR.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        'TBL Credit Note
                        ij = 0
                        While ij <= TBLCNote.Rows.Count - 1
                            RowCNote = TBLCNote.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> RowCNote.Primary_Currency Then
                                TBLCNote.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While

                        'TBL Debit Note
                        ij = 0
                        While ij <= TBLDNote.Rows.Count - 1
                            RowDNote = TBLDNote.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> RowDNote.Primary_Currency Then
                                TBLDNote.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                    '*****Kartu AP*****
                    If Me.ApakahAbaikanKriteriaFilterData() = False Then
                        'TBL AOP 
                        ij = 0
                        While ij <= TBL_AOP.Rows.Count - 1
                            Row_AOP = TBL_AOP.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> Row_AOP.Primary_Currency Then
                                TBL_AOP.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        'TBL Purchase Return
                        ij = 0
                        While ij <= TBL_AOP_PR.Rows.Count - 1
                            Row_AOP_PR = TBL_AOP_PR.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> Row_AOP_PR.Primary_Currency Then
                                TBL_AOP_PR.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        'TBL Credit Note
                        ij = 0
                        While ij <= TBLCNote.Rows.Count - 1
                            RowCNote = TBLCNote.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> RowCNote.Primary_Currency Then
                                TBLCNote.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                        'TBL Debit Note
                        ij = 0
                        While ij <= TBLDNote.Rows.Count - 1
                            RowDNote = TBLDNote.Rows(ij)
                            If Me.txtPrimaryCurrency.Text <> RowDNote.Primary_Currency Then
                                TBLDNote.Rows.RemoveAt(ij)
                            Else
                                ij += 1
                            End If
                        End While
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTransaksiAmountInvoice(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double))
        Dim DV01 As DataRowView
        Dim DV02 As DataRowView
        Try
            ListPrimarySOPONumber.Clear()
            'Get Transaction Amount Invoice
            If TBLBal_SalesAutomatic.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                For Each RowBal_SalesAutomatic In TBLBal_SalesAutomatic.Rows
                    TBLInv.DefaultView.RowFilter = Nothing
                    TBLInv.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesAutomatic.Primary_Invoice.ToString
                    If TBLInv.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_SalesAutomatic.Primary_Invoice) = False Then
                            ListPrimary.Add(RowBal_SalesAutomatic.Primary_Invoice)
                            RowInv = TBLInv.DefaultView(0).Row
                            'Input Transaction Amount
                            RowInv.Transaction_Amount += RowBal_SalesAutomatic.balance_original + RowBal_SalesAutomatic.received_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowInv.Primary_Transaksi.ToString & _
                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                            " AND [Primary Currency] = " & RowInv.Primary_Currency.ToString

                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                RowInv.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If
                            'Input Cust PO No
                            TBLInvItem.DefaultView.RowFilter = Nothing
                            TBLInvItem.DefaultView.RowFilter = "[Primary Transaksi] = " & RowInv.Primary_Transaksi.ToString
                            If TBLInvItem.DefaultView.Count > 0 Then
                                TBL_SO.DefaultView.RowFilter = Nothing
                                TBL_SO.DefaultView.Sort = Nothing
                                For Each DV01 In TBLInvItem.DefaultView
                                    RowInvItem = DV01.Row
                                    TBL_SO.DefaultView.RowFilter = "[Primary SO] = " & RowInvItem.Primary_SO.ToString
                                    TBL_SO.DefaultView.Sort = "[Cust PO No]"
                                    If TBL_SO.DefaultView.Count > 0 Then
                                        For Each DV02 In TBL_SO.DefaultView
                                            Row_SO = DV02.Row
                                            If String.IsNullOrEmpty(Row_SO.Cust_PO_No) = False Then
                                                If ListPrimarySOPONumber.Contains(Row_SO.Primary_SO) = False Then
                                                    ListPrimarySOPONumber.Add(Row_SO.Primary_SO)
                                                    If String.IsNullOrEmpty(RowInv.Cust_PO_Number) = False Then
                                                        RowInv.Cust_PO_Number += ", " + Row_SO.Cust_PO_No.ToString
                                                    Else
                                                        RowInv.Cust_PO_Number = Row_SO.Cust_PO_No.ToString
                                                    End If
                                                End If
                                            Else
                                                If ListPrimarySOPONumber.Contains(Row_SO.Primary_SO) = False Then
                                                    If String.IsNullOrEmpty(RowInv.Cust_PO_Number) = False Then
                                                        RowInv.Cust_PO_Number += ", " + Row_SO.No_Transaksi.ToString
                                                    Else
                                                        RowInv.Cust_PO_Number = Row_SO.No_Transaksi.ToString
                                                    End If
                                                    ListPrimarySOPONumber.Add(Row_SO.Primary_SO)
                                                End If

                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            'Input Baris Baru
                            If RowBal_SalesAutomatic.Primary_Currency <> 1 Then
                                If RowInv.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    RowInvDummy = TBLInv.NewRow
                                    For xy As Integer = 0 To TBLInv.Columns.Count - 1
                                        RowInvDummy.Item(xy) = RowInv.Item(xy)
                                    Next
                                    TBLInv.Rows.Add(RowInvDummy)
                                    'Currency Code-Nya IDR
                                    RowInvDummy = TBLInv.Rows(TBLInv.Rows.Count - 1)
                                    RowInvDummy.Primary_Currency = 1
                                    RowInvDummy.Transaction_Amount = 0
                                    RowInvDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_SalesAutomatic.Primary_Invoice) = True Then
                            If RowBal_SalesAutomatic.Primary_Currency = 1 Then
                                RowInv = TBLInv.DefaultView(0).Row
                                If RowInv.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    RowInv = TBLInv.Rows(TBLInv.Rows.Count - 1)
                                    RowInv.Transaction_Amount += RowBal_SalesAutomatic.balance_original + RowBal_SalesAutomatic.received_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowInv.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & RowInv.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        RowInv.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTransaksiAmountSalesReturn(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double))
        Dim DV01 As DataRowView
        Dim DV02 As DataRowView
        Try
            If TBLBal_SalesAutomaticSR.Rows.Count > 0 Then
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                For Each RowBal_SalesAutomaticSR In TBLBal_SalesAutomaticSR.Rows
                    TBL_SR.DefaultView.RowFilter = Nothing
                    TBL_SR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesAutomaticSR.Primary_Invoice.ToString
                    If TBL_SR.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_SalesAutomaticSR.Primary_Invoice) = False Then
                            ListPrimary.Add(RowBal_SalesAutomaticSR.Primary_Invoice)
                            Row_SR = TBL_SR.DefaultView(0).Row
                            'Input Transaction Amount
                            Row_SR.Transaction_Amount += RowBal_SalesAutomaticSR.balance_original + RowBal_SalesAutomaticSR.received_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_SR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_SR.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                Row_SR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If
                            'Input Baris Baru
                            If RowBal_SalesAutomaticSR.Primary_Currency <> 1 Then
                                If Row_SR.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    Row_SRDummy = TBL_SR.NewRow
                                    For xy As Integer = 0 To TBL_SR.Columns.Count - 1
                                        Row_SRDummy.Item(xy) = Row_SR.Item(xy)
                                    Next
                                    TBL_SR.Rows.Add(Row_SRDummy)
                                    'Currency Code-Nya IDR
                                    Row_SRDummy = TBL_SR.Rows(TBL_SR.Rows.Count - 1)
                                    Row_SRDummy.Primary_Currency = 1
                                    Row_SRDummy.Transaction_Amount = 0
                                    Row_SRDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_SalesAutomaticSR.Primary_Invoice) = True Then
                            If RowBal_SalesAutomaticSR.Primary_Currency = 1 Then
                                Row_SR = TBL_SR.DefaultView(0).Row
                                If Row_SR.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    Row_SR = TBL_SR.Rows(TBL_SR.Rows.Count - 1)
                                    Row_SR.Transaction_Amount += RowBal_SalesAutomaticSR.balance_original + RowBal_SalesAutomaticSR.received_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_SR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_SR.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        Row_SR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTransaksiAmountHistoricalSales()
        Dim DV01 As Data.DataRowView
        Try
            If TBL_HisSales.Rows.Count > 0 Then
                For Each Row_HisSales In TBL_HisSales.Rows
                    TBLBal_SalesManual.DefaultView.RowFilter = Nothing
                    TBLBal_SalesManual.DefaultView.RowFilter = "[Primary Invoice] = " & Row_HisSales.Primary_Transaksi.ToString & " AND [Primary Customer] = " & Row_HisSales.Primary_Customer.ToString & " AND [Primary Currency] = " & Row_HisSales.Primary_Currency.ToString
                    If TBLBal_SalesManual.DefaultView.Count > 0 Then
                        For Each DV01 In TBLBal_SalesManual.DefaultView
                            RowBal_SalesManual = DV01.Row
                            Row_HisSales.Transaction_Amount += RowBal_SalesManual.balance_original + RowBal_SalesManual.received_original
                        Next
                    End If
                    TBL_Pay.DefaultView.RowFilter = Nothing
                    TBL_Pay.DefaultView.Sort = Nothing
                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisSales.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisSales.Primary_Currency.ToString
                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                    If TBL_Pay.DefaultView.Count > 0 Then
                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                        Row_HisSales.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTransaksiAmountHistoricalSalesReturn()
        Dim DV01 As Data.DataRowView
        Try
            If TBL_HisSR.Rows.Count > 0 Then
                For Each Row_HisSR In TBL_HisSR.Rows
                    TBLBal_SalesManualSR.DefaultView.RowFilter = Nothing
                    TBLBal_SalesManualSR.DefaultView.RowFilter = "[Primary Invoice] = " & Row_HisSR.Primary_Transaksi.ToString & " AND [Primary Customer] = " & Row_HisSR.Primary_Customer.ToString & " AND [Primary Currency] = " & Row_HisSR.Primary_Currency.ToString
                    If TBLBal_SalesManualSR.DefaultView.Count > 0 Then
                        For Each DV01 In TBLBal_SalesManualSR.DefaultView
                            RowBal_SalesManualSR = DV01.Row
                            Row_HisSR.Transaction_Amount += RowBal_SalesManualSR.balance_original + RowBal_SalesManualSR.received_original
                        Next
                    End If
                    TBL_Pay.DefaultView.RowFilter = Nothing
                    TBL_Pay.DefaultView.Sort = Nothing
                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisSR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisSR.Primary_Currency.ToString
                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                    If TBL_Pay.DefaultView.Count > 0 Then
                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                        Row_HisSR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub GetTransaksiAmountCNoteSales(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double), ByVal PrimaryMin As Double, ByVal PrimaryMax As Double)
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView
        Dim DTAInvX As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim tblInvX As New DseLaporanARPaid.tbltrans_salesinvoiceDataTable
        Dim rowInvX As DseLaporanARPaid.tbltrans_salesinvoiceRow
        Dim DTAInvItemX As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim tblInvItemX As New DseLaporanARPaid.tbltrans_salesinvoice_itemDataTable
        Dim rowInvItemX As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
        Dim DTA_SOX As New DseLaporanARPaidTableAdapters.tbltrans_soTableAdapter
        Dim TBL_SOX As New DseLaporanARPaid.tbltrans_soDataTable
        Dim Row_SOX As DseLaporanARPaid.tbltrans_soRow
        Try
            'Get Transaction Amount Invoice
            If Me.TBLBal_SalesCNote.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                DTA_SOX.FillByCDNote(TBL_SOX)
                DTAInvX.FillByCDNote(tblInvX, PrimaryMin, PrimaryMax)
                DTAInvItemX.FillByCDNote(tblInvItemX, PrimaryMin, PrimaryMax)

               
                For Each RowBal_SalesCNote In TBLBal_SalesCNote.Rows
                    TBLCNote.DefaultView.RowFilter = Nothing
                    TBLCNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesCNote.Primary_Invoice.ToString
                    If TBLCNote.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_SalesCNote.Primary_Invoice) = False Then
                            ListPrimary.Add(RowBal_SalesCNote.Primary_Invoice)
                            RowCNote = TBLCNote.DefaultView(0).Row
                            'Input Transaction Amount
                            RowCNote.Transaction_Amount += RowBal_SalesCNote.balance_original + RowBal_SalesCNote.received_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowCNote.Primary_Transaksi.ToString & _
                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                            " AND [Primary Currency] = " & RowCNote.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                RowCNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If

                            If RowCNote.primary_invoice <> -1 Then
                                'Input Tax Currency
                                tblInvX.DefaultView.RowFilter = Nothing
                                tblInvX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCNote.primary_invoice.ToString
                                If tblInvX.DefaultView.Count > 0 Then
                                    rowInvX = tblInvX.Rows(0)
                                    RowCNote.Tax_Currency = rowInvX.Tax_Currency
                                End If
                                'Input Cust PO No
                                tblInvItemX.DefaultView.RowFilter = Nothing
                                tblInvItemX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCNote.primary_invoice.ToString
                                If tblInvItemX.DefaultView.Count > 0 Then
                                    TBL_SOX.DefaultView.RowFilter = Nothing
                                    TBL_SOX.DefaultView.Sort = Nothing
                                    For Each DV01 In tblInvItemX.DefaultView
                                        rowInvItemX = DV01.Row
                                        TBL_SOX.DefaultView.RowFilter = "[Primary SO] = " & rowInvItemX.Primary_SO.ToString
                                        TBL_SOX.DefaultView.Sort = "[Cust PO No]"
                                        If TBL_SOX.DefaultView.Count > 0 Then
                                            For Each DV02 In TBL_SOX.DefaultView
                                                Row_SOX = DV02.Row
                                                If String.IsNullOrEmpty(Row_SOX.Cust_PO_No) = False Then
                                                    If ListPrimarySOPONumber.Contains(Row_SOX.Primary_SO) = False Then
                                                        ListPrimarySOPONumber.Add(Row_SOX.Primary_SO)
                                                        If String.IsNullOrEmpty(RowCNote.Cust_PO_Number) = False Then
                                                            RowCNote.Cust_PO_Number += ", " + Row_SOX.Cust_PO_No.ToString
                                                        Else
                                                            RowCNote.Cust_PO_Number = Row_SOX.Cust_PO_No.ToString
                                                        End If
                                                    End If
                                                Else
                                                    If String.IsNullOrEmpty(RowCNote.Cust_PO_Number) = False Then
                                                        RowCNote.Cust_PO_Number += ", " + Row_SOX.No_Transaksi.ToString
                                                    Else
                                                        RowCNote.Cust_PO_Number = Row_SOX.No_Transaksi.ToString
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Else
                                RowCNote.Cust_PO_Number = "-"
                            End If


                            'Input Baris Baru
                            If RowBal_SalesCNote.Primary_Currency <> 1 Then
                                If RowCNote.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    RowCNoteDummy = TBLCNote.NewRow
                                    For xy As Integer = 0 To TBLCNote.Columns.Count - 1
                                        RowCNoteDummy.Item(xy) = RowCNote.Item(xy)
                                    Next
                                    TBLCNote.Rows.Add(RowCNoteDummy)
                                    'Currency Code-Nya IDR
                                    RowCNoteDummy = TBLCNote.Rows(TBLCNote.Rows.Count - 1)
                                    RowCNoteDummy.Primary_Currency = 1
                                    RowCNoteDummy.Transaction_Amount = 0
                                    RowCNoteDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_SalesCNote.Primary_Invoice) = True Then
                            If RowBal_SalesCNote.Primary_Currency = 1 Then
                                RowCNote = TBLCNote.DefaultView(0).Row
                                If RowCNote.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    RowCNote = TBLCNote.Rows(TBLCNote.Rows.Count - 1)
                                    RowCNote.Transaction_Amount += RowBal_SalesCNote.balance_original + RowBal_SalesCNote.received_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowCNote.Primary_Transaksi.ToString & _
                                                                    " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                                    " AND [Primary Currency] = " & RowCNote.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        RowCNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            'Dim Ina As New Formerna
            'Ina.DataGridView1.DataSource = Nothing
            'Ina.DataGridView2.DataSource = Nothing
            'Ina.DataGridView3.DataSource = Nothing

            'TBLCNote.DefaultView.RowFilter = Nothing
            'TBLBal_SalesCNote.DefaultView.RowFilter = Nothing
            'TBL_Pay.DefaultView.RowFilter = Nothing

            'Ina.DataGridView1.DataSource = TBLCNote
            'Ina.DataGridView2.DataSource = TBLBal_SalesCNote
            'Ina.DataGridView3.DataSource = TBL_Pay
            'Ina.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvX.Dispose()
            tblInvX.Dispose()
            DTAInvItemX.Dispose()
            tblInvItemX.Dispose()
            DTA_SOX.Dispose()
            TBL_SOX.Dispose()
        End Try
    End Sub

    Private Sub GetTransaksiAmountDNoteSales(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double), ByVal PrimaryMin As Double, ByVal PrimaryMax As Double)
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView
        Dim DTAInvX As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim tblInvX As New DseLaporanARPaid.tbltrans_salesinvoiceDataTable
        Dim rowInvX As DseLaporanARPaid.tbltrans_salesinvoiceRow
        Dim DTAInvItemX As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim tblInvItemX As New DseLaporanARPaid.tbltrans_salesinvoice_itemDataTable
        Dim rowInvItemX As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
        Dim DTA_SOX As New DseLaporanARPaidTableAdapters.tbltrans_soTableAdapter
        Dim TBL_SOX As New DseLaporanARPaid.tbltrans_soDataTable
        Dim Row_SOX As DseLaporanARPaid.tbltrans_soRow
        Try
            'Get Transaction Amount Invoice
            If Me.TBLBal_SalesDNote.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                DTA_SOX.FillByCDNote(TBL_SOX)
                DTAInvX.FillByCDNote(tblInvX, PrimaryMin, PrimaryMax)
                DTAInvItemX.FillByCDNote(tblInvItemX, PrimaryMin, PrimaryMax)

                For Each RowBal_SalesDNote In TBLBal_SalesDNote.Rows
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesDNote.Primary_Invoice.ToString
                    If TBLDNote.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_SalesDNote.Primary_Invoice) = False Then
                            ListPrimary.Add(RowBal_SalesDNote.Primary_Invoice)
                            RowDNote = TBLDNote.DefaultView(0).Row
                            'Input Transaction Amount
                            RowDNote.Transaction_Amount += RowBal_SalesDNote.balance_original + RowBal_SalesDNote.received_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                            " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If

                            If RowDNote.primary_invoice <> -1 Then
                                DTAInvX.FillByCDNote(tblInvX, PrimaryMin, PrimaryMax)
                                'Input Tax Currency
                                tblInvX.DefaultView.RowFilter = Nothing
                                tblInvX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
                                If tblInvX.DefaultView.Count > 0 Then
                                    rowInvX = tblInvX.Rows(0)
                                    RowDNote.Tax_Currency = rowInvX.Tax_Currency
                                End If
                                'Input Cust PO No
                                tblInvItemX.DefaultView.RowFilter = Nothing
                                tblInvItemX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
                                If tblInvItemX.DefaultView.Count > 0 Then
                                    TBL_SOX.DefaultView.RowFilter = Nothing
                                    TBL_SOX.DefaultView.Sort = Nothing
                                    For Each DV01 In tblInvItemX.DefaultView
                                        rowInvItemX = DV01.Row
                                        TBL_SOX.DefaultView.RowFilter = "[Primary SO] = " & rowInvItemX.Primary_SO.ToString
                                        TBL_SOX.DefaultView.Sort = "[Cust PO No]"
                                        If TBL_SOX.DefaultView.Count > 0 Then
                                            For Each DV02 In TBL_SOX.DefaultView
                                                Row_SOX = DV02.Row
                                                If String.IsNullOrEmpty(Row_SOX.Cust_PO_No) = False Then
                                                    If ListPrimarySOPONumber.Contains(Row_SOX.Primary_SO) = False Then
                                                        ListPrimarySOPONumber.Add(Row_SOX.Primary_SO)
                                                        If String.IsNullOrEmpty(RowDNote.Cust_PO_Number) = False Then
                                                            RowDNote.Cust_PO_Number += ", " + Row_SOX.Cust_PO_No.ToString
                                                        Else
                                                            RowDNote.Cust_PO_Number = Row_SOX.Cust_PO_No.ToString
                                                        End If
                                                    End If
                                                Else
                                                    If String.IsNullOrEmpty(RowDNote.Cust_PO_Number) = False Then
                                                        RowDNote.Cust_PO_Number += ", " + Row_SOX.No_Transaksi.ToString
                                                    Else
                                                        RowDNote.Cust_PO_Number = Row_SOX.No_Transaksi.ToString
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Else
                                RowDNote.Cust_PO_Number = "-"
                            End If


                            'Input Baris Baru
                            If RowBal_SalesDNote.Primary_Currency <> 1 Then
                                If RowDNote.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    RowDNoteDummy = TBLDNote.NewRow
                                    For xy As Integer = 0 To TBLDNote.Columns.Count - 1
                                        RowDNoteDummy.Item(xy) = RowDNote.Item(xy)
                                    Next
                                    TBLDNote.Rows.Add(RowDNoteDummy)
                                    'Currency Code-Nya IDR
                                    RowDNoteDummy = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
                                    RowDNoteDummy.Primary_Currency = 1
                                    RowDNoteDummy.Transaction_Amount = 0
                                    RowDNoteDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_SalesDNote.Primary_Invoice) = True Then
                            If RowBal_SalesDNote.Primary_Currency = 1 Then
                                RowDNote = TBLDNote.DefaultView(0).Row
                                If RowDNote.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    RowDNote = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
                                    RowDNote.Transaction_Amount += RowBal_SalesDNote.balance_original + RowBal_SalesDNote.received_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
                                                                    " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                                    " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvX.Dispose()
            tblInvX.Dispose()
            DTAInvItemX.Dispose()
            tblInvItemX.Dispose()
            DTA_SOX.Dispose()
            TBL_SOX.Dispose()
        End Try
    End Sub


    Private Sub GetTransaksiAmountPurchase(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double))
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView

        Try
            ListPrimarySOPONumber.Clear()
            'Get Transaction Amount Purchase
            If TBLBal_PurcAutomatic.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                For Each RowBal_PurcAutomatic In TBLBal_PurcAutomatic.Rows
                    TBL_AOP.DefaultView.RowFilter = Nothing
                    TBL_AOP.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcAutomatic.Primary_AOP.ToString
                    If TBL_AOP.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_PurcAutomatic.Primary_AOP) = False Then
                            ListPrimary.Add(RowBal_PurcAutomatic.Primary_AOP)
                            Row_AOP = TBL_AOP.DefaultView(0).Row
                            'Input Transaction Amount
                            Row_AOP.Transaction_Amount += RowBal_PurcAutomatic.balance_original + RowBal_PurcAutomatic.paid_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                Row_AOP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If
                            'Cust PO No
                            TBL_AOP_Item.DefaultView.RowFilter = Nothing
                            TBL_AOP_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & Row_AOP.Primary_Transaksi.ToString
                            If TBL_AOP_Item.DefaultView.Count > 0 Then
                                TBL_PO.DefaultView.RowFilter = Nothing
                                TBL_PO.DefaultView.Sort = Nothing
                                For Each DV01 In TBL_AOP_Item.DefaultView
                                    Row_AOP_Item = DV01.Row
                                    TBL_PO.DefaultView.RowFilter = "[Primary PO] = " & Row_AOP_Item.Primary_PO.ToString
                                    TBL_PO.DefaultView.Sort = "[No Transaksi]"
                                    If TBL_PO.DefaultView.Count > 0 Then
                                        For Each DV02 In TBL_PO.DefaultView
                                            Row_PO = DV02.Row
                                            If String.IsNullOrEmpty(Row_PO.No_Transaksi) = False Then
                                                If ListPrimarySOPONumber.Contains(Row_PO.Primary_PO) = False Then
                                                    ListPrimarySOPONumber.Add(Row_PO.Primary_PO)
                                                    If String.IsNullOrEmpty(Row_AOP.PO_Number) = False Then
                                                        Row_AOP.PO_Number += ", " + Row_PO.No_Transaksi.ToString
                                                    Else
                                                        Row_AOP.PO_Number = Row_PO.No_Transaksi.ToString
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            'Input Baris Baru
                            If RowBal_PurcAutomatic.Primary_Currency <> 1 Then
                                If Row_AOP.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    Row_AOPDummy = TBL_AOP.NewRow
                                    For xy As Integer = 0 To TBL_AOP.Columns.Count - 1
                                        Row_AOPDummy.Item(xy) = Row_AOP.Item(xy)
                                    Next
                                    TBL_AOP.Rows.Add(Row_AOPDummy)
                                    'Currency Code-Nya IDR
                                    Row_AOPDummy = TBL_AOP.Rows(TBL_AOP.Rows.Count - 1)
                                    Row_AOPDummy.Primary_Currency = 1
                                    Row_AOPDummy.Transaction_Amount = 0
                                    Row_AOPDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_PurcAutomatic.Primary_AOP) = True Then
                            If RowBal_PurcAutomatic.Primary_Currency = 1 Then
                                Row_AOP = TBL_AOP.DefaultView(0).Row
                                If Row_AOP.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    Row_AOP = TBL_AOP.Rows(TBL_AOP.Rows.Count - 1)
                                    Row_AOP.Transaction_Amount += RowBal_PurcAutomatic.balance_original + RowBal_PurcAutomatic.paid_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        Row_AOP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub GetTransaksiAmountPurchaseReturn(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double))
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView

        Try
            If TBLBal_PurcAutomaticPR.Rows.Count > 0 Then
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                For Each RowBal_PurcAutomaticPR In TBLBal_PurcAutomaticPR.Rows
                    TBL_AOP_PR.DefaultView.RowFilter = Nothing
                    TBL_AOP_PR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcAutomaticPR.Primary_AOP.ToString
                    If TBL_AOP_PR.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_PurcAutomaticPR.Primary_AOP) = False Then
                            ListPrimary.Add(RowBal_PurcAutomaticPR.Primary_AOP)
                            Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
                            'Input Transaction Amount
                            Row_AOP_PR.Transaction_Amount += RowBal_PurcAutomaticPR.balance_original + RowBal_PurcAutomaticPR.paid_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP_PR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP_PR.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                Row_AOP_PR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If
                            'Input Baris Baru
                            If RowBal_PurcAutomaticPR.Primary_Currency <> 1 Then
                                If Row_AOP_PR.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    Row_AOP_PRDummy = TBL_AOP_PR.NewRow
                                    For xy As Integer = 0 To TBL_AOP_PR.Columns.Count - 1
                                        Row_AOP_PRDummy.Item(xy) = Row_AOP_PR.Item(xy)
                                    Next
                                    TBL_AOP_PR.Rows.Add(Row_AOP_PRDummy)
                                    'Currency Code-Nya IDR
                                    Row_AOP_PRDummy = TBL_AOP_PR.Rows(TBL_AOP_PR.Rows.Count - 1)
                                    Row_AOP_PRDummy.Primary_Currency = 1
                                    Row_AOP_PRDummy.Transaction_Amount = 0
                                    Row_AOP_PRDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_PurcAutomaticPR.Primary_AOP) = True Then
                            If RowBal_PurcAutomaticPR.Primary_Currency = 1 Then
                                Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
                                If Row_AOP_PR.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    Row_AOP_PR = TBL_AOP_PR.Rows(TBL_AOP_PR.Rows.Count - 1)
                                    Row_AOP_PR.Transaction_Amount += RowBal_PurcAutomaticPR.balance_original + RowBal_PurcAutomaticPR.paid_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP_PR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP_PR.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        Row_AOP_PR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    'Private Sub GetTransaksiAmountHistoricalPurchase()
    '    Dim Dv01 As Data.DataRowView

    '    Try
    '        If TBL_HisPurc.Rows.Count > 0 Then
    '            For Each Row_HisPurc In TBL_HisPurc.Rows
    '                TBLBal_PurcManual.DefaultView.RowFilter = Nothing
    '                TBLBal_PurcManual.DefaultView.RowFilter = "[Primary AOP] = " & Row_HisPurc.Primary_Transaksi.ToString & " AND [Primary Supplier] = " & Row_HisPurc.Primary_Supplier.ToString & " AND [Primary Currency] = " & Row_HisPurc.Primary_Currency.ToString
    '                If TBLBal_PurcManual.DefaultView.Count > 0 Then
    '                    For Each Dv01 In TBLBal_PurcManual.DefaultView
    '                        RowBal_PurcManual = Dv01.Row
    '                        Row_HisPurc.Transaction_Amount += RowBal_PurcManual.balance_original + RowBal_PurcManual.paid_original
    '                    Next
    '                End If
    '                TBL_Pay.DefaultView.RowFilter = Nothing
    '                TBL_Pay.DefaultView.Sort = Nothing
    '                Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
    '                TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisPurc.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisPurc.Primary_Currency.ToString
    '                TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                If TBL_Pay.DefaultView.Count > 0 Then
    '                    Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                    Row_HisPurc.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try

    'End Sub

    'Private Sub GetTransaksiAmountHistoricalPurchaseReturn()
    '    Dim Dv01 As Data.DataRowView

    '    Try
    '        If TBL_HisPR.Rows.Count > 0 Then
    '            For Each Row_HisPR In TBL_HisPR.Rows
    '                TBLBal_PurcManualPR.DefaultView.RowFilter = Nothing
    '                TBLBal_PurcManualPR.DefaultView.RowFilter = "[Primary AOP] = " & Row_HisPR.Primary_Transaksi.ToString & " AND [Primary Supplier] = " & Row_HisPR.Primary_Supplier.ToString & " AND [Primary Currency] = " & Row_HisPR.Primary_Currency.ToString
    '                If TBLBal_PurcManualPR.DefaultView.Count > 0 Then
    '                    For Each Dv01 In TBLBal_PurcManualPR.DefaultView
    '                        RowBal_PurcManualPR = Dv01.Row
    '                        Row_HisPR.Transaction_Amount += RowBal_PurcManualPR.balance_original + RowBal_PurcManualPR.paid_original
    '                    Next
    '                End If
    '                TBL_Pay.DefaultView.RowFilter = Nothing
    '                TBL_Pay.DefaultView.Sort = Nothing
    '                Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '                TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisPR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisPR.Primary_Currency.ToString
    '                TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                If TBL_Pay.DefaultView.Count > 0 Then
    '                    Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                    Row_HisPR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub GetTransaksiAmountCNotePurchase(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double), ByVal PrimaryMin As Double, ByVal PrimaryMax As Double)
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView
        Dim DTAPurcX As New DseLaporanARPaidTableAdapters.tbltrans_aopTableAdapter
        Dim tblPurcX As New DseLaporanARPaid.tbltrans_aopDataTable
        Dim rowPurcX As DseLaporanARPaid.tbltrans_aopRow
        Dim DTAPurcItemX As New DseLaporanARPaidTableAdapters.tbltrans_aop_itemTableAdapter
        Dim tblPurcItemX As New DseLaporanARPaid.tbltrans_aop_itemDataTable
        Dim rowPurcItemX As DseLaporanARPaid.tbltrans_aop_itemRow
        Dim DTA_POX As New DseLaporanARPaidTableAdapters.tbltrans_poTableAdapter
        Dim TBL_POX As New DseLaporanARPaid.tbltrans_poDataTable
        Dim Row_POX As DseLaporanARPaid.tbltrans_poRow

        Try
            'Get Transaction Amount Purchase
            If Me.TBLBal_PurcCNote.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                DTAPurcItemX.FillBy_PrimaryMaxMin(tblPurcItemX, PrimaryMax, PrimaryMin)
                DTAPurcX.FillBy_CDNote(tblPurcX, PrimaryMin, PrimaryMax)
                DTA_POX.FillBy_CDNote(TBL_POX)

                For Each RowBal_PurcCNote In TBLBal_PurcCNote.Rows
                    TBLCNote.DefaultView.RowFilter = Nothing
                    TBLCNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcCNote.Primary_AOP.ToString
                    If TBLCNote.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_PurcCNote.Primary_AOP) = False Then
                            ListPrimary.Add(RowBal_PurcCNote.Primary_AOP)
                            RowCNote = TBLCNote.DefaultView(0).Row
                            'Input Transaction Amount
                            RowCNote.Transaction_Amount += RowBal_PurcCNote.balance_original + RowBal_PurcCNote.paid_original
                            'Input Paid Date
                            Me.TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowCNote.Primary_Transaksi.ToString & _
                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                            " AND [Primary Currency] = " & RowCNote.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                RowCNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If

                            If RowCNote.primary_invoice <> -1 Then
                                'Input Tax Currency
                                tblPurcX.DefaultView.RowFilter = Nothing
                                tblPurcX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCNote.primary_invoice.ToString
                                If tblPurcX.DefaultView.Count > 0 Then
                                    rowPurcX = tblPurcX.Rows(0)
                                    RowCNote.Tax_Currency = rowPurcX.Tax_Currency
                                End If

                                'Cust PO No
                                tblPurcItemX.DefaultView.RowFilter = Nothing
                                tblPurcItemX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCNote.primary_invoice.ToString
                                If tblPurcItemX.DefaultView.Count > 0 Then
                                    TBL_POX.DefaultView.RowFilter = Nothing
                                    TBL_POX.DefaultView.Sort = Nothing
                                    For Each DV01 In tblPurcItemX.DefaultView
                                        rowPurcItemX = DV01.Row
                                        TBL_POX.DefaultView.RowFilter = "[Primary PO] = " & rowPurcItemX.Primary_PO.ToString
                                        TBL_POX.DefaultView.Sort = "[No Transaksi]"
                                        If TBL_POX.DefaultView.Count > 0 Then
                                            For Each DV02 In TBL_POX.DefaultView
                                                Row_POX = DV02.Row
                                                If String.IsNullOrEmpty(Row_POX.No_Transaksi) = False Then
                                                    If ListPrimarySOPONumber.Contains(Row_POX.Primary_PO) = False Then
                                                        ListPrimarySOPONumber.Add(Row_POX.Primary_PO)
                                                        If String.IsNullOrEmpty(RowCNote.Cust_PO_Number) = False Then
                                                            RowCNote.Cust_PO_Number += ", " + Row_POX.No_Transaksi.ToString
                                                        Else
                                                            RowCNote.Cust_PO_Number = Row_POX.No_Transaksi.ToString
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                                'Input Baris Baru
                                If RowBal_PurcCNote.Primary_Currency <> 1 Then
                                    If RowCNote.Tax_Currency = 0 Then
                                        'Tambah Baris Buat IDR
                                        RowCNoteDummy = TBLCNote.NewRow
                                        For xy As Integer = 0 To TBLCNote.Columns.Count - 1
                                            RowCNoteDummy.Item(xy) = RowCNote.Item(xy)
                                        Next
                                        TBLCNote.Rows.Add(RowCNoteDummy)
                                        'Currency Code-Nya IDR
                                        RowCNoteDummy = TBLCNote.Rows(TBLCNote.Rows.Count - 1)
                                        RowCNoteDummy.Primary_Currency = 1
                                        RowCNoteDummy.Transaction_Amount = 0
                                        RowCNoteDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                End If
                            Else
                                RowCNote.Cust_PO_Number = "-"
                                RowCNote.Tax_Currency = 0
                            End If

                        ElseIf ListPrimary.Contains(RowBal_PurcCNote.Primary_AOP) = True Then
                            If RowBal_PurcCNote.Primary_Currency = 1 Then
                                RowCNote = TBLCNote.DefaultView(0).Row
                                If RowCNote.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    RowCNote = TBLCNote.Rows(TBLCNote.Rows.Count - 1)
                                    RowCNote.Transaction_Amount += RowBal_PurcCNote.balance_original + RowBal_PurcCNote.paid_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowCNote.Primary_Transaksi.ToString & _
                                                                    " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                                    " AND [Primary Currency] = " & RowCNote.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        RowCNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPurcX.Dispose()
            tblPurcX.Dispose()
            DTAPurcItemX.Dispose()
            tblPurcItemX.Dispose()
            DTA_POX.Dispose()
            TBL_POX.Dispose()
        End Try

    End Sub

    Private Sub GetTransaksiAmountDNotePurchase(ByVal ListPrimary As List(Of Double), ByVal ListPrimaryCurr As List(Of String), ByVal ListPrimarySOPONumber As List(Of Double), ByVal PrimaryMin As Double, ByVal PrimaryMax As Double)
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView
        Dim DTAInvX As New DseLaporanARPaidTableAdapters.tbltrans_aopTableAdapter
        Dim tblInvX As New DseLaporanARPaid.tbltrans_aopDataTable
        Dim rowInvX As DseLaporanARPaid.tbltrans_aopRow
        Dim DTAInvItemX As New DseLaporanARPaidTableAdapters.tbltrans_aop_itemTableAdapter
        Dim tblInvItemX As New DseLaporanARPaid.tbltrans_aop_itemDataTable
        Dim rowInvItemX As DseLaporanARPaid.tbltrans_aop_itemRow
        Dim DTA_SOX As New DseLaporanARPaidTableAdapters.tbltrans_poTableAdapter
        Dim TBL_SOX As New DseLaporanARPaid.tbltrans_poDataTable
        Dim Row_SOX As DseLaporanARPaid.tbltrans_poRow
        Try
            'Get Transaction Amount Invoice
            If Me.TBLBal_PurcDNote.Rows.Count > 0 Then
                ListPrimarySOPONumber.Clear()
                ListPrimaryCurr.Clear()
                ListPrimary.Clear()
                DTA_SOX.FillBy_CDNote(TBL_SOX)
                DTAInvItemX.FillBy_PrimaryMaxMin(tblInvItemX, PrimaryMin, PrimaryMax)

                For Each RowBal_PurcDNote In TBLBal_PurcDNote.Rows
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcDNote.Primary_AOP.ToString
                    If TBLDNote.DefaultView.Count > 0 Then
                        If ListPrimary.Contains(RowBal_PurcDNote.Primary_AOP) = False Then
                            ListPrimary.Add(RowBal_PurcDNote.Primary_AOP)
                            RowDNote = TBLDNote.DefaultView(0).Row
                            'Input Transaction Amount
                            RowDNote.Transaction_Amount += RowBal_PurcDNote.balance_original + RowBal_PurcDNote.paid_original
                            'Input Paid Date
                            TBL_Pay.DefaultView.RowFilter = Nothing
                            TBL_Pay.DefaultView.Sort = Nothing
                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                            " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                            If TBL_Pay.DefaultView.Count > 0 Then
                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                            End If

                            If RowDNote.primary_invoice <> -1 Then
                                DTAInvX.FillBy_CDNote(tblInvX, PrimaryMin, PrimaryMax)
                                'Input Tax Currency
                                tblInvX.DefaultView.RowFilter = Nothing
                                tblInvX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
                                If tblInvX.DefaultView.Count > 0 Then
                                    rowInvX = tblInvX.Rows(0)
                                    RowDNote.Tax_Currency = rowInvX.Tax_Currency
                                End If
                                'Input Cust PO No
                                tblInvItemX.DefaultView.RowFilter = Nothing
                                tblInvItemX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
                                If tblInvItemX.DefaultView.Count > 0 Then
                                    TBL_SOX.DefaultView.RowFilter = Nothing
                                    TBL_SOX.DefaultView.Sort = Nothing
                                    For Each DV01 In tblInvItemX.DefaultView
                                        rowInvItemX = DV01.Row
                                        TBL_SOX.DefaultView.RowFilter = "[Primary SO] = " & rowInvItemX.Primary_PO.ToString
                                        TBL_SOX.DefaultView.Sort = "[Cust PO No]"
                                        If TBL_SOX.DefaultView.Count > 0 Then
                                            For Each DV02 In TBL_SOX.DefaultView
                                                Row_SOX = DV02.Row
                                                If String.IsNullOrEmpty(Row_SOX.No_Transaksi) = False Then
                                                    If ListPrimarySOPONumber.Contains(Row_SOX.Primary_PO) = False Then
                                                        ListPrimarySOPONumber.Add(Row_SOX.Primary_PO)
                                                        If String.IsNullOrEmpty(RowDNote.Cust_PO_Number) = False Then
                                                            RowDNote.Cust_PO_Number += ", " + Row_SOX.No_Transaksi.ToString
                                                        Else
                                                            RowDNote.Cust_PO_Number = Row_SOX.No_Transaksi.ToString
                                                        End If
                                                    End If
                                                Else
                                                    If String.IsNullOrEmpty(RowDNote.Cust_PO_Number) = False Then
                                                        RowDNote.Cust_PO_Number += ", " + Row_SOX.No_Transaksi.ToString
                                                    Else
                                                        RowDNote.Cust_PO_Number = Row_SOX.No_Transaksi.ToString
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Else
                                RowDNote.Cust_PO_Number = "-"
                            End If


                            'Input Baris Baru
                            If RowBal_PurcDNote.Primary_Currency <> 1 Then
                                If RowDNote.Tax_Currency = 0 Then
                                    'Tambah Baris Buat IDR
                                    RowDNoteDummy = TBLDNote.NewRow
                                    For xy As Integer = 0 To TBLDNote.Columns.Count - 1
                                        RowDNoteDummy.Item(xy) = RowDNote.Item(xy)
                                    Next
                                    TBLDNote.Rows.Add(RowDNoteDummy)
                                    'Currency Code-Nya IDR
                                    RowDNoteDummy = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
                                    RowDNoteDummy.Primary_Currency = 1
                                    RowDNoteDummy.Transaction_Amount = 0
                                    RowDNoteDummy.Paid_Date = "1-Jan-1950"
                                End If
                            End If
                        ElseIf ListPrimary.Contains(RowBal_PurcDNote.Primary_AOP) = True Then
                            If RowBal_PurcDNote.Primary_Currency = 1 Then
                                RowDNote = TBLDNote.DefaultView(0).Row
                                If RowDNote.Tax_Currency = 0 Then
                                    'Input Pasangan USD yang IDR-Nya
                                    RowDNote = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
                                    RowDNote.Transaction_Amount += RowBal_PurcDNote.balance_original + RowBal_PurcDNote.paid_original
                                    'Input Paid Date
                                    TBL_Pay.DefaultView.RowFilter = Nothing
                                    TBL_Pay.DefaultView.Sort = Nothing
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
                                                                    " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
                                                                    " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                                    If TBL_Pay.DefaultView.Count > 0 Then
                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                                        RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvX.Dispose()
            tblInvX.Dispose()
            DTAInvItemX.Dispose()
            tblInvItemX.Dispose()
            DTA_SOX.Dispose()
            TBL_SOX.Dispose()
        End Try


        'Dim DV01 As Data.DataRowView
        'Dim DV02 As Data.DataRowView
        'Dim DTAPurcX As New DseLaporanARPaidTableAdapters.tbltrans_aopTableAdapter
        'Dim tblPurcX As New DseLaporanARPaid.tbltrans_aopDataTable
        'Dim rowPurcX As DseLaporanARPaid.tbltrans_aopRow
        'Dim DTAPurcItemX As New DseLaporanARPaidTableAdapters.tbltrans_aop_itemTableAdapter
        'Dim tblPurcItemX As New DseLaporanARPaid.tbltrans_aop_itemDataTable
        'Dim rowPurcItemX As DseLaporanARPaid.tbltrans_aop_itemRow
        'Dim DTA_POX As New DseLaporanARPaidTableAdapters.tbltrans_poTableAdapter
        'Dim TBL_POX As New DseLaporanARPaid.tbltrans_poDataTable
        'Dim Row_POX As DseLaporanARPaid.tbltrans_poRow

        'Try
        '    'Get Transaction Amount Purchase
        '    If Me.TBLBal_PurcDNote.Rows.Count > 0 Then
        '        ListPrimarySOPONumber.Clear()
        '        ListPrimaryCurr.Clear()
        '        ListPrimary.Clear()
        '        DTAPurcItemX.FillBy_PrimaryMaxMin(tblPurcItemX, PrimaryMax, PrimaryMin)
        '        DTAPurcX.FillBy_CDNote(tblPurcX, PrimaryMin, PrimaryMax)
        '        DTA_POX.FillBy_CDNote(TBL_POX)

        '        For Each RowBal_PurcDNote In TBLBal_PurcDNote.Rows
        '            TBLDNote.DefaultView.RowFilter = Nothing
        '            TBLDNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcDNote.Primary_AOP.ToString
        '            If TBLDNote.DefaultView.Count > 0 Then
        '                If ListPrimary.Contains(RowBal_PurcDNote.Primary_AOP) = False Then
        '                    ListPrimary.Add(RowBal_PurcDNote.Primary_AOP)
        '                    RowDNote = TBLDNote.DefaultView(0).Row
        '                    If Not IsDBNull(RowDNote.Primary_Transaksi) Then
        '                        'Input Transaction Amount
        '                        RowDNote.Transaction_Amount += RowBal_PurcDNote.balance_original + RowBal_PurcDNote.paid_original
        '                        'Input Paid Date
        '                        TBL_Pay.DefaultView.RowFilter = Nothing
        '                        TBL_Pay.DefaultView.Sort = Nothing
        '                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
        '                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
        '                                                        " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
        '                                                        " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
        '                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
        '                        If TBL_Pay.DefaultView.Count > 0 Then
        '                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
        '                            RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
        '                        End If

        '                        If RowDNote.primary_invoice <> -1 Then
        '                            'Input Tax Currency
        '                            tblPurcX.DefaultView.RowFilter = Nothing
        '                            tblPurcX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
        '                            If tblPurcX.DefaultView.Count > 0 Then
        '                                rowPurcX = tblPurcX.Rows(0)
        '                                RowDNote.Tax_Currency = rowPurcX.Tax_Currency
        '                            End If

        '                            'Cust PO No
        '                            tblPurcItemX.DefaultView.RowFilter = Nothing
        '                            tblPurcItemX.DefaultView.RowFilter = "[Primary Transaksi] = " & RowDNote.primary_invoice.ToString
        '                            If tblPurcItemX.DefaultView.Count > 0 Then
        '                                TBL_POX.DefaultView.RowFilter = Nothing
        '                                TBL_POX.DefaultView.Sort = Nothing
        '                                For Each DV01 In tblPurcItemX.DefaultView
        '                                    rowPurcItemX = DV01.Row
        '                                    TBL_POX.DefaultView.RowFilter = "[Primary PO] = " & rowPurcItemX.Primary_PO.ToString
        '                                    TBL_POX.DefaultView.Sort = "[No Transaksi]"
        '                                    If TBL_POX.DefaultView.Count > 0 Then
        '                                        For Each DV02 In TBL_POX.DefaultView
        '                                            Row_POX = DV02.Row
        '                                            If String.IsNullOrEmpty(Row_POX.No_Transaksi) = False Then
        '                                                If ListPrimarySOPONumber.Contains(Row_POX.Primary_PO) = False Then
        '                                                    ListPrimarySOPONumber.Add(Row_POX.Primary_PO)
        '                                                    If String.IsNullOrEmpty(RowDNote.Cust_PO_Number) = False Then
        '                                                        RowDNote.Cust_PO_Number += ", " + Row_POX.No_Transaksi.ToString
        '                                                    Else
        '                                                        RowDNote.Cust_PO_Number = Row_POX.No_Transaksi.ToString
        '                                                    End If
        '                                                End If
        '                                            End If
        '                                        Next
        '                                    End If
        '                                Next
        '                            End If


        '                            'Input Baris Baru
        '                            If RowBal_PurcDNote.Primary_Currency <> 1 Then
        '                                If RowDNote.Tax_Currency = 0 Then
        '                                    'Tambah Baris Buat IDR
        '                                    RowDNoteDummy = TBLDNote.NewRow
        '                                    For xy As Integer = 0 To TBLDNote.Columns.Count - 1
        '                                        RowDNoteDummy.Item(xy) = RowDNote.Item(xy)
        '                                    Next
        '                                    TBLDNote.Rows.Add(RowDNoteDummy)
        '                                    'Currency Code-Nya IDR
        '                                    RowDNoteDummy = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
        '                                    RowDNoteDummy.Primary_Currency = 1
        '                                    RowDNoteDummy.Transaction_Amount = 0
        '                                    RowDNoteDummy.Paid_Date = "1-Jan-1950"
        '                                End If
        '                            End If

        '                        Else
        '                            RowDNote.Cust_PO_Number = "-"
        '                            RowDNote.Tax_Currency = 0
        '                        End If
        '                    End If


        '                ElseIf ListPrimary.Contains(RowBal_PurcDNote.Primary_AOP) = True Then
        '                    If RowBal_PurcDNote.Primary_Currency = 1 Then
        '                        RowDNote = TBLDNote.DefaultView(0).Row
        '                        If RowDNote.Tax_Currency = 0 Then
        '                            'Input Pasangan USD yang IDR-Nya
        '                            RowDNote = TBLDNote.Rows(TBLDNote.Rows.Count - 1)
        '                            RowDNote.Transaction_Amount += RowBal_PurcDNote.balance_original + RowBal_PurcDNote.paid_original
        '                            'Input Paid Date
        '                            TBL_Pay.DefaultView.RowFilter = Nothing
        '                            TBL_Pay.DefaultView.Sort = Nothing
        '                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
        '                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowDNote.Primary_Transaksi.ToString & _
        '                                                            " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & _
        '                                                            " AND [Primary Currency] = " & RowDNote.Primary_Currency.ToString
        '                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
        '                            If TBL_Pay.DefaultView.Count > 0 Then
        '                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
        '                                RowDNote.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
        '                            End If
        '                        End If
        '                    End If
        '                End If
        '            End If
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DTAPurcX.Dispose()
        '    tblPurcX.Dispose()
        '    DTAPurcItemX.Dispose()
        '    tblPurcItemX.Dispose()
        '    DTA_POX.Dispose()
        '    TBL_POX.Dispose()
        'End Try

    End Sub

    Private Sub GetPrimaryMinMaxAR()
        Try
            'Primary Min Max
            TBLInv.DefaultView.RowFilter = Nothing
            TBLInv.DefaultView.Sort = Nothing
            TBLInv.DefaultView.Sort = "[Primary Transaksi]"
            If TBLInv.DefaultView.Count > 0 Then
                RowInv = Nothing
                RowInv = TBLInv.DefaultView(0).Row
                Me.PrimaryMinInv = RowInv.Primary_Transaksi
                RowInv = Nothing
                RowInv = TBLInv.DefaultView(TBLInv.DefaultView.Count - 1).Row
                Me.PrimaryMaxInv = RowInv.Primary_Transaksi
            Else
                Me.PrimaryMinInv = -1
                Me.PrimaryMaxInv = -1
            End If

            TBL_SR.DefaultView.RowFilter = Nothing
            TBL_SR.DefaultView.Sort = Nothing
            TBL_SR.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_SR.DefaultView.Count > 0 Then
                Row_SR = Nothing
                Row_SR = TBL_SR.DefaultView(0).Row
                Me.PrimaryMinSR = Row_SR.Primary_Transaksi
                Row_SR = Nothing
                Row_SR = TBL_SR.DefaultView(TBL_SR.DefaultView.Count - 1).Row
                Me.PrimaryMaxSR = Row_SR.Primary_Transaksi
            Else
                Me.PrimaryMinSR = -1
                Me.PrimaryMaxSR = -1
            End If

            TBL_HisSales.DefaultView.RowFilter = Nothing
            TBL_HisSales.DefaultView.Sort = Nothing
            TBL_HisSales.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_HisSales.DefaultView.Count > 0 Then
                Row_HisSales = Nothing
                Row_HisSales = TBL_HisSales.DefaultView(0).Row
                Me.PrimaryMinHistoricalSales = Row_HisSales.Primary_Transaksi
                Row_HisSales = Nothing
                Row_HisSales = TBL_HisSales.DefaultView(TBL_HisSales.DefaultView.Count - 1).Row
                Me.PrimaryMaxHistoricalSales = Row_HisSales.Primary_Transaksi
            Else
                Me.PrimaryMinHistoricalSales = -1
                Me.PrimaryMaxHistoricalSales = -1
            End If

            TBL_HisSR.DefaultView.RowFilter = Nothing
            TBL_HisSR.DefaultView.Sort = Nothing
            TBL_HisSR.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_HisSR.DefaultView.Count > 0 Then
                Row_HisSR = Nothing
                Row_HisSR = TBL_HisSR.DefaultView(0).Row
                Me.PrimaryMinHistoricalSalesReturn = Row_HisSR.Primary_Transaksi
                Row_HisSR = Nothing
                Row_HisSR = TBL_HisSR.DefaultView(TBL_HisSR.DefaultView.Count - 1).Row
                Me.PrimaryMaxHistoricalSalesReturn = Row_HisSR.Primary_Transaksi
            Else
                Me.PrimaryMinHistoricalSalesReturn = -1
                Me.PrimaryMaxHistoricalSalesReturn = -1
            End If

            Me.TBLCNote.DefaultView.RowFilter = Nothing
            TBLCNote.DefaultView.Sort = Nothing
            TBLCNote.DefaultView.Sort = "[Primary Transaksi]"
            If TBLCNote.DefaultView.Count > 0 Then
                RowCNote = Nothing
                RowCNote = TBLCNote.DefaultView(0).Row
                Me.PrimaryMinCreditNote = RowCNote.Primary_Transaksi
                RowCNote = Nothing
                RowCNote = TBLCNote.DefaultView(TBLCNote.DefaultView.Count - 1).Row
                Me.PrimaryMaxCreditNote = RowCNote.Primary_Transaksi
            Else
                Me.PrimaryMinCreditNote = -1
                Me.PrimaryMaxCreditNote = -1
            End If

            Me.TBLDNote.DefaultView.RowFilter = Nothing
            TBLDNote.DefaultView.Sort = Nothing
            TBLDNote.DefaultView.Sort = "[Primary Transaksi]"
            If TBLDNote.DefaultView.Count > 0 Then
                RowDNote = Nothing
                RowDNote = TBLDNote.DefaultView(0).Row
                Me.PrimaryMinDebitNote = RowDNote.Primary_Transaksi
                RowDNote = Nothing
                RowDNote = TBLDNote.DefaultView(TBLDNote.DefaultView.Count - 1).Row
                Me.PrimaryMaxDebitNote = RowDNote.Primary_Transaksi
            Else
                Me.PrimaryMinCreditNote = -1
                Me.PrimaryMaxCreditNote = -1
            End If



            'Me.TBLDNote.DefaultView.RowFilter = Nothing
            'TBLDNote.DefaultView.Sort = Nothing
            'TBLDNote.DefaultView.Sort = "[Primary Transaksi]"
            'If TBLDNote.DefaultView.Count > 1 Then
            '    Me.PrimaryMinDebitNote = -1
            '    Me.PrimaryMaxDebitNote = -1
            '    RowDNote = Nothing
            '    RowDNote = TBLDNote.DefaultView(0).Row
            '    If IsNumeric(RowDNote.Primary_Transaksi) Then
            '        Me.PrimaryMinDebitNote = RowDNote.Primary_Transaksi
            '        Me.PrimaryMaxDebitNote = RowDNote.Primary_Transaksi
            '    End If

            '    If IsNumeric(TBLDNote.DefaultView(TBLDNote.DefaultView.Count - 2).Item(0)) Then
            '        RowDNote = Nothing
            '        RowDNote = TBLDNote.DefaultView(TBLDNote.DefaultView.Count - 2).Row
            '        Me.PrimaryMaxDebitNote = RowDNote.Primary_Transaksi
            '    End If
            'Else
            '    Me.PrimaryMinDebitNote = -1
            '    Me.PrimaryMaxDebitNote = -1
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetPrimaryMinMaxAP()

        Try
            TBL_AOP.DefaultView.RowFilter = Nothing
            TBL_AOP.DefaultView.Sort = Nothing
            TBL_AOP.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_AOP.DefaultView.Count > 0 Then
                Row_AOP = TBL_AOP.DefaultView(0).Row
                Me.PrimaryMinInv = Row_AOP.Primary_Transaksi
                Row_AOP = TBL_AOP.DefaultView(TBL_AOP.DefaultView.Count - 1).Row
                Me.PrimaryMaxInv = Row_AOP.Primary_Transaksi
            Else
                Me.PrimaryMinInv = -1
                Me.PrimaryMaxInv = -1
            End If

            TBL_AOP_PR.DefaultView.RowFilter = Nothing
            TBL_AOP_PR.DefaultView.Sort = Nothing
            TBL_AOP_PR.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_AOP_PR.DefaultView.Count > 0 Then
                Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
                Me.PrimaryMinSR = Row_AOP_PR.Primary_Transaksi
                Row_AOP_PR = TBL_AOP_PR.DefaultView(TBL_AOP_PR.DefaultView.Count - 1).Row
                Me.PrimaryMaxSR = Row_AOP_PR.Primary_Transaksi
            Else
                Me.PrimaryMinSR = -1
                Me.PrimaryMaxSR = -1
            End If

            TBL_HisPurc.DefaultView.RowFilter = Nothing
            TBL_HisPurc.DefaultView.Sort = Nothing
            TBL_HisPurc.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_HisPurc.DefaultView.Count > 0 Then
                Row_HisPurc = TBL_HisPurc.DefaultView(0).Row
                Me.PrimaryMinHistoricalSales = Row_HisPurc.Primary_Transaksi
                Row_HisPurc = TBL_HisPurc.DefaultView(TBL_HisPurc.DefaultView.Count - 1).Row
                Me.PrimaryMaxHistoricalSales = Row_HisPurc.Primary_Transaksi
            Else
                Me.PrimaryMinHistoricalSales = -1
                Me.PrimaryMaxHistoricalSales = -1
            End If

            TBL_HisPR.DefaultView.RowFilter = Nothing
            TBL_HisPR.DefaultView.Sort = Nothing
            TBL_HisPR.DefaultView.Sort = "[Primary Transaksi]"
            If TBL_HisPR.DefaultView.Count > 0 Then
                Row_HisPR = TBL_HisPR.DefaultView(0).Row
                Me.PrimaryMinHistoricalSalesReturn = Row_HisPR.Primary_Transaksi
                Row_HisPR = TBL_HisPR.DefaultView(TBL_HisPR.DefaultView.Count - 1).Row
                Me.PrimaryMaxHistoricalSalesReturn = Row_HisPR.Primary_Transaksi
            Else
                Me.PrimaryMinHistoricalSalesReturn = -1
                Me.PrimaryMaxHistoricalSalesReturn = -1
            End If

            Me.TBLCNote.DefaultView.RowFilter = Nothing
            TBLCNote.DefaultView.Sort = Nothing
            TBLCNote.DefaultView.Sort = "[Primary Transaksi]"
            If TBLCNote.DefaultView.Count > 1 Then
                RowCNote = TBLCNote.DefaultView(0).Row
                Me.PrimaryMinCreditNote = RowCNote.Primary_Transaksi
                RowCNote = TBLCNote.DefaultView(TBLCNote.DefaultView.Count - 2).Row
                Me.PrimaryMaxCreditNote = RowCNote.Primary_Transaksi
            Else
                Me.PrimaryMinCreditNote = -1
                Me.PrimaryMaxCreditNote = -1
            End If

            Me.TBLDNote.DefaultView.RowFilter = Nothing
            TBLDNote.DefaultView.Sort = Nothing
            TBLDNote.DefaultView.Sort = "[Primary Transaksi]"
            If TBLDNote.DefaultView.Count > 1 Then
                RowDNote = TBLDNote.DefaultView(0).Row
                Me.PrimaryMinDebitNote = RowDNote.Primary_Transaksi
                RowDNote = TBLDNote.DefaultView(TBLDNote.DefaultView.Count - 2).Row
                Me.PrimaryMaxDebitNote = RowDNote.Primary_Transaksi
            Else
                Me.PrimaryMinDebitNote = -1
                Me.PrimaryMaxDebitNote = -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub HapusDataTidakSesuaiKriteria(ByVal mAbaikanHistorical As Boolean, ByVal mKriteria As String, ByVal mFilter02 As String)

        Dim Ketemu As Boolean = False
        Dim ij As Integer = 0
        Dim DV01 As Data.DataRowView
        Dim DV02 As Data.DataRowView

        Try
            '***Hapus Data Yang Tidak Sesuai Kriteria FILTER***
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                '*****Kartu AR*****
                If mAbaikanHistorical = False Then
                    If Me.ApakahAbaikanKriteriaFilterData() = False Then
                        'TBL Invoice 
                        If TBLInv.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBLInv.Rows.Count - 1
                                RowInv = TBLInv.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> RowInv.Primary_Customer Then
                                            TBLInv.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> RowInv.primary_akun_ar Then
                                            TBLInv.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If
                       
                        'TBL Sales Return
                        If TBL_SR.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBL_SR.Rows.Count - 1
                                Row_SR = TBL_SR.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> Row_SR.Primary_Customer Then
                                            TBL_SR.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> Row_SR.primary_akun_ar Then
                                            TBL_SR.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If
                       
                        'TBL Historical Sales
                        If TBL_HisSales.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBL_HisSales.Rows.Count - 1
                                Row_HisSales = TBL_HisSales.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCurrency.Text <> "-1" Then
                                        If Me.txtPrimaryCurrency.Text <> Row_HisSales.Primary_Currency Then
                                            TBL_HisSales.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> Row_HisSales.Primary_Customer Then
                                            TBL_HisSales.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> Row_HisSales.primary_akun_ar Then
                                            TBL_HisSales.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If



                        'TBL Historical Sales Return
                        If TBL_HisSR.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBL_HisSR.Rows.Count - 1
                                Row_HisSR = TBL_HisSR.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCurrency.Text <> "-1" Then
                                        If Me.txtPrimaryCurrency.Text <> Row_HisSR.Primary_Currency Then
                                            TBL_HisSR.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> Row_HisSR.Primary_Customer Then
                                            TBL_HisSR.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> Row_HisSR.primary_akun_ar Then
                                            TBL_HisSR.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If
                      

                        'TBL Credit Note
                        If TBLCNote.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBLCNote.Rows.Count - 1
                                RowCNote = TBLCNote.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCurrency.Text <> "-1" Then
                                        If Me.txtPrimaryCurrency.Text <> RowCNote.Primary_Currency Then
                                            TBLCNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> RowCNote.Primary_Customer Then
                                            TBLCNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> RowCNote.primary_akun_ar Then
                                            TBLCNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If
                       

                        'TBL Debit Note
                        If TBLDNote.Rows.Count > 0 Then
                            ij = 0
                            While ij <= TBLDNote.Rows.Count - 1
                                RowDNote = TBLDNote.Rows(ij)
                                Ketemu = False
                                If Ketemu = False Then
                                    If Me.txtPrimaryCurrency.Text <> "-1" Then
                                        If Me.txtPrimaryCurrency.Text <> RowDNote.Primary_Currency Then
                                            TBLDNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryCust.Text <> "-1" Then
                                        If Me.txtPrimaryCust.Text <> RowDNote.Primary_Customer Then
                                            TBLDNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    If Me.txtPrimaryArAccount.Text <> "-1" Then
                                        If Me.txtPrimaryArAccount.Text <> RowDNote.primary_akun_ar Then
                                            TBLDNote.Rows.RemoveAt(ij)
                                            Ketemu = True
                                        End If
                                    End If
                                End If
                                If Ketemu = False Then
                                    ij += 1
                                End If
                            End While
                        End If
                    End If
                        
                    Me.GetPrimaryMinMax_SalesOrder(TBLInvItem)
                ElseIf mAbaikanHistorical = True Then
                    'TBL Invoice
                    TBLInvItem.DefaultView.RowFilter = Nothing
                    TBLInvItem.DefaultView.RowFilter = mFilter02
                    If TBLInvItem.DefaultView.Count > 0 Then
                        mKriteria = ""
                        For Each DV01 In TBLInvItem.DefaultView
                            RowInvItem = DV01.Row
                            If String.IsNullOrEmpty(mKriteria) = False Then
                                mKriteria += " AND [Primary Transaksi] <> " & RowInvItem.Primary_Transaksi.ToString
                            Else
                                mKriteria = "[Primary Transaksi] <> " & RowInvItem.Primary_Transaksi.ToString
                            End If
                        Next
                        TBLInv.DefaultView.RowFilter = Nothing
                        TBLInv.DefaultView.RowFilter = mKriteria
                        ij = 0
                        While ij <= TBLInv.DefaultView.Count - 1
                            DV02 = TBLInv.DefaultView(ij)
                            RowInv = DV02.Row
                            TBLInv.Rows.Remove(RowInv)
                        End While
                        ij = 0
                        While ij <= TBLInv.Rows.Count - 1
                            Ketemu = False
                            RowInv = TBLInv.Rows(ij)
                            If Ketemu = False Then
                                If Me.txtPrimaryCust.Text <> "-1" Then
                                    If Me.txtPrimaryCust.Text <> RowInv.Primary_Customer Then
                                        TBLInv.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> RowInv.primary_akun_ar Then
                                        TBLInv.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                    Else
                        TBLInv.Rows.Clear()
                    End If

                    'TBL Sales Return
                    TBL_SRItem.DefaultView.RowFilter = Nothing
                    TBL_SRItem.DefaultView.RowFilter = mFilter02
                    If TBL_SRItem.DefaultView.Count > 0 Then
                        mKriteria = ""
                        For Each DV01 In TBL_SRItem.DefaultView
                            Row_SRItem = DV01.Row
                            If String.IsNullOrEmpty(mKriteria) = False Then
                                mKriteria += " AND [Primary Transaksi] <> " & Row_SRItem.Primary_Transaksi.ToString
                            Else
                                mKriteria = "[Primary Transaksi] <> " & Row_SRItem.Primary_Transaksi.ToString
                            End If
                        Next
                        TBL_SR.DefaultView.RowFilter = Nothing
                        TBL_SR.DefaultView.RowFilter = mKriteria
                        ij = 0
                        While ij <= TBL_SR.DefaultView.Count - 1
                            DV02 = TBL_SR.DefaultView(ij)
                            Row_SR = DV02.Row
                            TBL_SR.Rows.Remove(Row_SR)
                        End While
                        ij = 0
                        While ij <= TBL_SR.Rows.Count - 1
                            Ketemu = False
                            Row_SR = TBL_SR.Rows(ij)
                            If Ketemu = False Then
                                If Me.txtPrimaryCust.Text <> "-1" Then
                                    If Me.txtPrimaryCust.Text <> Row_SR.Primary_Customer Then
                                        TBL_SR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_SR.primary_akun_ar Then
                                        TBL_SR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                    Else
                        TBL_SR.Rows.Clear()
                    End If
                    '****
                End If
                Me.GetPrimaryMinMax_SalesOrder(TBLInvItem)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                '*****Kartu AP*****
                If mAbaikanHistorical = False Then
                    If Me.ApakahAbaikanKriteriaFilterData() = False Then
                        'TBL AOP 
                        ij = 0
                        While ij <= TBL_AOP.Rows.Count - 1
                            Row_AOP = TBL_AOP.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_AOP.Primary_Supplier Then
                                        TBL_AOP.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_AOP.primary_akun_ap Then
                                        TBL_AOP.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                        'TBL Purchase Return
                        ij = 0
                        While ij <= TBL_AOP_PR.Rows.Count - 1
                            Row_AOP_PR = TBL_AOP_PR.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_AOP_PR.Primary_Supplier Then
                                        TBL_AOP_PR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_AOP_PR.primary_akun_ap Then
                                        TBL_AOP_PR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                        'TBL Historical Purchase
                        ij = 0
                        While ij <= TBL_HisPurc.Rows.Count - 1
                            Row_HisPurc = TBL_HisPurc.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.txtPrimaryCurrency.Text <> "-1" Then
                                    If Me.txtPrimaryCurrency.Text <> Row_HisPurc.Primary_Currency Then
                                        TBL_HisPurc.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_HisPurc.Primary_Supplier Then
                                        TBL_HisPurc.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_HisPurc.primary_akun_ap Then
                                        TBL_HisPurc.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                        'TBL Historical Purchase Return 
                        ij = 0
                        While ij <= TBL_HisPR.Rows.Count - 1
                            Row_HisPR = TBL_HisPR.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.txtPrimaryCurrency.Text <> "-1" Then
                                    If Me.txtPrimaryCurrency.Text <> Row_HisPR.Primary_Currency Then
                                        TBL_HisPR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_HisPR.Primary_Supplier Then
                                        TBL_HisPR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_HisPR.primary_akun_ap Then
                                        TBL_HisPR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While

                        'TBL Credit note
                        ij = 0
                        While ij <= TBLCNote.Rows.Count - 1
                            RowCNote = TBLCNote.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.txtPrimaryCurrency.Text <> "-1" Then
                                    If Me.txtPrimaryCurrency.Text <> RowCNote.Primary_Currency Then
                                        TBLCNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> RowCNote.Primary_Supplier Then
                                        TBLCNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> RowCNote.primary_akun_ap Then
                                        TBLCNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While

                        'TBL Debit note
                        ij = 0
                        While ij <= TBLDNote.Rows.Count - 1
                            RowDNote = TBLDNote.Rows(ij)
                            Ketemu = False
                            If Ketemu = False Then
                                If Me.txtPrimaryCurrency.Text <> "-1" Then
                                    If Me.txtPrimaryCurrency.Text <> RowDNote.Primary_Currency Then
                                        TBLDNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> RowDNote.Primary_Supplier Then
                                        TBLDNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> RowDNote.primary_akun_ap Then
                                        TBLDNote.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                        '****
                    End If
                    Me.GetPrimaryMinMax_PurchaseOrder(TBL_AOP_Item)
                ElseIf mAbaikanHistorical = True Then
                    'TBL AOP
                    TBL_AOP_Item.DefaultView.RowFilter = Nothing
                    TBL_AOP_Item.DefaultView.RowFilter = mFilter02
                    If TBL_AOP_Item.DefaultView.Count > 0 Then
                        mKriteria = ""
                        For Each DV01 In TBL_AOP_Item.DefaultView
                            Row_AOP_Item = DV01.Row
                            If String.IsNullOrEmpty(mKriteria) = False Then
                                mKriteria += " AND [Primary Transaksi] <> " & Row_AOP_Item.Primary_Transaksi.ToString
                            Else
                                mKriteria = "[Primary Transaksi] <> " & Row_AOP_Item.Primary_Transaksi.ToString
                            End If
                        Next
                        TBL_AOP.DefaultView.RowFilter = Nothing
                        TBL_AOP.DefaultView.RowFilter = mKriteria
                        ij = 0
                        While ij <= TBL_AOP.DefaultView.Count - 1
                            DV02 = TBL_AOP.DefaultView(ij)
                            Row_AOP = DV02.Row
                            TBL_AOP.Rows.Remove(Row_AOP)
                        End While
                        ij = 0
                        While ij <= TBL_AOP.Rows.Count - 1
                            Ketemu = False
                            Row_AOP = TBL_AOP.Rows(ij)
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_AOP.Primary_Supplier Then
                                        TBL_AOP.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_AOP.primary_akun_ap Then
                                        TBL_AOP.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                    Else
                        TBL_AOP.Rows.Clear()
                    End If
                    'TBL Purchase Return
                    TBL_AOP_PR_Item.DefaultView.RowFilter = Nothing
                    TBL_AOP_PR_Item.DefaultView.RowFilter = mFilter02
                    If TBL_AOP_PR_Item.DefaultView.Count > 0 Then
                        mKriteria = ""
                        For Each DV01 In TBL_AOP_PR_Item.DefaultView
                            row_AOP_PR_Item = DV01.Row
                            If String.IsNullOrEmpty(mKriteria) = False Then
                                mKriteria += " AND [Primary Transaksi] <> " & row_AOP_PR_Item.Primary_Transaksi.ToString
                            Else
                                mKriteria = "[Primary Transaksi] <> " & row_AOP_PR_Item.Primary_Transaksi.ToString
                            End If
                        Next
                        TBL_AOP_PR.DefaultView.RowFilter = Nothing
                        TBL_AOP_PR.DefaultView.RowFilter = mKriteria
                        ij = 0
                        While ij <= TBL_AOP_PR.DefaultView.Count - 1
                            DV02 = TBL_AOP_PR.DefaultView(ij)
                            Row_AOP_PR = DV02.Row
                            TBL_AOP_PR.Rows.Remove(Row_AOP_PR)
                        End While
                        ij = 0
                        While ij <= TBL_AOP_PR.Rows.Count - 1
                            Ketemu = False
                            Row_AOP_PR = TBL_AOP_PR.Rows(ij)
                            If Ketemu = False Then
                                If Me.TxtPrimarySupplier.Text <> "-1" Then
                                    If Me.TxtPrimarySupplier.Text <> Row_AOP_PR.Primary_Supplier Then
                                        TBL_AOP_PR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                If Me.txtPrimaryArAccount.Text <> "-1" Then
                                    If Me.txtPrimaryArAccount.Text <> Row_AOP_PR.primary_akun_ap Then
                                        TBL_AOP_PR.Rows.RemoveAt(ij)
                                        Ketemu = True
                                    End If
                                End If
                            End If
                            If Ketemu = False Then
                                ij += 1
                            End If
                        End While
                    Else
                        TBL_AOP_PR.Rows.Clear()
                    End If
                    '****
                End If
                Me.GetPrimaryMinMax_PurchaseOrder(TBL_AOP_Item)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim Report_AR_AP As New CrpLaporanCustomerPaidNSupplierPaid
        Dim DS As New DseLaporanARPaid
        Dim TBLCetak As DseLaporanARPaid.TblCetakDataTable = DS.TblCetak
        Dim RowCetak As DseLaporanARPaid.TblCetakRow
        Dim TBLCetakTotal As DseLaporanARPaid.TblTotalDataTable = DS.TblTotal
        Dim RowCetakTotal As DseLaporanARPaid.TblTotalRow
        Dim TBLDummy As New DseLaporanARPaid.TblCetakDataTable
        Dim RowDummy As DseLaporanARPaid.TblCetakRow
        Dim mFilter01 As String = ""
        Dim mFilter02 As String = ""
        Dim mKriteria As String = ""
        Dim mJudul As String = ""
        Dim mJudulKolom As String = ""
        Dim Ketemu As Boolean = False
        Dim DV01 As DataRowView
        Dim DV02 As DataRowView
        Dim ij As Integer = 0
        Dim mAbaikanHistorical As Boolean = False
        Dim ListPrimarySOPONumber As New List(Of Double)
        Dim ListPrimaryCurr As New List(Of String)
        Dim ListPrimary As New List(Of Double)
        Dim EndLicenseDate As Date = #1/1/1950#

        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpAsOf.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpAsOf.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
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

            'cek:
            '1. tanggal dari <= tgl sampai
            If Me.dtpFrom.Value.Date > Me.dtpFrom.Value.Date Then
                MsgBox("Invalid document date.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            '2. tgl as of >= tgl sampai
            If Me.DtpAsOf.Value.Date < Me.dtpTo.Value.Date Then
                MsgBox("'As of date' must be later than 'to' date, please input date again.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            '3. untuk akun harus tipe piutang dagang, jika kartunya ar
            '   untuk akun harus tipe hutang dagang, jika kartunya ap
            '   udah di setting di main menu
            '*****Connection string***** 
            DTABal_SalesAutomatic.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_SalesAutomaticSR.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_SalesManual.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_SalesManualSR.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_SalesDNote.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_SalesCNote.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcAutomatic.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcAutomaticPR.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcManual.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcManualPR.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcDNote.Connection.ConnectionString = Core_Module.GetConstring
            DTABal_PurcCNote.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Pay.Connection.ConnectionString = Core_Module.GetConstring
            DTADNote.Connection.ConnectionString = Core_Module.GetConstring
            DTACNote.Connection.ConnectionString = Core_Module.GetConstring
            DTASaldo_AR.Connection.ConnectionString = Core_Module.GetConstring
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTAInvItem.Connection.ConnectionString = Core_Module.GetConstring
            DTA_SR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_SRItem.Connection.ConnectionString = Core_Module.GetConstring
            DTA_HisSales.Connection.ConnectionString = Core_Module.GetConstring
            DTA_HisSR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_SO.Connection.ConnectionString = Core_Module.GetConstring
            DTASaldo_AP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP_PR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP_PR_Item.Connection.ConnectionString = Core_Module.GetConstring
            DTA_HisPurc.Connection.ConnectionString = Core_Module.GetConstring
            DTA_HisPR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PO.Connection.ConnectionString = Core_Module.GetConstring
            'aul add
            DTA_HisSalesDP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_HisPurchDP.Connection.ConnectionString = Core_Module.GetConstring

            'load data
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                mJudul = "Account Receivable Balance Report"
                mJudulKolom = "Customer"
                If Me.CboPayment.SelectedIndex = 0 Then
                    'Unfinish
                    DTASaldo_AR.Fill(TBLSaldo_AR, Me.TipeKartu, Me.DtpAsOf.Value.Date)
                ElseIf Me.CboPayment.SelectedIndex = 1 Then
                    'All data
                    DTASaldo_AR.FillBy_All(TBLSaldo_AR, Me.TipeKartu, Me.DtpAsOf.Value.Date)
                End If
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                mJudul = "Account Payable Balance Report"
                mJudulKolom = "Supplier"
                If Me.CboPayment.SelectedIndex = 0 Then
                    'Unfinish
                    DTASaldo_AP.Fill(TBLSaldo_AP, Me.TipeKartu, Me.DtpAsOf.Value.Date)
                ElseIf Me.CboPayment.SelectedIndex = 1 Then
                    'All data
                    DTASaldo_AP.FillBy_All(TBLSaldo_AP, Me.TipeKartu, Me.DtpAsOf.Value.Date)
                End If
            End If

            '1. sales/purchase invoice
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                DTAInv.Fill(TBLInv, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                DTAInvItem.Fill(TBLInvItem, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                DTA_AOP.Fill(TBL_AOP, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                DTA_AOP_Item.Fill(TBL_AOP_Item, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            End If
            '2. sales/purchase return
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                DTA_SR.Fill(TBL_SR, Me.TipeStockInSR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipeSR)
                DTA_SRItem.Fill(TBL_SRItem, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipeSR)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                DTA_AOP_PR.Fill(TBL_AOP_PR, Me.TipePR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                DTA_AOP_PR_Item.Fill(TBL_AOP_PR_Item, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipePR)
            End If
            '3. historical sales/purchase
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                DTA_HisSales.Fill(TBL_HisSales, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                DTA_HisPurc.Fill(TBL_HisPurc, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            End If
            '4. historical sales/purchase return
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                DTA_HisSR.Fill(TBL_HisSR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                DTA_HisPR.Fill(TBL_HisPR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            End If
            '5. credit note/debit note
            DTADNote.Fill(TBLDNote, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu, ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE)
            DTACNote.Fill(TBLCNote, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE)
            Dim PrimaryMinInvoice As Double = -1
            Dim PrimaryMaxInvoice As Double = -1
            'aul add
            '5.1. historical sales/purchase DP
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                DTA_HisSalesDP.Fill(TBL_HisSalesDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                DTA_HisPurchDP.Fill(TBL_HisPurchDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            End If


            '6. filter customer, ar/ap account, currency, code no, so number, supplier, so no pada tabel 1,2,3,4,5
            mAbaikanHistorical = Me.AbaikanHistorical(Me.TipeKartu)
            '***FILTER*** ????
            mFilter01 = Me.GetFilterHeader()
            mFilter02 = Me.GetFilterItem()

            '-----------------------------------------------------------------------------------------------------------------------------------------
            Me.HapusDataTidakSesuaiKriteria(mAbaikanHistorical, mKriteria, mFilter02)
            '-----------------------------------------------------------------------------------------------------------------------------------------
            '6.1. isikan tabel balance untuk masing2 tabel 1,2,3,4,5 cari nilai transaction amount
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Me.GetPrimaryMinMaxAR()
                DTA_SO.Fill(TBL_SO, Me.PrimaryMaxSO, Me.PrimaryMinSO)
                DTA_Pay.Fill(TBL_Pay, Me.TipeJurnalAR, Me.dtpFrom.Value.Date, Me.DtpAsOf.Value.Date)
                DTABal_SalesAutomatic.Fill(TBLBal_SalesAutomatic, Me.PrimaryMaxInv, Me.PrimaryMinInv, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic)
                DTABal_SalesAutomaticSR.Fill(TBLBal_SalesAutomaticSR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR)
                DTABal_SalesManual.Fill(TBLBal_SalesManual, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual)
                DTABal_SalesManualSR.Fill(TBLBal_SalesManualSR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR)
                Me.DTABal_SalesDNote.Fill(TBLBal_SalesDNote, Me.PrimaryMaxDebitNote, Me.PrimaryMinDebitNote, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note)
                Me.DTABal_SalesCNote.Fill(TBLBal_SalesCNote, Me.PrimaryMaxCreditNote, Me.PrimaryMinCreditNote, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note)

                ' MsgBox(PrimaryMinCreditNote & "," & PrimaryMaxDebitNote)


                'Get Transaction Amount Invoice
                GetTransaksiAmountInvoice(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber)

                'Get Transaction Amount Sales Return
                GetTransaksiAmountSalesReturn(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber)

                'Historical Sales Invoice dan Sales Return
                If mAbaikanHistorical = False Then
                    GetTransaksiAmountHistoricalSales()
                    GetTransaksiAmountHistoricalSalesReturn()
                    'aul add
                    GetTransaksiAmountHistoricalDP(ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice)
                End If

                PrimaryMinInvoice = -1
                PrimaryMaxInvoice = -1
                PrimaryMinInvoice = Me.GetPrimaryMinInvoice(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)
                PrimaryMaxInvoice = Me.GetPrimaryMaxInvoice(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)

                Me.GetTransaksiAmountCNoteSales(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber, PrimaryMinInvoice, PrimaryMaxInvoice)
                Me.GetTransaksiAmountDNoteSales(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber, PrimaryMinInvoice, PrimaryMaxInvoice)
                '***************
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                'Primary Min Max
                Me.GetPrimaryMinMaxAP()
                DTA_PO.Fill(TBL_PO, Me.PrimaryMinSO, Me.PrimaryMaxSO)
                DTA_Pay.Fill(TBL_Pay, Me.TipeJurnalAP, Me.dtpFrom.Value.Date, Me.DtpAsOf.Value.Date)
                DTABal_PurcAutomatic.Fill(TBLBal_PurcAutomatic, Me.PrimaryMaxInv, Me.PrimaryMinInv, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic)
                DTABal_PurcAutomaticPR.Fill(TBLBal_PurcAutomaticPR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR)
                DTABal_PurcManual.Fill(TBLBal_PurcManual, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual)
                DTABal_PurcManualPR.Fill(TBLBal_PurcManualPR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR)
                Me.DTABal_PurcDNote.Fill(TBLBal_PurcDNote, Me.PrimaryMaxDebitNote, Me.PrimaryMinDebitNote, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note)
                Me.DTABal_PurcCNote.Fill(TBLBal_PurcCNote, Me.PrimaryMaxCreditNote, Me.PrimaryMinCreditNote, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note)



                'Get Transaction Amount Purchase
                GetTransaksiAmountPurchase(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber)

                'Get Transaction Amount Purchase Return
                GetTransaksiAmountPurchaseReturn(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber)

                If mAbaikanHistorical = False Then
                    'GetTransaksiAmountHistoricalPurchase()
                    'GetTransaksiAmountHistoricalPurchaseReturn()
                    'aul add
                    GetTransaksiAmountHistoricalDP(ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP)
                End If
                PrimaryMinInvoice = -1
                PrimaryMaxInvoice = -1
                PrimaryMinInvoice = Me.GetPrimaryMinInvoice(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)
                PrimaryMaxInvoice = Me.GetPrimaryMaxInvoice(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)
                Me.GetTransaksiAmountCNotePurchase(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber, PrimaryMinInvoice, PrimaryMaxInvoice)
                Me.GetTransaksiAmountDNotePurchase(ListPrimary, ListPrimaryCurr, ListPrimarySOPONumber, PrimaryMinInvoice, PrimaryMaxInvoice)
                '***************
            End If



            'filter khusus currency 
            Me.FilterDataByPrimaryCurrency(TBLInv, TBL_SR, Me.TBLCNote, Me.TBLDNote, TBL_AOP, TBL_AOP_PR)
            ' MsgBox(TBLCNote)
            '6.2. isikan tabel jurnal item payment untuk dapet paid date (udah diatas)
            '6.3. isikan cust po number utnuk invoice (udah diatas)
            '7. tabel saldo/saldo1 (udah diatas)
            '8. tentukan nilai primary min dan max untuk tabel invoice, sales return, historical sales, dan historical sales return (Udah diatas)
            '***************
            '9. buang baris dalam tabel saldo yang tidak sesuai dengan primary min dan primary max untuk setiap data
            '10. untuk setiap data dalam tabel saldo, pindahkan ke dalam tabel cetak

            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If TBLSaldo_AR.Rows.Count > 0 Then
                    For Each RowSaldo_AR In TBLSaldo_AR.Rows
                        If RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                            TBLInv.DefaultView.RowFilter = Nothing
                            TBLInv.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
                            If TBLInv.DefaultView.Count > 0 Then
                                RowInv = TBLInv.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                RowDummy.Primary_Cust = RowInv.Primary_Customer
                                RowDummy.Customer_Name = RowInv.Customer_Name
                                RowDummy.Primary_Invoice = RowInv.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                RowDummy.Invoice_No = RowInv.No_Transaksi
                                RowDummy.Invoice_Date = RowInv.Tanggal_Transaksi
                                RowDummy._Cust__PO_No = RowInv.Cust_PO_Number
                                RowDummy.Payment_Term = RowInv.Payment_Term
                                RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
                                RowDummy.Transaction_Amount = RowInv.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    RowDummy.Paid_Date = RowInv.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                            TBL_SR.DefaultView.RowFilter = Nothing
                            TBL_SR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
                            If TBL_SR.DefaultView.Count > 0 Then
                                Row_SR = TBL_SR.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                RowDummy.Primary_Cust = Row_SR.Primary_Customer
                                RowDummy.Customer_Name = Row_SR.Customer_Name
                                RowDummy.Primary_Invoice = Row_SR.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                                RowDummy.Invoice_No = Row_SR.No_Transaksi
                                RowDummy.Invoice_Date = Row_SR.Tanggal_Transaksi
                                RowDummy._Cust__PO_No = ""
                                RowDummy.Payment_Term = Row_SR.Payment_Term
                                RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
                                RowDummy.Transaction_Amount = Row_SR.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    RowDummy.Paid_Date = Row_SR.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                            TBLDNote.DefaultView.RowFilter = Nothing
                            TBLDNote.DefaultView.Sort = Nothing
                            TBLDNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & _
                            " AND [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                            " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
                            TBLDNote.DefaultView.Sort = "[Primary Transaksi]"
                            If TBLDNote.DefaultView.Count > 0 Then
                                RowDNote = TBLDNote.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                RowDummy.Primary_Cust = RowDNote.Primary_Customer
                                RowDummy.Customer_Name = RowDNote.customer_debet_note
                                RowDummy.Primary_Invoice = RowDNote.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                                RowDummy.Invoice_No = RowDNote.no_transksi
                                RowDummy.Invoice_Date = RowDNote.tgl
                                RowDummy._Cust__PO_No = RowDNote.Cust_PO_Number
                                RowDummy.Payment_Term = DateDiff(DateInterval.Day, RowDNote.due_date, RowDNote.tgl, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
                                RowDummy.Transaction_Amount = RowDNote.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                If IsDate(RowDNote.Paid_Date) Then
                                    RowDummy.Paid_Date = RowDNote.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                            TBLCNote.DefaultView.RowFilter = Nothing
                            TBLCNote.DefaultView.Sort = Nothing
                            TBLCNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & _
                                                             " AND [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                                                             " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
                            TBLCNote.DefaultView.Sort = "[Primary Transaksi]"
                            If TBLCNote.DefaultView.Count > 0 Then
                                RowCNote = TBLCNote.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                RowDummy.Primary_Cust = RowSaldo_AR.primary_currency
                                RowDummy.Customer_Name = RowSaldo_AR.Customer_Name
                                RowDummy.Primary_Invoice = RowCNote.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                                RowDummy.Invoice_No = RowCNote.no_transksi
                                RowDummy.Invoice_Date = RowCNote.tgl
                                RowDummy._Cust__PO_No = RowCNote.Cust_PO_Number
                                RowDummy.Payment_Term = DateDiff(DateInterval.Day, RowCNote.due_date, RowCNote.tgl, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                '   RowDummy.Payment_Term = Format(RowCNote.due_date - RowCNote.tgl, "dd-MMM-yyyy")
                                RowDummy.Due_Date = Format(RowCNote.due_date, "dd-MMM-yyyy")
                                RowDummy.Transaction_Amount = RowCNote.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                If IsDate(RowCNote.Paid_Date) Then
                                    'Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                                    RowDummy.Paid_Date = RowCNote.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        End If
                        If mAbaikanHistorical = False Then
                            If RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
                                TBL_HisSales.DefaultView.RowFilter = Nothing
                                TBL_HisSales.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString
                                If TBL_HisSales.DefaultView.Count > 0 Then
                                    Row_HisSales = TBL_HisSales.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                    RowDummy.Primary_Cust = Row_HisSales.Primary_Customer
                                    RowDummy.Customer_Name = Row_HisSales.Customer_Name
                                    RowDummy.Primary_Invoice = Row_HisSales.Primary_Transaksi
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                                    RowDummy.Invoice_No = Row_HisSales.Invoice_Number
                                    RowDummy.Invoice_Date = Row_HisSales.Invoice_Date
                                    RowDummy._Cust__PO_No = ""

                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisSales.Invoice_Date, Row_HisSales.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisSales.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisSales.Transaction_Amount
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisSales.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                            ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                                TBL_HisSR.DefaultView.RowFilter = Nothing
                                TBL_HisSR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString
                                If TBL_HisSR.DefaultView.Count > 0 Then
                                    Row_HisSR = TBL_HisSR.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                    RowDummy.Primary_Cust = Row_HisSR.Primary_Customer
                                    RowDummy.Customer_Name = Row_HisSR.Customer_Name
                                    RowDummy.Primary_Invoice = Row_HisSR.Primary_Transaksi
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                                    RowDummy.Invoice_No = Row_HisSR.AOSR_Number
                                    RowDummy.Invoice_Date = Row_HisSR.AOSR_Date
                                    RowDummy._Cust__PO_No = ""

                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisSR.AOSR_Date, Row_HisSR.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisSR.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisSR.Transaction_Amount
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisSR.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                                'aul add
                            ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual Then
                                TBL_HisSalesDP.DefaultView.RowFilter = Nothing
                                TBL_HisSalesDP.DefaultView.RowFilter = "[Primary] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString
                                If TBL_HisSalesDP.DefaultView.Count > 0 Then
                                    Row_HisSalesDP = TBL_HisSalesDP.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
                                    RowDummy.Primary_Cust = Row_HisSalesDP.Primary_Pihak_Acuan
                                    RowDummy.Customer_Name = Row_HisSalesDP.Pihak_Acuan_Name
                                    RowDummy.Primary_Invoice = Row_HisSalesDP.Primary
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
                                    RowDummy.Invoice_No = Row_HisSalesDP.DP_No
                                    RowDummy.Invoice_Date = Row_HisSalesDP.DP_Date
                                    RowDummy._Cust__PO_No = ""

                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisSalesDP.DP_Date, Row_HisSalesDP.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisSalesDP.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisSalesDP.Transaction_Amount
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AR.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisSalesDP.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If TBLSaldo_AP.Rows.Count > 0 Then
                    For Each RowSaldo_AP In TBLSaldo_AP.Rows
                        If RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                            TBL_AOP.DefaultView.RowFilter = Nothing
                            TBL_AOP.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & _
                                                            " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
                            If TBL_AOP.DefaultView.Count > 0 Then
                                Row_AOP = TBL_AOP.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                RowDummy.Primary_Invoice = Row_AOP.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
                                If Me.chkUseSupplieInvoiceNo.Checked = True Then
                                    RowDummy.Invoice_No = Row_AOP.Supplier_Invoice_No
                                Else
                                    RowDummy.Invoice_No = Row_AOP.No_Transaksi
                                End If
                                RowDummy.Invoice_Date = Row_AOP.Tanggal_Transaksi
                                RowDummy._Cust__PO_No = Row_AOP.PO_Number
                                RowDummy.Payment_Term = Row_AOP.Payment_Term
                                RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
                                RowDummy.Transaction_Amount = Row_AOP.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    RowDummy.Paid_Date = Row_AOP.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                            TBL_AOP_PR.DefaultView.RowFilter = Nothing
                            TBL_AOP_PR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & _
                                                               " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
                            If TBL_AOP_PR.DefaultView.Count > 0 Then
                                Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                RowDummy.Primary_Invoice = Row_AOP_PR.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
                                RowDummy.Invoice_No = Row_AOP_PR.No_Transaksi
                                RowDummy.Invoice_Date = Row_AOP_PR.Tanggal_Transaksi
                                RowDummy._Cust__PO_No = ""
                                RowDummy.Payment_Term = Row_AOP_PR.Payment_Term
                                RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
                                RowDummy.Transaction_Amount = Row_AOP_PR.Transaction_Amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    RowDummy.Paid_Date = Row_AOP_PR.Paid_Date
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                            TBLDNote.DefaultView.RowFilter = Nothing
                            TBLDNote.DefaultView.Sort = Nothing
                            TBLDNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & _
                                                             " AND [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                                                             " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
                            TBLDNote.DefaultView.Sort = "[Primary Transaksi]"
                            If TBLDNote.DefaultView.Count > 0 Then
                                RowDNote = TBLDNote.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                RowDummy.Primary_Cust = RowSaldo_AP.primary_currency
                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                RowDummy.Primary_Invoice = RowDNote.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                                RowDummy.Invoice_No = RowDNote.no_transksi
                                RowDummy.Invoice_Date = RowDNote.tgl
                                RowDummy._Cust__PO_No = ""
                                RowDummy.Payment_Term = DateDiff(DateInterval.Day, RowDNote.due_date, RowDNote.tgl, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                RowDummy.Due_Date = Format(RowDNote.due_date, "dd-MMM-yyyy")
                                RowDummy.Transaction_Amount = RowDNote.amount * -1
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                                    RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AP.Primary_Transaksi_Asal, RowSaldo_AP.primary_currency)
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                            TBLCNote.DefaultView.RowFilter = Nothing
                            TBLCNote.DefaultView.Sort = Nothing
                            TBLCNote.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & _
                                                             " AND [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                                                             " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
                            TBLCNote.DefaultView.Sort = "[Primary Transaksi]"
                            If TBLCNote.DefaultView.Count > 0 Then
                                RowCNote = TBLCNote.DefaultView(0).Row
                                RowDummy = TBLDummy.NewRow
                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                RowDummy.Primary_Cust = RowSaldo_AP.primary_currency
                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                RowDummy.Primary_Invoice = RowCNote.Primary_Transaksi
                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                                RowDummy.Invoice_No = RowCNote.no_transksi
                                RowDummy.Invoice_Date = RowCNote.tgl
                                RowDummy._Cust__PO_No = ""
                                RowDummy.Payment_Term = DateDiff(DateInterval.Day, RowCNote.due_date, RowCNote.tgl, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                RowDummy.Due_Date = Format(RowCNote.due_date, "dd-MMM-yyyy")
                                RowDummy.Transaction_Amount = RowCNote.amount
                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                If RowDummy.Payment_Amount > 0 Then
                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                                    RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AP.Primary_Transaksi_Asal, RowSaldo_AP.primary_currency)
                                Else
                                    RowDummy.Paid_Date = "1-Jan-1950"
                                End If
                                TBLDummy.Rows.Add(RowDummy)
                            End If
                        End If
                        If mAbaikanHistorical = False Then
                            If RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Then
                                TBL_HisPurc.DefaultView.RowFilter = Nothing
                                TBL_HisPurc.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString
                                If TBL_HisPurc.DefaultView.Count > 0 Then
                                    Row_HisPurc = TBL_HisPurc.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                    RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
                                    RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                    RowDummy.Primary_Invoice = Row_HisPurc.Primary_Transaksi
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                                    RowDummy.Invoice_No = Row_HisPurc.AOP_Number
                                    RowDummy.Invoice_Date = Row_HisPurc.AOP_Date
                                    RowDummy._Cust__PO_No = ""

                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisPurc.AOP_Date, Row_HisPurc.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisPurc.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisPurc.Balance
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisPurc.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                            ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
                                TBL_HisPR.DefaultView.RowFilter = Nothing
                                TBL_HisPR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString
                                If TBL_HisPR.DefaultView.Count > 0 Then
                                    Row_HisPR = TBL_HisPR.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                    RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
                                    RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                    RowDummy.Primary_Invoice = Row_HisPR.Primary_Transaksi
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                                    RowDummy.Invoice_No = Row_HisPR.AOPR_Number
                                    RowDummy.Invoice_Date = Row_HisPR.AOPR_Date
                                    RowDummy._Cust__PO_No = ""


                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisPR.AOPR_Date, Row_HisPR.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisPR.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisPR.Balance
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisPR.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                                'aul add
                            ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual Then
                                TBL_HisPurchDP.DefaultView.RowFilter = Nothing
                                TBL_HisPurchDP.DefaultView.RowFilter = "[Primary] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString
                                If TBL_HisPurchDP.DefaultView.Count > 0 Then
                                    Row_HisPurchDP = TBL_HisPurchDP.DefaultView(0).Row
                                    RowDummy = TBLDummy.NewRow
                                    RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
                                    RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
                                    RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
                                    RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
                                    RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
                                    RowDummy.Primary_Invoice = Row_HisPurchDP.Primary
                                    RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                                    RowDummy.Invoice_No = Row_HisPurchDP.DP_No
                                    RowDummy.Invoice_Date = Row_HisPurchDP.DP_Date
                                    RowDummy._Cust__PO_No = ""

                                    RowDummy.Payment_Term = DateDiff(DateInterval.Day, Row_HisPurchDP.DP_Date, Row_HisPurchDP.Due_Date, FirstDayOfWeek.Monday, FirstWeekOfYear.Jan1)
                                    RowDummy.Due_Date = Row_HisPurchDP.Due_Date
                                    RowDummy.Transaction_Amount = Row_HisPurchDP.Transaction_Amount
                                    RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
                                    RowDummy.Outstanding = RowSaldo_AP.Balance_Original
                                    If RowDummy.Payment_Amount > 0 Then
                                        RowDummy.Paid_Date = Row_HisPurchDP.Paid_Date
                                    Else
                                        RowDummy.Paid_Date = "1-Jan-1950"
                                    End If
                                    TBLDummy.Rows.Add(RowDummy)
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            '***************
            '11. tampilkan ke dalam crv
            TBLDummy.DefaultView.Sort = Nothing
            TBLDummy.DefaultView.Sort = "[Currency Code], [Customer Name], [Invoice Date], [Invoice No]"
            If TBLDummy.DefaultView.Count > 0 Then
                For Each DV01 In TBLDummy.DefaultView
                    RowDummy = DV01.Row
                    If ListPrimaryCurr.Contains(RowDummy.Currency_Code) = False Then
                        ListPrimaryCurr.Add(RowDummy.Currency_Code)
                        RowCetakTotal = TBLCetakTotal.NewRow
                        RowCetakTotal.CurrCode = RowDummy.Currency_Code
                        RowCetakTotal.CurrName = RowDummy.Currency_Name
                        RowCetakTotal.TotalAmount = RowDummy.Transaction_Amount
                        RowCetakTotal.TotalPayment = RowDummy.Payment_Amount
                        RowCetakTotal.TotalOutstanding = RowDummy.Outstanding
                        RowCetakTotal.DefPrec = RowDummy.DefPrec_Currency
                        TBLCetakTotal.Rows.Add(RowCetakTotal)
                    Else
                        TBLCetakTotal.DefaultView.RowFilter = Nothing
                        TBLCetakTotal.DefaultView.RowFilter = "[CurrCode] = '" & RowDummy.Currency_Code.ToString & "'"
                        If TBLCetakTotal.DefaultView.Count > 0 Then
                            RowCetakTotal = TBLCetakTotal.DefaultView(0).Row
                            RowCetakTotal.TotalAmount += RowDummy.Transaction_Amount
                            RowCetakTotal.TotalPayment += RowDummy.Payment_Amount
                            RowCetakTotal.TotalOutstanding += RowDummy.Outstanding
                        End If
                    End If
                    RowCetak = TBLCetak.NewRow
                    For i As Integer = 0 To TBLCetak.Columns.Count - 1
                        RowCetak.Item(i) = RowDummy.Item(i)
                    Next
                    TBLCetak.Rows.Add(RowCetak)
                Next
            End If

            Report_AR_AP.SetDataSource(DS)
            Report_AR_AP.Subreports(0).SetDataSource(DS)
            Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamHeader.ParameterFieldName, mJudul)
            Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamGroup.ParameterFieldName, mJudulKolom)
            Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamStartDate.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd-MMM-yyyy"))
            Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamEndDate.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd-MMM-yyyy"))
            Me.crvAP.ReportSource = Report_AR_AP
            Me.crvAP.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimarySOPONumber = Nothing
            ListPrimaryCurr = Nothing
            ListPrimary = Nothing
            DTABal_SalesAutomatic.Dispose()
            TBLBal_SalesAutomatic.Dispose()
            DTABal_SalesAutomaticSR.Dispose()
            TBLBal_SalesAutomaticSR.Dispose()
            DTABal_SalesManual.Dispose()
            TBLBal_SalesManual.Dispose()
            DTABal_SalesManualSR.Dispose()
            TBLBal_SalesManualSR.Dispose()
            DTABal_SalesCNote.Dispose()
            TBLBal_SalesCNote.Dispose()
            DTABal_SalesDNote.Dispose()
            TBLBal_SalesDNote.Dispose()
            DTABal_PurcAutomatic.Dispose()
            TBLBal_PurcAutomatic.Dispose()
            DTABal_PurcAutomaticPR.Dispose()
            TBLBal_PurcAutomaticPR.Dispose()
            DTABal_PurcManual.Dispose()
            TBLBal_PurcManual.Dispose()
            DTABal_PurcManualPR.Dispose()
            TBLBal_PurcManualPR.Dispose()
            DTABal_PurcCNote.Dispose()
            TBLBal_PurcCNote.Dispose()
            DTABal_PurcDNote.Dispose()
            TBLBal_PurcDNote.Dispose()
            DS.Dispose()
            DTA_Pay.Dispose()
            TBL_Pay.Dispose()
            DTADNote.Dispose()
            TBLDNote.Dispose()
            DTACNote.Dispose()
            TBLCNote.Dispose()
            DTASaldo_AR.Dispose()
            TBLSaldo_AR.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTAInvItem.Dispose()
            TBLInvItem.Dispose()
            DTA_SR.Dispose()
            TBL_SR.Dispose()
            DTA_SRItem.Dispose()
            TBL_SRItem.Dispose()
            DTA_HisSales.Dispose()
            TBL_HisSales.Dispose()
            DTA_HisSR.Dispose()
            TBL_HisSR.Dispose()
            DTA_SO.Dispose()
            TBL_SO.Dispose()
            DTASaldo_AP.Dispose()
            TBLSaldo_AP.Dispose()
            DTA_AOP.Dispose()
            TBL_AOP.Dispose()
            DTA_AOP_Item.Dispose()
            TBL_AOP_Item.Dispose()
            DTA_AOP_PR.Dispose()
            TBL_AOP_PR.Dispose()
            DTA_AOP_PR_Item.Dispose()
            TBL_AOP_PR_Item.Dispose()
            DTA_HisPurc.Dispose()
            TBL_HisPurc.Dispose()
            DTA_HisPR.Dispose()
            TBL_HisPR.Dispose()
            DTA_PO.Dispose()
            TBL_PO.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryMaxInvoice(ByVal Tanggal0 As Date, ByVal Tanggal1 As Date, ByVal TipeKartu As ClsTransaksiCreditNote.EnumTipeKartu) As Double
        Dim DTA As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
        Try
            If DTA.ScalarQuery_getPrimaryMaxInvoice(Tanggal0, Tanggal1, TipeKartu).HasValue Then
                Return DTA.ScalarQuery_getPrimaryMaxInvoice(Tanggal0, Tanggal1, TipeKartu).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryMinInvoice(ByVal Tanggal0 As Date, ByVal Tanggal1 As Date, ByVal TipeKartu As ClsTransaksiCreditNote.EnumTipeKartu) As Double
        Dim DTA As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
        Try
            If DTA.ScalarQuery_getPrimaryMinInvoice(Tanggal0, Tanggal1, TipeKartu).HasValue Then
                Return DTA.ScalarQuery_getPrimaryMinInvoice(Tanggal0, Tanggal1, TipeKartu).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
    '    Dim Report_AR_AP As New CrpLaporanCustomerPaidNSupplierPaid
    '    Dim DS As New DseLaporanARPaid
    '    Dim TBLCetak As DseLaporanARPaid.TblCetakDataTable = DS.TblCetak
    '    Dim RowCetak As DseLaporanARPaid.TblCetakRow
    '    Dim TBLCetakTotal As DseLaporanARPaid.TblTotalDataTable = DS.TblTotal
    '    Dim RowCetakTotal As DseLaporanARPaid.TblTotalRow
    '    Dim TBLDummy As New DseLaporanARPaid.TblCetakDataTable
    '    Dim RowDummy As DseLaporanARPaid.TblCetakRow
    '    Dim mFilter01 As String = ""
    '    Dim mFilter02 As String = ""
    '    Dim mKriteria As String = ""
    '    Dim mJudul As String = ""
    '    Dim mJudulKolom As String = ""
    '    Dim Ketemu As Boolean = False
    '    Dim DV01 As DataRowView
    '    Dim DV02 As DataRowView
    '    Dim ij As Integer = 0
    '    Dim mAbaikanHistorical As Boolean = False
    '    Dim ListPrimarySOPONumber As New List(Of Double)
    '    Dim ListPrimaryCurr As New List(Of String)
    '    Dim ListPrimary As New List(Of Double)
    '    '*********ALL DATA*********
    '    Dim DTA_Pay As New DseLaporanARPaidTableAdapters.Item_PaymentTableAdapter
    '    Dim TBL_Pay As New DseLaporanARPaid.Item_PaymentDataTable
    '    Dim Row_Pay As DseLaporanARPaid.Item_PaymentRow
    '    'Dim DTA_NoteDebit As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
    '    'Dim TBL_NoteDebit As New DseLaporanARPaid.tbltrans_credit_note1DataTable
    '    'Dim Row_NoteDebit As DseLaporanARPaid.tbltrans_credit_note1Row
    '    'Dim DTA_NoteKredit As New DseLaporanARPaidTableAdapters.tbltrans_credit_note1TableAdapter
    '    'Dim TBL_NoteKredit As New DseLaporanARPaid.tbltrans_credit_note1DataTable
    '    'Dim Row_NoteKredit As DseLaporanARPaid.tbltrans_credit_note1Row
    '    'Laporan Customer PAID (AR)
    '    Dim DTASaldo_AR As New DseLaporanARPaidTableAdapters.TBLSaldoTableAdapter
    '    Dim TBLSaldo_AR As New DseLaporanARPaid.TBLSaldoDataTable
    '    Dim RowSaldo_AR As DseLaporanARPaid.TBLSaldoRow
    '    Dim DTAInv As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoiceTableAdapter
    '    Dim TBLInv As New DseLaporanARPaid.tbltrans_salesinvoiceDataTable
    '    Dim RowInv As DseLaporanARPaid.tbltrans_salesinvoiceRow
    '    Dim RowInvDummy As DseLaporanARPaid.tbltrans_salesinvoiceRow
    '    Dim DTAInvItem As New DseLaporanARPaidTableAdapters.tbltrans_salesinvoice_itemTableAdapter
    '    Dim TBLInvItem As New DseLaporanARPaid.tbltrans_salesinvoice_itemDataTable
    '    Dim RowInvItem As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
    '    Dim DTA_SR As New DseLaporanARPaidTableAdapters.tbltransaksi_sales_return1TableAdapter
    '    Dim TBL_SR As New DseLaporanARPaid.tbltransaksi_sales_return1DataTable
    '    Dim Row_SR As DseLaporanARPaid.tbltransaksi_sales_return1Row
    '    Dim Row_SRDummy As DseLaporanARPaid.tbltransaksi_sales_return1Row
    '    Dim DTA_SRItem As New DseLaporanARPaidTableAdapters.tbltrans_salesreturn_itemTableAdapter
    '    Dim TBL_SRItem As New DseLaporanARPaid.tbltrans_salesreturn_itemDataTable
    '    Dim Row_SRItem As DseLaporanARPaid.tbltrans_salesinvoice_itemRow
    '    Dim DTA_HisSales As New DseLaporanARPaidTableAdapters.tbltrans_historicalsalesTableAdapter
    '    Dim TBL_HisSales As New DseLaporanARPaid.tbltrans_historicalsalesDataTable
    '    Dim Row_HisSales As DseLaporanARPaid.tbltrans_historicalsalesRow
    '    Dim DTA_HisSR As New DseLaporanARPaidTableAdapters.tbltrans_historical_sales_returnTableAdapter
    '    Dim TBL_HisSR As New DseLaporanARPaid.tbltrans_historical_sales_returnDataTable
    '    Dim Row_HisSR As DseLaporanARPaid.tbltrans_historical_sales_returnRow
    '    Dim DTA_SO As New DseLaporanARPaidTableAdapters.tbltrans_soTableAdapter
    '    Dim TBL_SO As New DseLaporanARPaid.tbltrans_soDataTable
    '    Dim Row_SO As DseLaporanARPaid.tbltrans_soRow
    '    'Laporan Supplier PAID (AP)
    '    Dim DTASaldo_AP As New DseLaporanARPaidTableAdapters.TBLSaldo1TableAdapter
    '    Dim TBLSaldo_AP As New DseLaporanARPaid.TBLSaldo1DataTable
    '    Dim RowSaldo_AP As DseLaporanARPaid.TBLSaldo1Row
    '    Dim DTA_AOP As New DseLaporanARPaidTableAdapters.tbltrans_aopTableAdapter
    '    Dim TBL_AOP As New DseLaporanARPaid.tbltrans_aopDataTable
    '    Dim Row_AOP As DseLaporanARPaid.tbltrans_aopRow
    '    Dim Row_AOPDummy As DseLaporanARPaid.tbltrans_aopRow
    '    Dim DTA_AOP_Item As New DseLaporanARPaidTableAdapters.tbltrans_aop_itemTableAdapter
    '    Dim TBL_AOP_Item As New DseLaporanARPaid.tbltrans_aop_itemDataTable
    '    Dim Row_AOP_Item As DseLaporanARPaid.tbltrans_aop_itemRow
    '    Dim DTA_AOP_PR As New DseLaporanARPaidTableAdapters.tbltrans_purchasereturnTableAdapter
    '    Dim TBL_AOP_PR As New DseLaporanARPaid.tbltrans_purchasereturnDataTable
    '    Dim Row_AOP_PR As DseLaporanARPaid.tbltrans_purchasereturnRow
    '    Dim Row_AOP_PRDummy As DseLaporanARPaid.tbltrans_purchasereturnRow
    '    Dim DTA_AOP_PR_Item As New DseLaporanARPaidTableAdapters.tbltrans_purchasereturn_itemTableAdapter
    '    Dim TBL_AOP_PR_Item As New DseLaporanARPaid.tbltrans_purchasereturn_itemDataTable
    '    Dim row_AOP_PR_Item As DseLaporanARPaid.tbltrans_purchasereturn_itemRow
    '    Dim DTA_HisPurc As New DseLaporanARPaidTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
    '    Dim TBL_HisPurc As New DseLaporanARPaid.tbltrans_historicalpurchasebalancemanualDataTable
    '    Dim Row_HisPurc As DseLaporanARPaid.tbltrans_historicalpurchasebalancemanualRow
    '    Dim DTA_HisPR As New DseLaporanARPaidTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
    '    Dim TBL_HisPR As New DseLaporanARPaid.tbltrans_historicalpurchasereturnbalancemanualDataTable
    '    Dim Row_HisPR As DseLaporanARPaid.tbltrans_historicalpurchasereturnbalancemanualRow
    '    Dim DTA_PO As New DseLaporanARPaidTableAdapters.tbltrans_poTableAdapter
    '    Dim TBL_PO As New DseLaporanARPaid.tbltrans_poDataTable
    '    Dim Row_PO As DseLaporanARPaid.tbltrans_poRow
    '    '*****TBL_Balance Sales*****
    '    Dim DTABal_SalesAutomatic As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    Dim TBLBal_SalesAutomatic As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    Dim RowBal_SalesAutomatic As DseLaporanARPaid.tbltrans_balance_salesRow
    '    Dim DTABal_SalesAutomaticSR As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    Dim TBLBal_SalesAutomaticSR As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    Dim RowBal_SalesAutomaticSR As DseLaporanARPaid.tbltrans_balance_salesRow
    '    Dim DTABal_SalesManual As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    Dim TBLBal_SalesManual As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    Dim RowBal_SalesManual As DseLaporanARPaid.tbltrans_balance_salesRow
    '    Dim DTABal_SalesManualSR As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    Dim TBLBal_SalesManualSR As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    Dim RowBal_SalesManualSR As DseLaporanARPaid.tbltrans_balance_salesRow
    '    'Dim DTABal_SalesDebitNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    'Dim TBLBal_SalesDebitNote As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    'Dim RowBal_SalesDebitNote As DseLaporanARPaid.tbltrans_balance_salesRow
    '    'Dim DTABal_SalesKreditNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_salesTableAdapter
    '    'Dim TBLBal_SalesKreditNote As New DseLaporanARPaid.tbltrans_balance_salesDataTable
    '    'Dim RowBal_SalesKreditNote As DseLaporanARPaid.tbltrans_balance_salesRow
    '    '*****TBL Balance Purchase*****
    '    Dim DTABal_PurcAutomatic As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    Dim TBLBal_PurcAutomatic As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    Dim RowBal_PurcAutomatic As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    Dim DTABal_PurcAutomaticPR As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    Dim TBLBal_PurcAutomaticPR As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    Dim RowBal_PurcAutomaticPR As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    Dim DTABal_PurcManual As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    Dim TBLBal_PurcManual As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    Dim RowBal_PurcManual As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    Dim DTABal_PurcManualPR As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    Dim TBLBal_PurcManualPR As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    Dim RowBal_PurcManualPR As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    'Dim DTABal_PurcDebitNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    'Dim TBLBal_PurcDebitNote As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    'Dim RowBal_PurcDebitNote As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    'Dim DTABal_PurcKreditNote As New DseLaporanARPaidTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    'Dim TBLBal_PurcKreditNote As New DseLaporanARPaid.tbltrans_balance_purchaseDataTable
    '    'Dim RowBal_PurcKreditNote As DseLaporanARPaid.tbltrans_balance_purchaseRow
    '    Try
    '        'cek:
    '        '1. tanggal dari <= tgl sampai
    '        If Me.dtpFrom.Value.Date > Me.dtpFrom.Value.Date Then
    '            MsgBox("Document date from to is not valid.", MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If
    '        '2. tgl as of >= tgl sampai
    '        If Me.DtpAsOf.Value.Date < Me.dtpTo.Value.Date Then
    '            MsgBox("As of date is not less then document date to, please input date again.", MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If
    '        '3. untuk akun harus tipe piutang dagang, jika kartunya ar
    '        '   untuk akun harus tipe hutang dagang, jika kartunya ap
    '        '   udah di setting di main menu
    '        '*****Connection string***** 
    '        DTABal_SalesAutomatic.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_SalesAutomaticSR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_SalesManual.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_SalesManualSR.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTABal_SalesDebitNote.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTABal_SalesKreditNote.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_PurcAutomatic.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_PurcAutomaticPR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_PurcManual.Connection.ConnectionString = Core_Module.GetConstring
    '        DTABal_PurcManualPR.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTABal_PurcDebitNote.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTABal_PurcKreditNote.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_Pay.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTA_NoteDebit.Connection.ConnectionString = Core_Module.GetConstring
    '        'DTA_NoteKredit.Connection.ConnectionString = Core_Module.GetConstring
    '        DTASaldo_AR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAInv.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAInvItem.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_SR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_SRItem.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_HisSales.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_HisSR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_SO.Connection.ConnectionString = Core_Module.GetConstring
    '        DTASaldo_AP.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_AOP.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_AOP_Item.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_AOP_PR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_AOP_PR_Item.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_HisPurc.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_HisPR.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_PO.Connection.ConnectionString = Core_Module.GetConstring
    '        'load data
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            mJudul = "Account Receivable Balance Report"
    '            mJudulKolom = "Customer"
    '            If Me.CboPayment.SelectedIndex = 0 Then
    '                'Unfinish
    '                DTASaldo_AR.Fill(TBLSaldo_AR, Me.TipeKartu, Me.DtpAsOf.Value.Date)
    '            ElseIf Me.CboPayment.SelectedIndex = 1 Then
    '                'All data
    '                DTASaldo_AR.FillBy_All(TBLSaldo_AR, Me.TipeKartu, Me.DtpAsOf.Value.Date)
    '            End If
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            mJudul = "Account Payable Balance Report"
    '            mJudulKolom = "Supplier"
    '            If Me.CboPayment.SelectedIndex = 0 Then
    '                'Unfinish
    '                DTASaldo_AP.Fill(TBLSaldo_AP, Me.TipeKartu, Me.DtpAsOf.Value.Date)
    '            ElseIf Me.CboPayment.SelectedIndex = 1 Then
    '                'All data
    '                DTASaldo_AP.FillBy_All(TBLSaldo_AP, Me.TipeKartu, Me.DtpAsOf.Value.Date)
    '            End If
    '        End If

    '        '1. sales/purchase invoice
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            DTAInv.Fill(TBLInv, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '            DTAInvItem.Fill(TBLInvItem, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            DTA_AOP.Fill(TBL_AOP, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '            DTA_AOP_Item.Fill(TBL_AOP_Item, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        End If
    '        '2. sales/purchase return
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            DTA_SR.Fill(TBL_SR, Me.TipeStockInSR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipeSR)
    '            DTA_SRItem.Fill(TBL_SRItem, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipeSR)
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            DTA_AOP_PR.Fill(TBL_AOP_PR, Me.TipePR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '            DTA_AOP_PR_Item.Fill(TBL_AOP_PR_Item, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date, Me.TipePR)
    '        End If
    '        '3. historical sales/purchase
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            DTA_HisSales.Fill(TBL_HisSales, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            DTA_HisPurc.Fill(TBL_HisPurc, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        End If
    '        '4. historical sales/purchase return
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            DTA_HisSR.Fill(TBL_HisSR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            DTA_HisPR.Fill(TBL_HisPR, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
    '        End If
    '        ''5. credit note/debit note
    '        'DTA_NoteDebit.Fill(TBL_NoteDebit, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)
    '        'DTA_NoteKredit.Fill(TBL_NoteKredit, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeKartu)

    '        '6. filter customer, ar/ap account, currency, code no, so number, supplier, so no pada tabel 1,2,3,4,5
    '        mAbaikanHistorical = Me.AbaikanHistorical(Me.TipeKartu)
    '        '***FILTER*** ????
    '        mFilter01 = Me.GetFilterHeader()
    '        mFilter02 = Me.GetFilterItem()

    '        '***Hapus Data Yang Tidak Sesuai Kriteria FILTER***
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            '*****Kartu AR*****
    '            If mAbaikanHistorical = False Then
    '                If Me.ApakahAbaikanKriteriaFilterData() = False Then
    '                    'TBL Invoice 
    '                    ij = 0
    '                    While ij <= TBLInv.Rows.Count - 1
    '                        RowInv = TBLInv.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> RowInv.Primary_Customer Then
    '                                    TBLInv.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> RowInv.primary_akun_ar Then
    '                                    TBLInv.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Sales Return
    '                    ij = 0
    '                    While ij <= TBL_SR.Rows.Count - 1
    '                        Row_SR = TBL_SR.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> Row_SR.Primary_Customer Then
    '                                    TBL_SR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_SR.primary_akun_ar Then
    '                                    TBL_SR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Historical Sales
    '                    ij = 0
    '                    While ij <= TBL_HisSales.Rows.Count - 1
    '                        Row_HisSales = TBL_HisSales.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCurrency.Text <> "-1" Then
    '                                If Me.txtPrimaryCurrency.Text <> Row_HisSales.Primary_Currency Then
    '                                    TBL_HisSales.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> Row_HisSales.Primary_Customer Then
    '                                    TBL_HisSales.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_HisSales.primary_akun_ar Then
    '                                    TBL_HisSales.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Historical Sales Return
    '                    ij = 0
    '                    While ij <= TBL_HisSR.Rows.Count - 1
    '                        Row_HisSR = TBL_HisSR.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCurrency.Text <> "-1" Then
    '                                If Me.txtPrimaryCurrency.Text <> Row_HisSR.Primary_Currency Then
    '                                    TBL_HisSR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> Row_HisSR.Primary_Customer Then
    '                                    TBL_HisSR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_HisSR.primary_akun_ar Then
    '                                    TBL_HisSR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    '****
    '                End If
    '                Me.GetPrimaryMinMax_SalesOrder(TBLInvItem)
    '            ElseIf mAbaikanHistorical = True Then
    '                'TBL Invoice
    '                TBLInvItem.DefaultView.RowFilter = Nothing
    '                TBLInvItem.DefaultView.RowFilter = mFilter02
    '                If TBLInvItem.DefaultView.Count > 0 Then
    '                    mKriteria = ""
    '                    For Each DV01 In TBLInvItem.DefaultView
    '                        RowInvItem = DV01.Row
    '                        If String.IsNullOrEmpty(mKriteria) = False Then
    '                            mKriteria += " AND [Primary Transaksi] <> " & RowInvItem.Primary_Transaksi.ToString
    '                        Else
    '                            mKriteria = "[Primary Transaksi] <> " & RowInvItem.Primary_Transaksi.ToString
    '                        End If
    '                    Next
    '                    TBLInv.DefaultView.RowFilter = Nothing
    '                    TBLInv.DefaultView.RowFilter = mKriteria
    '                    ij = 0
    '                    While ij <= TBLInv.DefaultView.Count - 1
    '                        DV02 = TBLInv.DefaultView(ij)
    '                        RowInv = DV02.Row
    '                        TBLInv.Rows.Remove(RowInv)
    '                    End While
    '                    ij = 0
    '                    While ij <= TBLInv.Rows.Count - 1
    '                        Ketemu = False
    '                        RowInv = TBLInv.Rows(ij)
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> RowInv.Primary_Customer Then
    '                                    TBLInv.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> RowInv.primary_akun_ar Then
    '                                    TBLInv.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                Else
    '                    TBLInv.Rows.Clear()
    '                End If
    '                'TBL Sales Return
    '                TBL_SRItem.DefaultView.RowFilter = Nothing
    '                TBL_SRItem.DefaultView.RowFilter = mFilter02
    '                If TBL_SRItem.DefaultView.Count > 0 Then
    '                    mKriteria = ""
    '                    For Each DV01 In TBL_SRItem.DefaultView
    '                        Row_SRItem = DV01.Row
    '                        If String.IsNullOrEmpty(mKriteria) = False Then
    '                            mKriteria += " AND [Primary Transaksi] <> " & Row_SRItem.Primary_Transaksi.ToString
    '                        Else
    '                            mKriteria = "[Primary Transaksi] <> " & Row_SRItem.Primary_Transaksi.ToString
    '                        End If
    '                    Next
    '                    TBL_SR.DefaultView.RowFilter = Nothing
    '                    TBL_SR.DefaultView.RowFilter = mKriteria
    '                    ij = 0
    '                    While ij <= TBL_SR.DefaultView.Count - 1
    '                        DV02 = TBL_SR.DefaultView(ij)
    '                        Row_SR = DV02.Row
    '                        TBL_SR.Rows.Remove(Row_SR)
    '                    End While
    '                    ij = 0
    '                    While ij <= TBL_SR.Rows.Count - 1
    '                        Ketemu = False
    '                        Row_SR = TBL_SR.Rows(ij)
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCust.Text <> "-1" Then
    '                                If Me.txtPrimaryCust.Text <> Row_SR.Primary_Customer Then
    '                                    TBL_SR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_SR.primary_akun_ar Then
    '                                    TBL_SR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                Else
    '                    TBL_SR.Rows.Clear()
    '                End If
    '                '****
    '            End If
    '            Me.GetPrimaryMinMax_SalesOrder(TBLInvItem)
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            '*****Kartu AP*****
    '            If mAbaikanHistorical = False Then
    '                If Me.ApakahAbaikanKriteriaFilterData() = False Then
    '                    'TBL AOP 
    '                    ij = 0
    '                    While ij <= TBL_AOP.Rows.Count - 1
    '                        Row_AOP = TBL_AOP.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_AOP.Primary_Supplier Then
    '                                    TBL_AOP.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_AOP.primary_akun_ap Then
    '                                    TBL_AOP.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Purchase Return
    '                    ij = 0
    '                    While ij <= TBL_AOP_PR.Rows.Count - 1
    '                        Row_AOP_PR = TBL_AOP_PR.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_AOP_PR.Primary_Supplier Then
    '                                    TBL_AOP_PR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_AOP_PR.primary_akun_ap Then
    '                                    TBL_AOP_PR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Historical Purchase
    '                    ij = 0
    '                    While ij <= TBL_HisPurc.Rows.Count - 1
    '                        Row_HisPurc = TBL_HisPurc.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCurrency.Text <> "-1" Then
    '                                If Me.txtPrimaryCurrency.Text <> Row_HisPurc.Primary_Currency Then
    '                                    TBL_HisPurc.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_HisPurc.Primary_Supplier Then
    '                                    TBL_HisPurc.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_HisPurc.primary_akun_ap Then
    '                                    TBL_HisPurc.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    'TBL Historical Purchase Return 
    '                    ij = 0
    '                    While ij <= TBL_HisPR.Rows.Count - 1
    '                        Row_HisPR = TBL_HisPR.Rows(ij)
    '                        Ketemu = False
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryCurrency.Text <> "-1" Then
    '                                If Me.txtPrimaryCurrency.Text <> Row_HisPR.Primary_Currency Then
    '                                    TBL_HisPR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_HisPR.Primary_Supplier Then
    '                                    TBL_HisPR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_HisPR.primary_akun_ap Then
    '                                    TBL_HisPR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                    '****
    '                End If
    '                Me.GetPrimaryMinMax_PurchaseOrder(TBL_AOP_Item)
    '            ElseIf mAbaikanHistorical = True Then
    '                'TBL AOP
    '                TBL_AOP_Item.DefaultView.RowFilter = Nothing
    '                TBL_AOP_Item.DefaultView.RowFilter = mFilter02
    '                If TBL_AOP_Item.DefaultView.Count > 0 Then
    '                    mKriteria = ""
    '                    For Each DV01 In TBL_AOP_Item.DefaultView
    '                        Row_AOP_Item = DV01.Row
    '                        If String.IsNullOrEmpty(mKriteria) = False Then
    '                            mKriteria += " AND [Primary Transaksi] <> " & Row_AOP_Item.Primary_Transaksi.ToString
    '                        Else
    '                            mKriteria = "[Primary Transaksi] <> " & Row_AOP_Item.Primary_Transaksi.ToString
    '                        End If
    '                    Next
    '                    TBL_AOP.DefaultView.RowFilter = Nothing
    '                    TBL_AOP.DefaultView.RowFilter = mKriteria
    '                    ij = 0
    '                    While ij <= TBL_AOP.DefaultView.Count - 1
    '                        DV02 = TBL_AOP.DefaultView(ij)
    '                        Row_AOP = DV02.Row
    '                        TBL_AOP.Rows.Remove(Row_AOP)
    '                    End While
    '                    ij = 0
    '                    While ij <= TBL_AOP.Rows.Count - 1
    '                        Ketemu = False
    '                        Row_AOP = TBL_AOP.Rows(ij)
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_AOP.Primary_Supplier Then
    '                                    TBL_AOP.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_AOP.primary_akun_ap Then
    '                                    TBL_AOP.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                Else
    '                    TBL_AOP.Rows.Clear()
    '                End If
    '                'TBL Purchase Return
    '                TBL_AOP_PR_Item.DefaultView.RowFilter = Nothing
    '                TBL_AOP_PR_Item.DefaultView.RowFilter = mFilter02
    '                If TBL_AOP_PR_Item.DefaultView.Count > 0 Then
    '                    mKriteria = ""
    '                    For Each DV01 In TBL_AOP_PR_Item.DefaultView
    '                        row_AOP_PR_Item = DV01.Row
    '                        If String.IsNullOrEmpty(mKriteria) = False Then
    '                            mKriteria += " AND [Primary Transaksi] <> " & row_AOP_PR_Item.Primary_Transaksi.ToString
    '                        Else
    '                            mKriteria = "[Primary Transaksi] <> " & row_AOP_PR_Item.Primary_Transaksi.ToString
    '                        End If
    '                    Next
    '                    TBL_AOP_PR.DefaultView.RowFilter = Nothing
    '                    TBL_AOP_PR.DefaultView.RowFilter = mKriteria
    '                    ij = 0
    '                    While ij <= TBL_AOP_PR.DefaultView.Count - 1
    '                        DV02 = TBL_AOP_PR.DefaultView(ij)
    '                        Row_AOP_PR = DV02.Row
    '                        TBL_AOP_PR.Rows.Remove(Row_AOP_PR)
    '                    End While
    '                    ij = 0
    '                    While ij <= TBL_AOP_PR.Rows.Count - 1
    '                        Ketemu = False
    '                        Row_AOP_PR = TBL_AOP_PR.Rows(ij)
    '                        If Ketemu = False Then
    '                            If Me.TxtPrimarySupplier.Text <> "-1" Then
    '                                If Me.TxtPrimarySupplier.Text <> Row_AOP_PR.Primary_Supplier Then
    '                                    TBL_AOP_PR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            If Me.txtPrimaryArAccount.Text <> "-1" Then
    '                                If Me.txtPrimaryArAccount.Text <> Row_AOP_PR.primary_akun_ap Then
    '                                    TBL_AOP_PR.Rows.RemoveAt(ij)
    '                                    Ketemu = True
    '                                End If
    '                            End If
    '                        End If
    '                        If Ketemu = False Then
    '                            ij += 1
    '                        End If
    '                    End While
    '                Else
    '                    TBL_AOP_PR.Rows.Clear()
    '                End If
    '                '****
    '            End If
    '            Me.GetPrimaryMinMax_PurchaseOrder(TBL_AOP_Item)
    '        End If

    '        '6.1. isikan tabel balance untuk masing2 tabel 1,2,3,4,5 cari nilai transaction amount
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            'Primary Min Max
    '            TBLInv.DefaultView.RowFilter = Nothing
    '            TBLInv.DefaultView.Sort = Nothing
    '            TBLInv.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBLInv.DefaultView.Count > 0 Then
    '                RowInv = TBLInv.DefaultView(0).Row
    '                Me.PrimaryMinInv = RowInv.Primary_Transaksi
    '                RowInv = TBLInv.DefaultView(TBLInv.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxInv = RowInv.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinInv = -1
    '                Me.PrimaryMaxInv = -1
    '            End If
    '            TBL_SR.DefaultView.RowFilter = Nothing
    '            TBL_SR.DefaultView.Sort = Nothing
    '            TBL_SR.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_SR.DefaultView.Count > 0 Then
    '                Row_SR = TBL_SR.DefaultView(0).Row
    '                Me.PrimaryMinSR = Row_SR.Primary_Transaksi
    '                Row_SR = TBL_SR.DefaultView(TBL_SR.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxSR = Row_SR.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinSR = -1
    '                Me.PrimaryMaxSR = -1
    '            End If
    '            TBL_HisSales.DefaultView.RowFilter = Nothing
    '            TBL_HisSales.DefaultView.Sort = Nothing
    '            TBL_HisSales.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_HisSales.DefaultView.Count > 0 Then
    '                Row_HisSales = TBL_HisSales.DefaultView(0).Row
    '                Me.PrimaryMinHistoricalSales = Row_HisSales.Primary_Transaksi
    '                Row_HisSales = TBL_HisSales.DefaultView(TBL_HisSales.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxHistoricalSales = Row_HisSales.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinHistoricalSales = -1
    '                Me.PrimaryMaxHistoricalSales = -1
    '            End If
    '            TBL_HisSR.DefaultView.RowFilter = Nothing
    '            TBL_HisSR.DefaultView.Sort = Nothing
    '            TBL_HisSR.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_HisSR.DefaultView.Count > 0 Then
    '                Row_HisSR = TBL_HisSR.DefaultView(0).Row
    '                Me.PrimaryMinHistoricalSalesReturn = Row_HisSR.Primary_Transaksi
    '                Row_HisSR = TBL_HisSR.DefaultView(TBL_HisSR.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxHistoricalSalesReturn = Row_HisSR.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinHistoricalSalesReturn = -1
    '                Me.PrimaryMaxHistoricalSalesReturn = -1
    '            End If
    '            DTA_SO.Fill(TBL_SO, Me.PrimaryMaxSO, Me.PrimaryMinSO)
    '            DTA_Pay.Fill(TBL_Pay, Me.TipeJurnalAR, Me.dtpFrom.Value.Date, Me.DtpAsOf.Value.Date)
    '            DTABal_SalesAutomatic.Fill(TBLBal_SalesAutomatic, Me.PrimaryMaxInv, Me.PrimaryMinInv, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic)
    '            DTABal_SalesAutomaticSR.Fill(TBLBal_SalesAutomaticSR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR)
    '            DTABal_SalesManual.Fill(TBLBal_SalesManual, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual)
    '            DTABal_SalesManualSR.Fill(TBLBal_SalesManualSR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR)

    '            'Get Transaction Amount Invoice
    '            If TBLBal_SalesAutomatic.Rows.Count > 0 Then
    '                ListPrimarySOPONumber.Clear()
    '                ListPrimaryCurr.Clear()
    '                ListPrimary.Clear()
    '                For Each RowBal_SalesAutomatic In TBLBal_SalesAutomatic.Rows
    '                    TBLInv.DefaultView.RowFilter = Nothing
    '                    TBLInv.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesAutomatic.Primary_Invoice.ToString
    '                    If TBLInv.DefaultView.Count > 0 Then
    '                        If ListPrimary.Contains(RowBal_SalesAutomatic.Primary_Invoice) = False Then
    '                            ListPrimary.Add(RowBal_SalesAutomatic.Primary_Invoice)
    '                            RowInv = TBLInv.DefaultView(0).Row
    '                            'Input Transaction Amount
    '                            RowInv.Transaction_Amount += RowBal_SalesAutomatic.balance_original + RowBal_SalesAutomatic.received_original
    '                            'Input Paid Date
    '                            TBL_Pay.DefaultView.RowFilter = Nothing
    '                            TBL_Pay.DefaultView.Sort = Nothing
    '                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowInv.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & RowInv.Primary_Currency.ToString
    '                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                            If TBL_Pay.DefaultView.Count > 0 Then
    '                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                RowInv.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                            End If
    '                            'Input Cust PO No
    '                            TBLInvItem.DefaultView.RowFilter = Nothing
    '                            TBLInvItem.DefaultView.RowFilter = "[Primary Transaksi] = " & RowInv.Primary_Transaksi.ToString
    '                            If TBLInvItem.DefaultView.Count > 0 Then
    '                                TBL_SO.DefaultView.RowFilter = Nothing
    '                                TBL_SO.DefaultView.Sort = Nothing
    '                                For Each DV01 In TBLInvItem.DefaultView
    '                                    RowInvItem = DV01.Row
    '                                    TBL_SO.DefaultView.RowFilter = "[Primary SO] = " & RowInvItem.Primary_SO.ToString
    '                                    TBL_SO.DefaultView.Sort = "[Cust PO No]"
    '                                    If TBL_SO.DefaultView.Count > 0 Then
    '                                        For Each DV02 In TBL_SO.DefaultView
    '                                            Row_SO = DV02.Row
    '                                            If String.IsNullOrEmpty(Row_SO.Cust_PO_No) = False Then
    '                                                If ListPrimarySOPONumber.Contains(Row_SO.Primary_SO) = False Then
    '                                                    ListPrimarySOPONumber.Add(Row_SO.Primary_SO)
    '                                                    If String.IsNullOrEmpty(RowInv.Cust_PO_Number) = False Then
    '                                                        RowInv.Cust_PO_Number += ", " + Row_SO.Cust_PO_No.ToString
    '                                                    Else
    '                                                        RowInv.Cust_PO_Number = Row_SO.Cust_PO_No.ToString
    '                                                    End If
    '                                                End If
    '                                            Else
    '                                                If String.IsNullOrEmpty(RowInv.Cust_PO_Number) = False Then
    '                                                    RowInv.Cust_PO_Number += ", " + Row_SO.No_Transaksi.ToString
    '                                                Else
    '                                                    RowInv.Cust_PO_Number = Row_SO.No_Transaksi.ToString
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    End If
    '                                Next
    '                            End If
    '                            'Input Baris Baru
    '                            If RowBal_SalesAutomatic.Primary_Currency <> 1 Then
    '                                If RowInv.Tax_Currency = 0 Then
    '                                    'Tambah Baris Buat IDR
    '                                    RowInvDummy = TBLInv.NewRow
    '                                    For xy As Integer = 0 To TBLInv.Columns.Count - 1
    '                                        RowInvDummy.Item(xy) = RowInv.Item(xy)
    '                                    Next
    '                                    TBLInv.Rows.Add(RowInvDummy)
    '                                    'Currency Code-Nya IDR
    '                                    RowInvDummy = TBLInv.Rows(TBLInv.Rows.Count - 1)
    '                                    RowInvDummy.Primary_Currency = 1
    '                                    RowInvDummy.Transaction_Amount = 0
    '                                    RowInvDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                            End If
    '                        ElseIf ListPrimary.Contains(RowBal_SalesAutomatic.Primary_Invoice) = True Then
    '                            If RowBal_SalesAutomatic.Primary_Currency = 1 Then
    '                                RowInv = TBLInv.DefaultView(0).Row
    '                                If RowInv.Tax_Currency = 0 Then
    '                                    'Input Pasangan USD yang IDR-Nya
    '                                    RowInv = TBLInv.Rows(TBLInv.Rows.Count - 1)
    '                                    RowInv.Transaction_Amount += RowBal_SalesAutomatic.balance_original + RowBal_SalesAutomatic.received_original
    '                                    'Input Paid Date
    '                                    TBL_Pay.DefaultView.RowFilter = Nothing
    '                                    TBL_Pay.DefaultView.Sort = Nothing
    '                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & RowInv.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & RowInv.Primary_Currency.ToString
    '                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                                    If TBL_Pay.DefaultView.Count > 0 Then
    '                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                        RowInv.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            'Get Transaction Amount Sales Return
    '            If TBLBal_SalesAutomaticSR.Rows.Count > 0 Then
    '                ListPrimaryCurr.Clear()
    '                ListPrimary.Clear()
    '                For Each RowBal_SalesAutomaticSR In TBLBal_SalesAutomaticSR.Rows
    '                    TBL_SR.DefaultView.RowFilter = Nothing
    '                    TBL_SR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_SalesAutomaticSR.Primary_Invoice.ToString
    '                    If TBL_SR.DefaultView.Count > 0 Then
    '                        If ListPrimary.Contains(RowBal_SalesAutomaticSR.Primary_Invoice) = False Then
    '                            ListPrimary.Add(RowBal_SalesAutomaticSR.Primary_Invoice)
    '                            Row_SR = TBL_SR.DefaultView(0).Row
    '                            'Input Transaction Amount
    '                            Row_SR.Transaction_Amount += RowBal_SalesAutomaticSR.balance_original + RowBal_SalesAutomaticSR.received_original
    '                            'Input Paid Date
    '                            TBL_Pay.DefaultView.RowFilter = Nothing
    '                            TBL_Pay.DefaultView.Sort = Nothing
    '                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_SR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_SR.Primary_Currency.ToString
    '                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                            If TBL_Pay.DefaultView.Count > 0 Then
    '                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                Row_SR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                            End If
    '                            'Input Baris Baru
    '                            If RowBal_SalesAutomaticSR.Primary_Currency <> 1 Then
    '                                If Row_SR.Tax_Currency = 0 Then
    '                                    'Tambah Baris Buat IDR
    '                                    Row_SRDummy = TBL_SR.NewRow
    '                                    For xy As Integer = 0 To TBL_SR.Columns.Count - 1
    '                                        Row_SRDummy.Item(xy) = Row_SR.Item(xy)
    '                                    Next
    '                                    TBL_SR.Rows.Add(Row_SRDummy)
    '                                    'Currency Code-Nya IDR
    '                                    Row_SRDummy = TBL_SR.Rows(TBL_SR.Rows.Count - 1)
    '                                    Row_SRDummy.Primary_Currency = 1
    '                                    Row_SRDummy.Transaction_Amount = 0
    '                                    Row_SRDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                            End If
    '                        ElseIf ListPrimary.Contains(RowBal_SalesAutomaticSR.Primary_Invoice) = True Then
    '                            If RowBal_SalesAutomaticSR.Primary_Currency = 1 Then
    '                                Row_SR = TBL_SR.DefaultView(0).Row
    '                                If Row_SR.Tax_Currency = 0 Then
    '                                    'Input Pasangan USD yang IDR-Nya
    '                                    Row_SR = TBL_SR.Rows(TBL_SR.Rows.Count - 1)
    '                                    Row_SR.Transaction_Amount += RowBal_SalesAutomaticSR.balance_original + RowBal_SalesAutomaticSR.received_original
    '                                    'Input Paid Date
    '                                    TBL_Pay.DefaultView.RowFilter = Nothing
    '                                    TBL_Pay.DefaultView.Sort = Nothing
    '                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_SR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_SR.Primary_Currency.ToString
    '                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                                    If TBL_Pay.DefaultView.Count > 0 Then
    '                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                        Row_SR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            'Historical Sales Invoice dan Sales Return
    '            If mAbaikanHistorical = False Then
    '                If TBL_HisSales.Rows.Count > 0 Then
    '                    For Each Row_HisSales In TBL_HisSales.Rows
    '                        TBLBal_SalesManual.DefaultView.RowFilter = Nothing
    '                        TBLBal_SalesManual.DefaultView.RowFilter = "[Primary Invoice] = " & Row_HisSales.Primary_Transaksi.ToString & " AND [Primary Customer] = " & Row_HisSales.Primary_Customer.ToString & " AND [Primary Currency] = " & Row_HisSales.Primary_Currency.ToString
    '                        If TBLBal_SalesManual.DefaultView.Count > 0 Then
    '                            For Each DV01 In TBLBal_SalesManual.DefaultView
    '                                RowBal_SalesManual = DV01.Row
    '                                Row_HisSales.Transaction_Amount += RowBal_SalesManual.balance_original + RowBal_SalesManual.received_original
    '                            Next
    '                        End If
    '                        TBL_Pay.DefaultView.RowFilter = Nothing
    '                        TBL_Pay.DefaultView.Sort = Nothing
    '                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
    '                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisSales.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisSales.Primary_Currency.ToString
    '                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                        If TBL_Pay.DefaultView.Count > 0 Then
    '                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                            Row_HisSales.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                        End If
    '                    Next
    '                End If
    '                If TBL_HisSR.Rows.Count > 0 Then
    '                    For Each Row_HisSR In TBL_HisSR.Rows
    '                        TBLBal_SalesManualSR.DefaultView.RowFilter = Nothing
    '                        TBLBal_SalesManualSR.DefaultView.RowFilter = "[Primary Invoice] = " & Row_HisSR.Primary_Transaksi.ToString & " AND [Primary Customer] = " & Row_HisSR.Primary_Customer.ToString & " AND [Primary Currency] = " & Row_HisSR.Primary_Currency.ToString
    '                        If TBLBal_SalesManualSR.DefaultView.Count > 0 Then
    '                            For Each DV01 In TBLBal_SalesManualSR.DefaultView
    '                                RowBal_SalesManualSR = DV01.Row
    '                                Row_HisSR.Transaction_Amount += RowBal_SalesManualSR.balance_original + RowBal_SalesManualSR.received_original
    '                            Next
    '                        End If
    '                        TBL_Pay.DefaultView.RowFilter = Nothing
    '                        TBL_Pay.DefaultView.Sort = Nothing
    '                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisSR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisSR.Primary_Currency.ToString
    '                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                        If TBL_Pay.DefaultView.Count > 0 Then
    '                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                            Row_HisSR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                        End If
    '                    Next
    '                End If
    '                '
    '            End If
    '            '***************
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            'Primary Min Max
    '            TBL_AOP.DefaultView.RowFilter = Nothing
    '            TBL_AOP.DefaultView.Sort = Nothing
    '            TBL_AOP.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_AOP.DefaultView.Count > 0 Then
    '                Row_AOP = TBL_AOP.DefaultView(0).Row
    '                Me.PrimaryMinInv = Row_AOP.Primary_Transaksi
    '                Row_AOP = TBL_AOP.DefaultView(TBL_AOP.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxInv = Row_AOP.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinInv = -1
    '                Me.PrimaryMaxInv = -1
    '            End If
    '            TBL_AOP_PR.DefaultView.RowFilter = Nothing
    '            TBL_AOP_PR.DefaultView.Sort = Nothing
    '            TBL_AOP_PR.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_AOP_PR.DefaultView.Count > 0 Then
    '                Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
    '                Me.PrimaryMinSR = Row_AOP_PR.Primary_Transaksi
    '                Row_AOP_PR = TBL_AOP_PR.DefaultView(TBL_AOP_PR.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxSR = Row_AOP_PR.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinSR = -1
    '                Me.PrimaryMaxSR = -1
    '            End If
    '            TBL_HisPurc.DefaultView.RowFilter = Nothing
    '            TBL_HisPurc.DefaultView.Sort = Nothing
    '            TBL_HisPurc.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_HisPurc.DefaultView.Count > 0 Then
    '                Row_HisPurc = TBL_HisPurc.DefaultView(0).Row
    '                Me.PrimaryMinHistoricalSales = Row_HisPurc.Primary_Transaksi
    '                Row_HisPurc = TBL_HisPurc.DefaultView(TBL_HisPurc.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxHistoricalSales = Row_HisPurc.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinHistoricalSales = -1
    '                Me.PrimaryMaxHistoricalSales = -1
    '            End If
    '            TBL_HisPR.DefaultView.RowFilter = Nothing
    '            TBL_HisPR.DefaultView.Sort = Nothing
    '            TBL_HisPR.DefaultView.Sort = "[Primary Transaksi]"
    '            If TBL_HisPR.DefaultView.Count > 0 Then
    '                Row_HisPR = TBL_HisPR.DefaultView(0).Row
    '                Me.PrimaryMinHistoricalSalesReturn = Row_HisPR.Primary_Transaksi
    '                Row_HisPR = TBL_HisPR.DefaultView(TBL_HisPR.DefaultView.Count - 1).Row
    '                Me.PrimaryMaxHistoricalSalesReturn = Row_HisPR.Primary_Transaksi
    '            Else
    '                Me.PrimaryMinHistoricalSalesReturn = -1
    '                Me.PrimaryMaxHistoricalSalesReturn = -1
    '            End If
    '            DTA_PO.Fill(TBL_PO, Me.PrimaryMinSO, Me.PrimaryMaxSO)
    '            DTA_Pay.Fill(TBL_Pay, Me.TipeJurnalAP, Me.dtpFrom.Value.Date, Me.DtpAsOf.Value.Date)
    '            DTABal_PurcAutomatic.Fill(TBLBal_PurcAutomatic, Me.PrimaryMaxInv, Me.PrimaryMinInv, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic)
    '            DTABal_PurcAutomaticPR.Fill(TBLBal_PurcAutomaticPR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR)
    '            DTABal_PurcManual.Fill(TBLBal_PurcManual, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual)
    '            DTABal_PurcManualPR.Fill(TBLBal_PurcManualPR, Me.PrimaryMaxSR, Me.PrimaryMinSR, ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR)

    '            'Get Transaction Amount Purchase
    '            If TBLBal_PurcAutomatic.Rows.Count > 0 Then
    '                ListPrimarySOPONumber.Clear()
    '                ListPrimaryCurr.Clear()
    '                ListPrimary.Clear()
    '                For Each RowBal_PurcAutomatic In TBLBal_PurcAutomatic.Rows
    '                    TBL_AOP.DefaultView.RowFilter = Nothing
    '                    TBL_AOP.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcAutomatic.Primary_AOP.ToString
    '                    If TBL_AOP.DefaultView.Count > 0 Then
    '                        If ListPrimary.Contains(RowBal_PurcAutomatic.Primary_AOP) = False Then
    '                            ListPrimary.Add(RowBal_PurcAutomatic.Primary_AOP)
    '                            Row_AOP = TBL_AOP.DefaultView(0).Row
    '                            'Input Transaction Amount
    '                            Row_AOP.Transaction_Amount += RowBal_PurcAutomatic.balance_original + RowBal_PurcAutomatic.paid_original
    '                            'Input Paid Date
    '                            TBL_Pay.DefaultView.RowFilter = Nothing
    '                            TBL_Pay.DefaultView.Sort = Nothing
    '                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP.Primary_Currency.ToString
    '                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                            If TBL_Pay.DefaultView.Count > 0 Then
    '                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                Row_AOP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                            End If
    '                            'Cust PO No
    '                            TBL_AOP_Item.DefaultView.RowFilter = Nothing
    '                            TBL_AOP_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & Row_AOP.Primary_Transaksi.ToString
    '                            If TBL_AOP_Item.DefaultView.Count > 0 Then
    '                                TBL_PO.DefaultView.RowFilter = Nothing
    '                                TBL_PO.DefaultView.Sort = Nothing
    '                                For Each DV01 In TBL_AOP_Item.DefaultView
    '                                    Row_AOP_Item = DV01.Row
    '                                    TBL_PO.DefaultView.RowFilter = "[Primary PO] = " & Row_AOP_Item.Primary_PO.ToString
    '                                    TBL_PO.DefaultView.Sort = "[No Transaksi]"
    '                                    If TBL_PO.DefaultView.Count > 0 Then
    '                                        For Each DV02 In TBL_PO.DefaultView
    '                                            Row_PO = DV02.Row
    '                                            If String.IsNullOrEmpty(Row_PO.No_Transaksi) = False Then
    '                                                If ListPrimarySOPONumber.Contains(Row_PO.Primary_PO) = False Then
    '                                                    ListPrimarySOPONumber.Add(Row_PO.Primary_PO)
    '                                                    If String.IsNullOrEmpty(Row_AOP.PO_Number) = False Then
    '                                                        Row_AOP.PO_Number += ", " + Row_PO.No_Transaksi.ToString
    '                                                    Else
    '                                                        Row_AOP.PO_Number = Row_PO.No_Transaksi.ToString
    '                                                    End If
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    End If
    '                                Next
    '                            End If
    '                            'Input Baris Baru
    '                            If RowBal_PurcAutomatic.Primary_Currency <> 1 Then
    '                                If Row_AOP.Tax_Currency = 0 Then
    '                                    'Tambah Baris Buat IDR
    '                                    Row_AOPDummy = TBL_AOP.NewRow
    '                                    For xy As Integer = 0 To TBL_AOP.Columns.Count - 1
    '                                        Row_AOPDummy.Item(xy) = Row_AOP.Item(xy)
    '                                    Next
    '                                    TBL_AOP.Rows.Add(Row_AOPDummy)
    '                                    'Currency Code-Nya IDR
    '                                    Row_AOPDummy = TBL_AOP.Rows(TBL_AOP.Rows.Count - 1)
    '                                    Row_AOPDummy.Primary_Currency = 1
    '                                    Row_AOPDummy.Transaction_Amount = 0
    '                                    Row_AOPDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                            End If
    '                        ElseIf ListPrimary.Contains(RowBal_PurcAutomatic.Primary_AOP) = True Then
    '                            If RowBal_PurcAutomatic.Primary_Currency = 1 Then
    '                                Row_AOP = TBL_AOP.DefaultView(0).Row
    '                                If Row_AOP.Tax_Currency = 0 Then
    '                                    'Input Pasangan USD yang IDR-Nya
    '                                    Row_AOP = TBL_AOP.Rows(TBL_AOP.Rows.Count - 1)
    '                                    Row_AOP.Transaction_Amount += RowBal_PurcAutomatic.balance_original + RowBal_PurcAutomatic.paid_original
    '                                    'Input Paid Date
    '                                    TBL_Pay.DefaultView.RowFilter = Nothing
    '                                    TBL_Pay.DefaultView.Sort = Nothing
    '                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP.Primary_Currency.ToString
    '                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                                    If TBL_Pay.DefaultView.Count > 0 Then
    '                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                        Row_AOP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            'Get Transaction Amount Purchase Return
    '            If TBLBal_PurcAutomaticPR.Rows.Count > 0 Then
    '                ListPrimaryCurr.Clear()
    '                ListPrimary.Clear()
    '                For Each RowBal_PurcAutomaticPR In TBLBal_PurcAutomaticPR.Rows
    '                    TBL_AOP_PR.DefaultView.RowFilter = Nothing
    '                    TBL_AOP_PR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowBal_PurcAutomaticPR.Primary_AOP.ToString
    '                    If TBL_AOP_PR.DefaultView.Count > 0 Then
    '                        If ListPrimary.Contains(RowBal_PurcAutomaticPR.Primary_AOP) = False Then
    '                            ListPrimary.Add(RowBal_PurcAutomaticPR.Primary_AOP)
    '                            Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
    '                            'Input Transaction Amount
    '                            Row_AOP_PR.Transaction_Amount += RowBal_PurcAutomaticPR.balance_original + RowBal_PurcAutomaticPR.paid_original
    '                            'Input Paid Date
    '                            TBL_Pay.DefaultView.RowFilter = Nothing
    '                            TBL_Pay.DefaultView.Sort = Nothing
    '                            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP_PR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP_PR.Primary_Currency.ToString
    '                            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                            If TBL_Pay.DefaultView.Count > 0 Then
    '                                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                Row_AOP_PR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                            End If
    '                            'Input Baris Baru
    '                            If RowBal_PurcAutomaticPR.Primary_Currency <> 1 Then
    '                                If Row_AOP_PR.Tax_Currency = 0 Then
    '                                    'Tambah Baris Buat IDR
    '                                    Row_AOP_PRDummy = TBL_AOP_PR.NewRow
    '                                    For xy As Integer = 0 To TBL_AOP_PR.Columns.Count - 1
    '                                        Row_AOP_PRDummy.Item(xy) = Row_AOP_PR.Item(xy)
    '                                    Next
    '                                    TBL_AOP_PR.Rows.Add(Row_AOP_PRDummy)
    '                                    'Currency Code-Nya IDR
    '                                    Row_AOP_PRDummy = TBL_AOP_PR.Rows(TBL_AOP_PR.Rows.Count - 1)
    '                                    Row_AOP_PRDummy.Primary_Currency = 1
    '                                    Row_AOP_PRDummy.Transaction_Amount = 0
    '                                    Row_AOP_PRDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                            End If
    '                        ElseIf ListPrimary.Contains(RowBal_PurcAutomaticPR.Primary_AOP) = True Then
    '                            If RowBal_PurcAutomaticPR.Primary_Currency = 1 Then
    '                                Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
    '                                If Row_AOP_PR.Tax_Currency = 0 Then
    '                                    'Input Pasangan USD yang IDR-Nya
    '                                    Row_AOP_PR = TBL_AOP_PR.Rows(TBL_AOP_PR.Rows.Count - 1)
    '                                    Row_AOP_PR.Transaction_Amount += RowBal_PurcAutomaticPR.balance_original + RowBal_PurcAutomaticPR.paid_original
    '                                    'Input Paid Date
    '                                    TBL_Pay.DefaultView.RowFilter = Nothing
    '                                    TBL_Pay.DefaultView.Sort = Nothing
    '                                    Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                                    TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_AOP_PR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_AOP_PR.Primary_Currency.ToString
    '                                    TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                                    If TBL_Pay.DefaultView.Count > 0 Then
    '                                        Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                                        Row_AOP_PR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            If mAbaikanHistorical = False Then
    '                If TBL_HisPurc.Rows.Count > 0 Then
    '                    For Each Row_HisPurc In TBL_HisPurc.Rows
    '                        TBLBal_PurcManual.DefaultView.RowFilter = Nothing
    '                        TBLBal_PurcManual.DefaultView.RowFilter = "[Primary AOP] = " & Row_HisPurc.Primary_Transaksi.ToString & " AND [Primary Supplier] = " & Row_HisPurc.Primary_Supplier.ToString & " AND [Primary Currency] = " & Row_HisPurc.Primary_Currency.ToString
    '                        If TBLBal_PurcManual.DefaultView.Count > 0 Then
    '                            For Each DV01 In TBLBal_PurcManual.DefaultView
    '                                RowBal_PurcManual = DV01.Row
    '                                Row_HisPurc.Transaction_Amount += RowBal_PurcManual.balance_original + RowBal_PurcManual.paid_original
    '                            Next
    '                        End If
    '                        TBL_Pay.DefaultView.RowFilter = Nothing
    '                        TBL_Pay.DefaultView.Sort = Nothing
    '                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
    '                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisPurc.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisPurc.Primary_Currency.ToString
    '                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                        If TBL_Pay.DefaultView.Count > 0 Then
    '                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                            Row_HisPurc.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                        End If
    '                    Next
    '                End If
    '                If TBL_HisPR.Rows.Count > 0 Then
    '                    For Each Row_HisPR In TBL_HisPR.Rows
    '                        TBLBal_PurcManualPR.DefaultView.RowFilter = Nothing
    '                        TBLBal_PurcManualPR.DefaultView.RowFilter = "[Primary AOP] = " & Row_HisPR.Primary_Transaksi.ToString & " AND [Primary Supplier] = " & Row_HisPR.Primary_Supplier.ToString & " AND [Primary Currency] = " & Row_HisPR.Primary_Currency.ToString
    '                        If TBLBal_PurcManualPR.DefaultView.Count > 0 Then
    '                            For Each DV01 In TBLBal_PurcManualPR.DefaultView
    '                                RowBal_PurcManualPR = DV01.Row
    '                                Row_HisPR.Transaction_Amount += RowBal_PurcManualPR.balance_original + RowBal_PurcManualPR.paid_original
    '                            Next
    '                        End If
    '                        TBL_Pay.DefaultView.RowFilter = Nothing
    '                        TBL_Pay.DefaultView.Sort = Nothing
    '                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisPR.Primary_Transaksi.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisPR.Primary_Currency.ToString
    '                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
    '                        If TBL_Pay.DefaultView.Count > 0 Then
    '                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
    '                            Row_HisPR.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
    '                        End If
    '                    Next
    '                End If
    '                '
    '            End If
    '            '***************
    '        End If
    '        'filter khusus currency 
    '        Me.FilterDataByPrimaryCurrency(TBLInv, TBL_SR, TBL_AOP, TBL_AOP_PR)
    '        '6.2. isikan tabel jurnal item payment untuk dapet paid date (udah diatas)
    '        '6.3. isikan cust po number utnuk invoice (udah diatas)
    '        '7. tabel saldo/saldo1 (udah diatas)
    '        '8. tentukan nilai primary min dan max untuk tabel invoice, sales return, historical sales, dan historical sales return (Udah diatas)
    '        '***************
    '        '9. buang baris dalam tabel saldo yang tidak sesuai dengan primary min dan primary max untuk setiap data
    '        '10. untuk setiap data dalam tabel saldo, pindahkan ke dalam tabel cetak
    '        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
    '            If TBLSaldo_AR.Rows.Count > 0 Then
    '                For Each RowSaldo_AR In TBLSaldo_AR.Rows
    '                    If RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
    '                        TBLInv.DefaultView.RowFilter = Nothing
    '                        TBLInv.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
    '                        If TBLInv.DefaultView.Count > 0 Then
    '                            RowInv = TBLInv.DefaultView(0).Row
    '                            RowDummy = TBLDummy.NewRow
    '                            RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                            RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                            RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                            RowDummy.Primary_Cust = RowInv.Primary_Customer
    '                            RowDummy.Customer_Name = RowInv.Customer_Name
    '                            RowDummy.Primary_Invoice = RowInv.Primary_Transaksi
    '                            RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                            RowDummy.Invoice_No = RowInv.No_Transaksi
    '                            RowDummy.Invoice_Date = RowInv.Tanggal_Transaksi
    '                            RowDummy._Cust__PO_No = RowInv.Cust_PO_Number
    '                            RowDummy.Payment_Term = RowInv.Payment_Term
    '                            RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
    '                            RowDummy.Transaction_Amount = RowInv.Transaction_Amount
    '                            RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                            RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                            If RowDummy.Payment_Amount > 0 Then
    '                                RowDummy.Paid_Date = RowInv.Paid_Date
    '                            Else
    '                                RowDummy.Paid_Date = "1-Jan-1950"
    '                            End If
    '                            TBLDummy.Rows.Add(RowDummy)
    '                        End If
    '                    ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
    '                        TBL_SR.DefaultView.RowFilter = Nothing
    '                        TBL_SR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AR.primary_currency.ToString
    '                        If TBL_SR.DefaultView.Count > 0 Then
    '                            Row_SR = TBL_SR.DefaultView(0).Row
    '                            RowDummy = TBLDummy.NewRow
    '                            RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                            RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                            RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                            RowDummy.Primary_Cust = Row_SR.Primary_Customer
    '                            RowDummy.Customer_Name = Row_SR.Customer_Name
    '                            RowDummy.Primary_Invoice = Row_SR.Primary_Transaksi
    '                            RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                            RowDummy.Invoice_No = Row_SR.No_Transaksi
    '                            RowDummy.Invoice_Date = Row_SR.Tanggal_Transaksi
    '                            RowDummy._Cust__PO_No = ""
    '                            RowDummy.Payment_Term = Row_SR.Payment_Term
    '                            RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
    '                            RowDummy.Transaction_Amount = Row_SR.Transaction_Amount
    '                            RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                            RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                            If RowDummy.Payment_Amount > 0 Then
    '                                RowDummy.Paid_Date = Row_SR.Paid_Date
    '                            Else
    '                                RowDummy.Paid_Date = "1-Jan-1950"
    '                            End If
    '                            TBLDummy.Rows.Add(RowDummy)
    '                        End If
    '                        'ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.DebitNote Then
    '                        '    TBL_NoteDebit.DefaultView.RowFilter = Nothing
    '                        '    TBL_NoteDebit.DefaultView.Sort = Nothing
    '                        '    TBL_NoteDebit.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [tipe transaksi] = " & Me.TipeTransaksiDebitNote
    '                        '    TBL_NoteDebit.DefaultView.Sort = "[Primary Transaksi]"
    '                        '    If TBL_NoteDebit.DefaultView.Count > 0 Then
    '                        '        Row_NoteDebit = TBL_NoteDebit.DefaultView(0).Row
    '                        '        RowDummy = TBLDummy.NewRow
    '                        '        RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                        '        RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                        '        RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                        '        RowDummy.Primary_Cust = RowSaldo_AR.primary_currency
    '                        '        RowDummy.Customer_Name = RowSaldo_AR.Customer_Name
    '                        '        RowDummy.Primary_Invoice = Row_NoteDebit.Primary_Transaksi
    '                        '        RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
    '                        '        RowDummy.Invoice_No = Row_NoteDebit.no_transksi
    '                        '        RowDummy.Invoice_Date = Row_NoteDebit.tgl
    '                        '        RowDummy._Cust__PO_No = ""
    '                        '        RowDummy.Payment_Term = Format(Row_NoteDebit.due_date - Row_NoteDebit.tgl, "dd-MMM-yyyy")
    '                        '        RowDummy.Due_Date = Format(Row_NoteDebit.due_date, "dd-MMM-yyyy")
    '                        '        RowDummy.Transaction_Amount = Row_NoteDebit.amount
    '                        '        RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                        '        RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                        '        If RowDummy.Payment_Amount > 0 Then
    '                        '            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
    '                        '            RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AR.Primary_Transaksi_Asal, RowSaldo_AR.primary_currency)
    '                        '        Else
    '                        '            RowDummy.Paid_Date = "1-Jan-1950"
    '                        '        End If
    '                        '        TBLDummy.Rows.Add(RowDummy)
    '                        '    End If
    '                        'ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.CreditNote Then
    '                        '    TBL_NoteKredit.DefaultView.RowFilter = Nothing
    '                        '    TBL_NoteKredit.DefaultView.Sort = Nothing
    '                        '    TBL_NoteKredit.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString & " AND [tipe transaksi] = " & Me.TipeTransaksiKreditNote
    '                        '    TBL_NoteKredit.DefaultView.Sort = "[Primary Transaksi]"
    '                        '    If TBL_NoteKredit.DefaultView.Count > 0 Then
    '                        '        Row_NoteKredit = TBL_NoteKredit.DefaultView(0).Row
    '                        '        RowDummy = TBLDummy.NewRow
    '                        '        RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                        '        RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                        '        RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                        '        RowDummy.Primary_Cust = RowSaldo_AR.primary_currency
    '                        '        RowDummy.Customer_Name = RowSaldo_AR.Customer_Name
    '                        '        RowDummy.Primary_Invoice = Row_NoteKredit.Primary_Transaksi
    '                        '        RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
    '                        '        RowDummy.Invoice_No = Row_NoteKredit.no_transksi
    '                        '        RowDummy.Invoice_Date = Row_NoteKredit.tgl
    '                        '        RowDummy._Cust__PO_No = ""
    '                        '        RowDummy.Payment_Term = Format(Row_NoteKredit.due_date - Row_NoteKredit.tgl, "dd-MMM-yyyy")
    '                        '        RowDummy.Due_Date = Format(Row_NoteKredit.due_date, "dd-MMM-yyyy")
    '                        '        RowDummy.Transaction_Amount = Row_NoteKredit.amount * -1
    '                        '        RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                        '        RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                        '        If RowDummy.Payment_Amount > 0 Then
    '                        '            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
    '                        '            RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AR.Primary_Transaksi_Asal, RowSaldo_AR.primary_currency)
    '                        '        Else
    '                        '            RowDummy.Paid_Date = "1-Jan-1950"
    '                        '        End If
    '                        '        TBLDummy.Rows.Add(RowDummy)
    '                        '    End If
    '                    End If
    '                    If mAbaikanHistorical = False Then
    '                        If RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
    '                            TBL_HisSales.DefaultView.RowFilter = Nothing
    '                            TBL_HisSales.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString
    '                            If TBL_HisSales.DefaultView.Count > 0 Then
    '                                Row_HisSales = TBL_HisSales.DefaultView(0).Row
    '                                RowDummy = TBLDummy.NewRow
    '                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                                RowDummy.Primary_Cust = Row_HisSales.Primary_Customer
    '                                RowDummy.Customer_Name = Row_HisSales.Customer_Name
    '                                RowDummy.Primary_Invoice = Row_HisSales.Primary_Transaksi
    '                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
    '                                RowDummy.Invoice_No = Row_HisSales.Invoice_Number
    '                                RowDummy.Invoice_Date = Row_HisSales.Invoice_Date
    '                                RowDummy._Cust__PO_No = ""
    '                                RowDummy.Payment_Term = Format(Row_HisSales.Due_Date - Row_HisSales.Invoice_Date, "dd-MMM-yyyy")
    '                                RowDummy.Due_Date = Row_HisSales.Due_Date
    '                                RowDummy.Transaction_Amount = Row_HisSales.Transaction_Amount
    '                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                                If RowDummy.Payment_Amount > 0 Then
    '                                    RowDummy.Paid_Date = Row_HisSales.Paid_Date
    '                                Else
    '                                    RowDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                                TBLDummy.Rows.Add(RowDummy)
    '                            End If
    '                        ElseIf RowSaldo_AR.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
    '                            TBL_HisSR.DefaultView.RowFilter = Nothing
    '                            TBL_HisSR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AR.Primary_Transaksi_Asal.ToString
    '                            If TBL_HisSR.DefaultView.Count > 0 Then
    '                                Row_HisSR = TBL_HisSR.DefaultView(0).Row
    '                                RowDummy = TBLDummy.NewRow
    '                                RowDummy.Currency_Code = RowSaldo_AR.Currency_Code
    '                                RowDummy.Currency_Name = RowSaldo_AR.Currency_Name
    '                                RowDummy.DefPrec_Currency = RowSaldo_AR.Default_Precision
    '                                RowDummy.Primary_Cust = Row_HisSR.Primary_Customer
    '                                RowDummy.Customer_Name = Row_HisSR.Customer_Name
    '                                RowDummy.Primary_Invoice = Row_HisSR.Primary_Transaksi
    '                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '                                RowDummy.Invoice_No = Row_HisSR.AOSR_Number
    '                                RowDummy.Invoice_Date = Row_HisSR.AOSR_Date
    '                                RowDummy._Cust__PO_No = ""
    '                                RowDummy.Payment_Term = Format(Row_HisSR.Due_Date - Row_HisSR.AOSR_Date, "dd-MMM-yyyy")
    '                                RowDummy.Due_Date = Row_HisSR.Due_Date
    '                                RowDummy.Transaction_Amount = Row_HisSR.Transaction_Amount
    '                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AR.Balance_Original
    '                                RowDummy.Outstanding = RowSaldo_AR.Balance_Original
    '                                If RowDummy.Payment_Amount > 0 Then
    '                                    RowDummy.Paid_Date = Row_HisSR.Paid_Date
    '                                Else
    '                                    RowDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                                TBLDummy.Rows.Add(RowDummy)
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
    '            If TBLSaldo_AP.Rows.Count > 0 Then
    '                For Each RowSaldo_AP In TBLSaldo_AP.Rows
    '                    If RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
    '                        TBL_AOP.DefaultView.RowFilter = Nothing
    '                        TBL_AOP.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
    '                        If TBL_AOP.DefaultView.Count > 0 Then
    '                            Row_AOP = TBL_AOP.DefaultView(0).Row
    '                            RowDummy = TBLDummy.NewRow
    '                            RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                            RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                            RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                            RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
    '                            RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                            RowDummy.Primary_Invoice = Row_AOP.Primary_Transaksi
    '                            RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    '                            RowDummy.Invoice_No = Row_AOP.No_Transaksi
    '                            RowDummy.Invoice_Date = Row_AOP.Tanggal_Transaksi
    '                            RowDummy._Cust__PO_No = Row_AOP.PO_Number
    '                            RowDummy.Payment_Term = Row_AOP.Payment_Term
    '                            RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
    '                            RowDummy.Transaction_Amount = Row_AOP.Transaction_Amount
    '                            RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                            RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                            If RowDummy.Payment_Amount > 0 Then
    '                                RowDummy.Paid_Date = Row_AOP.Paid_Date
    '                            Else
    '                                RowDummy.Paid_Date = "1-Jan-1950"
    '                            End If
    '                            TBLDummy.Rows.Add(RowDummy)
    '                        End If
    '                    ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
    '                        TBL_AOP_PR.DefaultView.RowFilter = Nothing
    '                        TBL_AOP_PR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & " AND [Primary Currency] = " & RowSaldo_AP.primary_currency.ToString
    '                        If TBL_AOP_PR.DefaultView.Count > 0 Then
    '                            Row_AOP_PR = TBL_AOP_PR.DefaultView(0).Row
    '                            RowDummy = TBLDummy.NewRow
    '                            RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                            RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                            RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                            RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
    '                            RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                            RowDummy.Primary_Invoice = Row_AOP_PR.Primary_Transaksi
    '                            RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
    '                            RowDummy.Invoice_No = Row_AOP_PR.No_Transaksi
    '                            RowDummy.Invoice_Date = Row_AOP_PR.Tanggal_Transaksi
    '                            RowDummy._Cust__PO_No = ""
    '                            RowDummy.Payment_Term = Row_AOP_PR.Payment_Term
    '                            RowDummy.Due_Date = RowDummy.Invoice_Date + TimeSpan.FromDays(CDbl(RowDummy.Payment_Term))
    '                            RowDummy.Transaction_Amount = Row_AOP_PR.Transaction_Amount
    '                            RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                            RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                            If RowDummy.Payment_Amount > 0 Then
    '                                RowDummy.Paid_Date = Row_AOP_PR.Paid_Date
    '                            Else
    '                                RowDummy.Paid_Date = "1-Jan-1950"
    '                            End If
    '                            TBLDummy.Rows.Add(RowDummy)
    '                        End If
    '                        'ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.DebitNote Then
    '                        '    TBL_NoteDebit.DefaultView.RowFilter = Nothing
    '                        '    TBL_NoteDebit.DefaultView.Sort = Nothing
    '                        '    TBL_NoteDebit.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & " AND [tipe transaksi] = " & Me.TipeTransaksiDebitNote
    '                        '    TBL_NoteDebit.DefaultView.Sort = "[Primary Transaksi]"
    '                        '    If TBL_NoteDebit.DefaultView.Count > 0 Then
    '                        '        Row_NoteDebit = TBL_NoteDebit.DefaultView(0).Row
    '                        '        RowDummy = TBLDummy.NewRow
    '                        '        RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                        '        RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                        '        RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                        '        RowDummy.Primary_Cust = RowSaldo_AP.primary_currency
    '                        '        RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                        '        RowDummy.Primary_Invoice = Row_NoteDebit.Primary_Transaksi
    '                        '        RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
    '                        '        RowDummy.Invoice_No = Row_NoteDebit.no_transksi
    '                        '        RowDummy.Invoice_Date = Row_NoteDebit.tgl
    '                        '        RowDummy._Cust__PO_No = ""
    '                        '        RowDummy.Payment_Term = Format(Row_NoteDebit.due_date - Row_NoteDebit.tgl, "dd-MMM-yyyy")
    '                        '        RowDummy.Due_Date = Format(Row_NoteDebit.due_date, "dd-MMM-yyyy")
    '                        '        RowDummy.Transaction_Amount = Row_NoteDebit.amount * -1
    '                        '        RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                        '        RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                        '        If RowDummy.Payment_Amount > 0 Then
    '                        '            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
    '                        '            RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AP.Primary_Transaksi_Asal, RowSaldo_AP.primary_currency)
    '                        '        Else
    '                        '            RowDummy.Paid_Date = "1-Jan-1950"
    '                        '        End If
    '                        '        TBLDummy.Rows.Add(RowDummy)
    '                        '    End If
    '                        'ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.CreditNote Then
    '                        '    TBL_NoteKredit.DefaultView.RowFilter = Nothing
    '                        '    TBL_NoteKredit.DefaultView.Sort = Nothing
    '                        '    TBL_NoteKredit.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString & " AND [tipe transaksi] = " & Me.TipeTransaksiKreditNote
    '                        '    TBL_NoteKredit.DefaultView.Sort = "[Primary Transaksi]"
    '                        '    If TBL_NoteKredit.DefaultView.Count > 0 Then
    '                        '        Row_NoteKredit = TBL_NoteKredit.DefaultView(0).Row
    '                        '        RowDummy = TBLDummy.NewRow
    '                        '        RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                        '        RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                        '        RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                        '        RowDummy.Primary_Cust = RowSaldo_AP.primary_currency
    '                        '        RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                        '        RowDummy.Primary_Invoice = Row_NoteKredit.Primary_Transaksi
    '                        '        RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
    '                        '        RowDummy.Invoice_No = Row_NoteKredit.no_transksi
    '                        '        RowDummy.Invoice_Date = Row_NoteKredit.tgl
    '                        '        RowDummy._Cust__PO_No = ""
    '                        '        RowDummy.Payment_Term = Format(Row_NoteKredit.due_date - Row_NoteKredit.tgl, "dd-MMM-yyyy")
    '                        '        RowDummy.Due_Date = Format(Row_NoteKredit.due_date, "dd-MMM-yyyy")
    '                        '        RowDummy.Transaction_Amount = Row_NoteKredit.amount
    '                        '        RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                        '        RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                        '        If RowDummy.Payment_Amount > 0 Then
    '                        '            Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
    '                        '            RowDummy.Paid_Date = Me.GetPAIDDate(TBL_Pay, Me.TipeTransaksiAsal, RowSaldo_AP.Primary_Transaksi_Asal, RowSaldo_AP.primary_currency)
    '                        '        Else
    '                        '            RowDummy.Paid_Date = "1-Jan-1950"
    '                        '        End If
    '                        '        TBLDummy.Rows.Add(RowDummy)
    '                        '    End If
    '                    End If
    '                    If mAbaikanHistorical = False Then
    '                        If RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Then
    '                            TBL_HisPurc.DefaultView.RowFilter = Nothing
    '                            TBL_HisPurc.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString
    '                            If TBL_HisPurc.DefaultView.Count > 0 Then
    '                                Row_HisPurc = TBL_HisPurc.DefaultView(0).Row
    '                                RowDummy = TBLDummy.NewRow
    '                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                                RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
    '                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                                RowDummy.Primary_Invoice = Row_HisPurc.Primary_Transaksi
    '                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
    '                                RowDummy.Invoice_No = Row_HisPurc.AOP_Number
    '                                RowDummy.Invoice_Date = Row_HisPurc.AOP_Date
    '                                RowDummy._Cust__PO_No = ""
    '                                RowDummy.Payment_Term = Format(Row_HisPurc.Due_Date - Row_HisPurc.AOP_Date, "dd-MMM-yyyy")
    '                                RowDummy.Due_Date = Row_HisPurc.Due_Date
    '                                RowDummy.Transaction_Amount = Row_HisPurc.Transaction_Amount
    '                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                                If RowDummy.Payment_Amount > 0 Then
    '                                    RowDummy.Paid_Date = Row_HisPurc.Paid_Date
    '                                Else
    '                                    RowDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                                TBLDummy.Rows.Add(RowDummy)
    '                            End If
    '                        ElseIf RowSaldo_AP.Type_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
    '                            TBL_HisPR.DefaultView.RowFilter = Nothing
    '                            TBL_HisPR.DefaultView.RowFilter = "[Primary Transaksi] = " & RowSaldo_AP.Primary_Transaksi_Asal.ToString
    '                            If TBL_HisPR.DefaultView.Count > 0 Then
    '                                Row_HisPR = TBL_HisPR.DefaultView(0).Row
    '                                RowDummy = TBLDummy.NewRow
    '                                RowDummy.Currency_Code = RowSaldo_AP.Currency_Code
    '                                RowDummy.Currency_Name = RowSaldo_AP.Currency_Name
    '                                RowDummy.DefPrec_Currency = RowSaldo_AP.Default_Precision
    '                                RowDummy.Primary_Cust = RowSaldo_AP.Primary_Supplier
    '                                RowDummy.Customer_Name = RowSaldo_AP.Supplier_Name
    '                                RowDummy.Primary_Invoice = Row_HisPR.Primary_Transaksi
    '                                RowDummy.Tipe_Dokumen = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
    '                                RowDummy.Invoice_No = Row_HisPR.AOPR_Number
    '                                RowDummy.Invoice_Date = Row_HisPR.AOPR_Date
    '                                RowDummy._Cust__PO_No = ""
    '                                RowDummy.Payment_Term = Format(Row_HisPR.Due_Date - Row_HisPR.AOPR_Date, "dd-MMM-yyyy")
    '                                RowDummy.Due_Date = Row_HisPR.Due_Date
    '                                RowDummy.Transaction_Amount = Row_HisPR.Transaction_Amount
    '                                RowDummy.Payment_Amount = RowDummy.Transaction_Amount - RowSaldo_AP.Balance_Original
    '                                RowDummy.Outstanding = RowSaldo_AP.Balance_Original
    '                                If RowDummy.Payment_Amount > 0 Then
    '                                    RowDummy.Paid_Date = Row_HisPR.Paid_Date
    '                                Else
    '                                    RowDummy.Paid_Date = "1-Jan-1950"
    '                                End If
    '                                TBLDummy.Rows.Add(RowDummy)
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        '***************
    '        '11. tampilkan ke dalam crv
    '        TBLDummy.DefaultView.Sort = Nothing
    '        TBLDummy.DefaultView.Sort = "[Currency Code], [Customer Name], [Invoice Date], [Invoice No]"
    '        If TBLDummy.DefaultView.Count > 0 Then
    '            For Each DV01 In TBLDummy.DefaultView
    '                RowDummy = DV01.Row
    '                If ListPrimaryCurr.Contains(RowDummy.Currency_Code) = False Then
    '                    ListPrimaryCurr.Add(RowDummy.Currency_Code)
    '                    RowCetakTotal = TBLCetakTotal.NewRow
    '                    RowCetakTotal.CurrCode = RowDummy.Currency_Code
    '                    RowCetakTotal.CurrName = RowDummy.Currency_Name
    '                    RowCetakTotal.TotalAmount = RowDummy.Transaction_Amount
    '                    RowCetakTotal.TotalPayment = RowDummy.Payment_Amount
    '                    RowCetakTotal.TotalOutstanding = RowDummy.Outstanding
    '                    RowCetakTotal.DefPrec = RowDummy.DefPrec_Currency
    '                    TBLCetakTotal.Rows.Add(RowCetakTotal)
    '                Else
    '                    TBLCetakTotal.DefaultView.RowFilter = Nothing
    '                    TBLCetakTotal.DefaultView.RowFilter = "[CurrCode] = '" & RowDummy.Currency_Code.ToString & "'"
    '                    If TBLCetakTotal.DefaultView.Count > 0 Then
    '                        RowCetakTotal = TBLCetakTotal.DefaultView(0).Row
    '                        RowCetakTotal.TotalAmount += RowDummy.Transaction_Amount
    '                        RowCetakTotal.TotalPayment += RowDummy.Payment_Amount
    '                        RowCetakTotal.TotalOutstanding += RowDummy.Outstanding
    '                    End If
    '                End If
    '                RowCetak = TBLCetak.NewRow
    '                For i As Integer = 0 To TBLCetak.Columns.Count - 1
    '                    RowCetak.Item(i) = RowDummy.Item(i)
    '                Next
    '                TBLCetak.Rows.Add(RowCetak)
    '            Next
    '        End If

    '        Report_AR_AP.SetDataSource(DS)
    '        Report_AR_AP.Subreports(0).SetDataSource(DS)
    '        Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamHeader.ParameterFieldName, mJudul)
    '        Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamGroup.ParameterFieldName, mJudulKolom)
    '        Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamStartDate.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd-MMM-yyyy"))
    '        Report_AR_AP.SetParameterValue(Report_AR_AP.Parameter_ParamEndDate.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd-MMM-yyyy"))
    '        Me.crvAP.ReportSource = Report_AR_AP
    '        Me.crvAP.Show()
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ListPrimarySOPONumber = Nothing
    '        ListPrimaryCurr = Nothing
    '        ListPrimary = Nothing
    '        DTABal_SalesAutomatic.Dispose()
    '        TBLBal_SalesAutomatic.Dispose()
    '        DTABal_SalesAutomaticSR.Dispose()
    '        TBLBal_SalesAutomaticSR.Dispose()
    '        DTABal_SalesManual.Dispose()
    '        TBLBal_SalesManual.Dispose()
    '        DTABal_SalesManualSR.Dispose()
    '        TBLBal_SalesManualSR.Dispose()
    '        'DTABal_SalesDebitNote.Dispose()
    '        'TBLBal_SalesDebitNote.Dispose()
    '        'DTABal_SalesKreditNote.Dispose()
    '        'TBLBal_SalesKreditNote.Dispose()
    '        DTABal_PurcAutomatic.Dispose()
    '        TBLBal_PurcAutomatic.Dispose()
    '        DTABal_PurcAutomaticPR.Dispose()
    '        TBLBal_PurcAutomaticPR.Dispose()
    '        DTABal_PurcManual.Dispose()
    '        TBLBal_PurcManual.Dispose()
    '        DTABal_PurcManualPR.Dispose()
    '        TBLBal_PurcManualPR.Dispose()
    '        'DTABal_PurcDebitNote.Dispose()
    '        'TBLBal_PurcDebitNote.Dispose()
    '        'DTABal_PurcKreditNote.Dispose()
    '        'TBLBal_PurcKreditNote.Dispose()
    '        DS.Dispose()
    '        DTA_Pay.Dispose()
    '        TBL_Pay.Dispose()
    '        'DTA_NoteDebit.Dispose()
    '        'TBL_NoteDebit.Dispose()
    '        'DTA_NoteKredit.Dispose()
    '        'TBL_NoteKredit.Dispose()
    '        DTASaldo_AR.Dispose()
    '        TBLSaldo_AR.Dispose()
    '        DTAInv.Dispose()
    '        TBLInv.Dispose()
    '        DTAInvItem.Dispose()
    '        TBLInvItem.Dispose()
    '        DTA_SR.Dispose()
    '        TBL_SR.Dispose()
    '        DTA_SRItem.Dispose()
    '        TBL_SRItem.Dispose()
    '        DTA_HisSales.Dispose()
    '        TBL_HisSales.Dispose()
    '        DTA_HisSR.Dispose()
    '        TBL_HisSR.Dispose()
    '        DTA_SO.Dispose()
    '        TBL_SO.Dispose()
    '        DTASaldo_AP.Dispose()
    '        TBLSaldo_AP.Dispose()
    '        DTA_AOP.Dispose()
    '        TBL_AOP.Dispose()
    '        DTA_AOP_Item.Dispose()
    '        TBL_AOP_Item.Dispose()
    '        DTA_AOP_PR.Dispose()
    '        TBL_AOP_PR.Dispose()
    '        DTA_AOP_PR_Item.Dispose()
    '        TBL_AOP_PR_Item.Dispose()
    '        DTA_HisPurc.Dispose()
    '        TBL_HisPurc.Dispose()
    '        DTA_HisPR.Dispose()
    '        TBL_HisPR.Dispose()
    '        DTA_PO.Dispose()
    '        TBL_PO.Dispose()
    '        TBLDummy.Dispose()
    '    End Try
    'End Sub

    Private Function GetPAIDDate(ByVal TBL_Pay As DseLaporanARPaid.Item_PaymentDataTable, _
    ByVal TipeTransaksiAsal As ClsTransaksiSalesBalance.EnumTipeSalesBalance, ByVal Primary_Transaksi_Acuan As Double, _
    ByVal Primary_Currency As Double) As String
        Dim Retval As String = "1-Jan-1950"
        Dim Row_Pay As DseLaporanARPaid.Item_PaymentRow
        Try
            TBL_Pay.DefaultView.RowFilter = Nothing
            TBL_Pay.DefaultView.Sort = Nothing
            TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Primary_Transaksi_Acuan.ToString _
            & " AND [Tipe Document Acuan] = " & TipeTransaksiAsal & " AND [Primary Currency] = " & Primary_Currency.ToString
            TBL_Pay.DefaultView.Sort = "[Primary Journal]"
            If TBL_Pay.DefaultView.Count > 0 Then
                Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                Retval = Format(Row_Pay._Date, "dd-MMM-yyyy")
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1-Jan-1950"
        End Try
    End Function

    Private Sub FrmLaporanARPaid_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.BersihkanObjek()
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Me.chkUseSupplieInvoiceNo.Visible = False
                Me.chkUseSupplieInvoiceNo.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'FindCustomer
    Private Sub txtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyDown
        Dim objCode As New clsMasterCustomer
        Dim i As New FrmCariMasterCustomer
        Try
            If e.KeyValue = 13 Then
                If Me.txtCustCode.Text <> "" Then
                    If Me.ObjCust.ApakahSudahAda(Me.txtCustCode.Text) = True Then
                        Me.txtPrimaryCust.Text = Me.ObjCust.GetPrimary(Me.txtCustCode.Text)
                        Me.txtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.txtPrimaryCust.Text)
                        Me.txtCustDesc.Text = Me.ObjCust.GetCustName(Me.txtPrimaryCust.Text)
                        Me.txtArAccountCode.Focus()
                    Else
                        Me.btnFindCust.PerformClick()
                    End If
                Else
                    Me.btnFindCust.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.Leave
        Dim objCode As New clsMasterCustomer
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustCode.Text <> "" Then
                If Me.ObjCust.ApakahSudahAda(Me.txtCustCode.Text) = True Then
                    Me.txtPrimaryCust.Text = Me.ObjCust.GetPrimary(Me.txtCustCode.Text)
                    Me.txtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.txtPrimaryCust.Text)
                    Me.txtCustDesc.Text = Me.ObjCust.GetCustName(Me.txtPrimaryCust.Text)
                Else
                    If Me.txtCustCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCust.Text = i.PrimaryCustomer
                        Me.txtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.txtPrimaryCust.Text)
                        Me.txtCustDesc.Text = Me.ObjCust.GetCustName(Me.txtPrimaryCust.Text)
                    Else
                        Me.txtPrimaryCust.Text = "-1"
                        Me.txtCustCode.Clear()
                        Me.txtCustDesc.Clear()
                    End If
                End If
            Else
                Me.txtCustCode.Clear()
                Me.txtPrimaryCust.Text = "-1"
                Me.txtCustDesc.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCust_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCust.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCust.Text = i.PrimaryCustomer
                Me.txtCustCode.Text = Me.ObjCust.GetCustomerCode(Me.txtPrimaryCust.Text)
                Me.txtCustDesc.Text = Me.ObjCust.GetCustName(Me.txtPrimaryCust.Text)
                Me.txtArAccountCode.Focus()
            Else
                Me.txtPrimaryCust.Text = "-1"
                Me.txtCustCode.Focus()
                Me.txtCustDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindAccount
    Private Sub txtArAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtArAccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.txtArAccountCode.Text <> "" Then
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.txtArAccountCode.Text) = True Then
                        Me.txtPrimaryArAccount.Text = Me.ObjAkun.GetPrimary(Me.txtArAccountCode.Text)
                        Me.txtArAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.txtPrimaryArAccount.Text)
                        Me.txtArAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                        Me.txtCurrencyCode.Focus()
                    Else
                        Me.btnFindArAccount.PerformClick()
                    End If
                Else
                    Me.btnFindArAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub txtArAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtArAccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtArAccountCode.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.txtArAccountCode.Text) = True Then
                    Me.txtPrimaryArAccount.Text = Me.ObjAkun.GetPrimary(Me.txtArAccountCode.Text)
                    Me.txtArAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.txtPrimaryArAccount.Text)
                    Me.txtArAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                Else
                    If Me.txtArAccountCode.Text <> "" Then
                        i.txtCari.Text = Me.txtArAccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.txtPrimaryArAccount.Text = i.mPrimaryAkun
                        Me.txtArAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.txtPrimaryArAccount.Text)
                        Me.txtArAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                    Else
                        Me.txtPrimaryArAccount.Text = "-1"
                        Me.txtArAccountCode.Clear()
                        Me.txtArAccountDesc.Clear()
                    End If
                End If
            Else
                Me.txtArAccountCode.Clear()
                Me.txtPrimaryArAccount.Text = "-1"
                Me.txtArAccountDesc.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindArAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindArAccount.Click
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtArAccountCode.Text <> "" Then
                i.txtCari.Text = Me.txtArAccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.txtPrimaryArAccount.Text = i.mPrimaryAkun
                Me.txtArAccountCode.Text = Me.ObjAkun.GetAccountCode(Me.txtPrimaryArAccount.Text)
                Me.txtArAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.txtPrimaryArAccount.Text)
                Me.txtCurrencyCode.Focus()
            Else
                Me.txtPrimaryArAccount.Text = "-1"
                Me.txtArAccountCode.Focus()
                Me.txtArAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindCurrency
    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim objCode As New clsMasterCurrency
        Dim i As New FrmCariMasterCurrency
        Try
            If e.KeyValue = 13 Then
                If Me.txtCurrencyCode.Text <> "" Then
                    If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                        Me.txtPrimaryCurrency.Text = Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text)
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.txtCurrencyDesc.Text = Me.ObjCurr.GetDescription(Me.txtPrimaryCurrency.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.btnFindCurrency.PerformClick()
                    End If
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim objCode As New clsMasterCurrency
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.txtCurrencyDesc.Text = Me.ObjCurr.GetDescription(Me.txtPrimaryCurrency.Text)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.txtCurrencyDesc.Text = Me.ObjCurr.GetDescription(Me.txtPrimaryCurrency.Text)
                    Else
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Clear()
                        Me.txtCurrencyDesc.Clear()
                    End If
                End If
            Else
                Me.txtCurrencyCode.Clear()
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyDesc.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                i.txtCari.Text = Me.txtCurrencyCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCurrency > 0 Then
                Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.txtCurrencyDesc.Text = Me.ObjCurr.GetDescription(Me.txtPrimaryCurrency.Text)
                Me.TxtCodeNo.Focus()
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Focus()
                Me.txtCurrencyDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindProduct
    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim objCode As New clsMasterProduct
        Dim i As New frmCariMasterProduct
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeNo.Text <> "" Then
                    If Me.ObjProd.ApakahSudahAda(Me.TxtCodeNo.Text) = True Then
                        Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
                        Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtSONumber.Focus()
                    Else
                        Me.btnFindCodeNo.PerformClick()
                    End If
                Else
                    Me.btnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim objCode As New clsMasterProduct
        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProd.ApakahSudahAda(Me.TxtCodeNo.Text) = True Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
                    Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Else
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtCodeNo.Clear()
                        Me.TxtDesc.Clear()
                    End If
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtDesc.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.TxtSONumber.Focus()
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNo.Focus()
                Me.TxtDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindSalesOrder
    Private Function GetCustPO(ByVal vPrimarySO As Double) As String
        Dim DTAItem As New DseLaporanARPaidTableAdapters.tbltrans_soTableAdapter
        Dim TBLItem As New DseLaporanARPaid.tbltrans_soDataTable
        Dim RowItem As DseLaporanARPaid.tbltrans_soRow
        Dim Retval As String = ""
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.FillBy_PrimarySO(TBLItem, vPrimarySO)
            If TBLItem.Rows.Count > 0 Then
                For Each RowItem In TBLItem.Rows
                    If String.IsNullOrEmpty(Retval) = True Then
                        Retval = RowItem.Cust_PO_No
                    Else
                        Retval += ", " + RowItem.Cust_PO_No
                    End If
                Next
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Private Sub TxtSONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSONumber.KeyDown
        Dim objCode As New ClsTransaksiSalesOrder
        Dim i As New FrmCariTransaksiSO
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSONumber.Text <> "" Then
                    If Me.ObjSO.ApakahSudahAda(Me.TxtSONumber.Text) = True Then
                        Me.TxtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtSONumber.Text)
                        Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                        Me.TxtCustomerPONumber.Text = Me.GetCustPO(Me.TxtPrimarySO.Text)
                        Me.btnShow.Focus()
                    Else
                        Me.btnFindSO.PerformClick()
                    End If
                Else
                    Me.btnFindSO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.Leave
        Dim objCode As New ClsTransaksiSalesOrder
        Dim i As New FrmCariTransaksiSO
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSONumber.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.TxtSONumber.Text) = True Then
                    Me.TxtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtSONumber.Text)
                    Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                    Me.TxtCustomerPONumber.Text = Me.GetCustPO(Me.TxtPrimarySO.Text)
                Else
                    If Me.TxtSONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtSONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtPrimarySO.Text = i.PrimaryTransaksi
                        Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                        Me.TxtCustomerPONumber.Text = Me.GetCustPO(Me.TxtPrimarySO.Text)
                    Else
                        Me.TxtPrimarySO.Text = "-1"
                        Me.TxtSONumber.Clear()
                        Me.TxtCustomerPONumber.Clear()
                    End If
                End If
            Else
                Me.TxtSONumber.Clear()
                Me.TxtPrimarySO.Text = "-1"
                Me.TxtCustomerPONumber.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindSO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSO.Click
        Dim i As New FrmCariTransaksiSO
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtSONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.TxtPrimarySO.Text = i.PrimaryTransaksi
                Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                Me.TxtCustomerPONumber.Text = Me.GetCustPO(Me.TxtPrimarySO.Text)
                Me.btnShow.Focus()
            Else
                Me.TxtPrimarySO.Text = "-1"
                Me.TxtSONumber.Focus()
                Me.TxtCustomerPONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindSupplier
    Private Sub TxtSuppCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSuppCode.KeyDown
        Dim objCode As New clsMasterSupplier
        Dim i As New FrmCariMasterSupplier
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSuppCode.Text <> "" Then
                    If Me.ObjSupp.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                        Me.TxtPrimarySupplier.Text = Me.ObjSupp.GetPrimary(Me.TxtSuppCode.Text)
                        Me.TxtSuppCode.Text = Me.ObjSupp.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupp.GetSupplierName(Me.TxtPrimarySupplier.Text)
                        Me.txtArAccountCode.Focus()
                    Else
                        Me.btnFindSupplier.PerformClick()
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSuppCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppCode.Leave
        Dim objCode As New clsMasterSupplier
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                If Me.ObjSupp.ApakahSudahAda(Me.TxtSuppCode.Text) = True Then
                    Me.TxtPrimarySupplier.Text = Me.ObjSupp.GetPrimary(Me.TxtSuppCode.Text)
                    Me.TxtSuppCode.Text = Me.ObjSupp.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = Me.ObjSupp.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Else
                    If Me.TxtSuppCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSuppCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                        Me.TxtSuppCode.Text = Me.ObjSupp.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupp.GetSupplierName(Me.TxtPrimarySupplier.Text)
                    Else
                        Me.TxtPrimarySupplier.Text = "-1"
                        Me.TxtSuppCode.Clear()
                        Me.TxtSupplierName.Clear()
                    End If
                End If
            Else
                Me.TxtSuppCode.Clear()
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSupplierName.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSuppCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSuppCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimarySupplier.Text = i.PrimarySupplier
                Me.TxtSuppCode.Text = Me.ObjSupp.GetSupplierCode(Me.TxtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupp.GetSupplierName(Me.TxtPrimarySupplier.Text)
                Me.txtArAccountCode.Focus()
            Else
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSuppCode.Focus()
                Me.TxtSupplierName.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'FindPO
    Private Sub TxtPONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPONumber.KeyDown
        Dim objCode As New ClsTransaksiPO
        Dim i As New FrmCariTransaksiPO
        Try
            If e.KeyValue = 13 Then
                If Me.TxtPONumber.Text <> "" Then
                    If Me.ObjPO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                        Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtPONumber.Text)
                        Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                        Me.btnShow.Focus()
                    Else
                        Me.btnFindPO.PerformClick()
                    End If
                Else
                    Me.btnFindPO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtPONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.Leave
        Dim objCode As New ClsTransaksiPO
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPONumber.Text <> "" Then
                If Me.ObjPO.ApakahSudahAda(Me.TxtPONumber.Text) = True Then
                    Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtPONumber.Text)
                    Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Else
                    If Me.TxtPONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtPONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        Me.TxtPrimaryPO.Text = i.PrimaryPO
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtPONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.TxtPrimaryPO.Text = i.PrimaryPO
                Me.TxtPONumber.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Me.btnShow.Focus()
            Else
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtPONumber.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    'aul add
    Private Sub GetTransaksiAmountHistoricalDP(ByVal vTipeHisDP As ClsTransaksiBalanceManual_DP.EnumTipeDP)
        Dim DV01 As Data.DataRowView
        Try
            If vTipeHisDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                If TBL_HisPurchDP.Rows.Count > 0 Then
                    For Each Row_HisPurchDP In TBL_HisPurchDP.Rows
                        TBLBal_PurcManual.DefaultView.RowFilter = Nothing
                        TBLBal_PurcManual.DefaultView.RowFilter = "[Primary AOP] = " & Row_HisPurchDP.Primary.ToString & " AND [Primary Supplier] = " & Row_HisPurchDP.Primary_Pihak_Acuan.ToString & " AND [Primary Currency] = " & Row_HisPurchDP.primary_currency.ToString
                        If TBLBal_PurcManual.DefaultView.Count > 0 Then
                            For Each DV01 In TBLBal_PurcManual.DefaultView
                                RowBal_PurcManual = DV01.Row
                                Row_HisPurchDP.Transaction_Amount += RowBal_PurcManual.balance_original + RowBal_PurcManual.paid_original
                            Next
                        End If
                        TBL_Pay.DefaultView.RowFilter = Nothing
                        TBL_Pay.DefaultView.Sort = Nothing
                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisPurchDP.Primary.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisPurchDP.primary_currency.ToString
                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                        If TBL_Pay.DefaultView.Count > 0 Then
                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                            Row_HisPurchDP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                        End If
                    Next
                End If
            ElseIf vTipeHisDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice Then
                If TBL_HisSalesDP.Rows.Count > 0 Then
                    For Each Row_HisSalesDP In TBL_HisSalesDP.Rows
                        TBLBal_SalesManual.DefaultView.RowFilter = Nothing
                        TBLBal_SalesManual.DefaultView.RowFilter = "[Primary Invoice] = " & Row_HisSalesDP.Primary.ToString & " AND [Primary Customer] = " & Row_HisSalesDP.Primary_Pihak_Acuan.ToString & " AND [Primary Currency] = " & Row_HisSalesDP.primary_currency.ToString
                        If TBLBal_SalesManual.DefaultView.Count > 0 Then
                            For Each DV01 In TBLBal_SalesManual.DefaultView
                                RowBal_SalesManual = DV01.Row
                                Row_HisSalesDP.Transaction_Amount += RowBal_SalesManual.balance_original + RowBal_SalesManual.received_original
                            Next
                        End If
                        TBL_Pay.DefaultView.RowFilter = Nothing
                        TBL_Pay.DefaultView.Sort = Nothing
                        Me.TipeTransaksiAsal = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP
                        TBL_Pay.DefaultView.RowFilter = "[Primary Document Acuan] = " & Row_HisSalesDP.Primary.ToString & " AND [Tipe Document Acuan] = " & Me.TipeTransaksiAsal & " AND [Primary Currency] = " & Row_HisSalesDP.primary_currency.ToString
                        TBL_Pay.DefaultView.Sort = "[Primary Journal]"
                        If TBL_Pay.DefaultView.Count > 0 Then
                            Row_Pay = TBL_Pay.DefaultView(TBL_Pay.DefaultView.Count - 1).Row
                            Row_HisSalesDP.Paid_Date = Format(Row_Pay._Date, "dd-MMM-yyyy")
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class