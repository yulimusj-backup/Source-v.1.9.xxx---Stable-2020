Option Explicit On
Option Strict Off

Imports AccModule
Public Class FrmUpdate
    Dim ObjPO As New Purchase_Module.ClsTransaksiPOItem
    Dim DTAPO As New DseUpdatePOBTableAdapters.tbltrans_poTableAdapter
    Dim TBLPO As New DseUpdatePOB.tbltrans_poDataTable
    Dim rowPO As DseUpdatePOB.tbltrans_poRow
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim TBLInv As New DseUpdate_v2018.tbltrans_salesinvoiceDataTable

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdatePO()
        Try
            DTAPO.Fill(TBLPO)
            Me.ProgPOB.Value = 0
            Me.ProgPOB.Maximum = TBLPO.Rows.Count

            For Each rowPO In TBLPO.Rows
                Me.ProgPOB.Value += 1
                Me.Text = "Update PO: " + rowPO.No_Transaksi + " (Data " + Me.ProgPOB.Value.ToString + " of " + Me.ProgPOB.Maximum.ToString + ")"
                ObjPO.SetOrderBalance(rowPO.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.UpdatePO()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("PO Update Finish", MsgBoxStyle.Information, "Information")
        Me.Text = "Update Utility"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateSOB()
        Dim DTASO As New DseUpdatePOBTableAdapters.tbltrans_soTableAdapter
        Dim TBLSO As New DseUpdatePOB.tbltrans_soDataTable
        Dim rowSO As DseUpdatePOB.tbltrans_soRow
        Dim ObjSO As New Sales_Module.ClsTransaksiSalesOrderItem

        Try
            DTASO.Fill(TBLSO)
            Me.ProgSOB.Value = 0
            Me.ProgSOB.Maximum = TBLSO.Rows.Count

            For Each rowSO In TBLSO.Rows
                Me.ProgSOB.Value += 1
                Me.Text = "Update SOB: " + rowSO.No_Transaksi + " (Data " + Me.ProgSOB.Value.ToString + " of " + Me.ProgSOB.Maximum.ToString + ")"
                ObjSO.Set_Order_Balance(rowSO.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
            TBLSO.Dispose()
            ObjSO = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Me.UpdateSOB()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        MsgBox("SOB Update Finish", MsgBoxStyle.Information, "Information")
        Me.Text = "Update Utility"
    End Sub

    Private Sub FrmUpdatePOB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "Update Utility"
        Me.lblPrgStockCard.Text = "Pending"
        Me.LblProgress_UpdateNumberingColletionLetter.Text = "Pending"
        Me.LblProgress_UpdateNumberingMaterialPlanning.Text = "Pending"
        Me.LblUpdateSalesPerson.Text = "Ready"
        Me.dtpRefClosing.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpRefClosing.Value)
        Me.LblClearTable_Progress.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker3.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateDOSudahBuatInv()
        Dim DTA As New DseUpdatePOBTableAdapters.tbltrans_doTableAdapter
        Dim TBL As New DseUpdatePOB.tbltrans_doDataTable
        Dim row As DseUpdatePOB.tbltrans_doRow
        Dim DTA_Inv As New DseUpdatePOBTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim TBL_Inv As New DseUpdatePOB.tbltrans_salesinvoice_itemDataTable
        'Dim row_Inv As DseUpdatePOB.tbltrans_salesinvoice_itemRow

        Try
            DTA.Fill(TBL)
            DTA_Inv.Fill(TBL_Inv)
            Me.ProgUpdateDO1.Value = 0
            Me.ProgUpdateDO1.Maximum = TBL.Rows.Count + 1

            For Each row In TBL.Rows
                Me.ProgUpdateDO1.Value += 1
                Me.Text = "Update DO: " + row.No_Transaksi + " (Data " + Me.ProgUpdateDO1.Value.ToString + " of " + Me.ProgUpdateDO1.Maximum.ToString + ")"
                TBL_Inv.DefaultView.RowFilter = "[Primary DO] = " + row.Primary.ToString
                If TBL_Inv.DefaultView.Count = 0 Then
                    DTA.UpdateQuery_SetBelumBuatInvoice(row.Primary)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_Inv.Dispose()
            TBL_Inv.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Me.UpdateDOSudahBuatInv()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        MsgBox("DO Update Finish", MsgBoxStyle.Information, "Information")
        Me.Text = "Update Utility"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Control.CheckForIllegalCrossThreadCalls = False
        Me.BackgroundWorker4.RunWorkerAsync()
    End Sub

    Private Sub UpdateNoSeriFakturPajak()
        Dim DTACust As New DseUpdatePOBTableAdapters.tblmaster_customerTableAdapter
        Dim TBLCust As New DseUpdatePOB.tblmaster_customerDataTable
        Dim rowCust As DseUpdatePOB.tblmaster_customerRow
        Dim DTAInv As New DseUpdatePOBTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInv As New DseUpdatePOB.tbltrans_salesinvoiceDataTable
        Dim rowInv As DseUpdatePOB.tbltrans_salesinvoiceRow

        Try
            DTACust.Fill(TBLCust)
            DTAInv.Fill(TBLInv)

            'update data customer
            Me.Prog_SeriFakturPajak.Value = 0
            Me.Prog_SeriFakturPajak.Maximum = TBLCust.Rows.Count + TBLInv.Rows.Count

            'update cust
            For Each rowCust In TBLCust.Rows
                Me.Prog_SeriFakturPajak.Value += 1
                Me.Text = "Update customer: " + rowCust.Customer + " (data " + _
                Me.Prog_SeriFakturPajak.Value.ToString + " of " + Me.Prog_SeriFakturPajak.Maximum.ToString + ")"

                If rowCust.Default_Tax_Code.Length > 3 Then
                    DTACust.UpdateQuery(rowCust.Default_Tax_Code.Substring(0, 3), rowCust.Primary)
                End If
            Next

            'update data invoice
            For Each rowInv In TBLInv.Rows
                Me.Prog_SeriFakturPajak.Value += 1
                Me.Text = "Update Invoice: " + rowInv.No_Transaksi + " (data " + Me.Prog_SeriFakturPajak.Value.ToString + _
                " of " + Me.Prog_SeriFakturPajak.Maximum.ToString + ")"
                If rowInv.No_Seri_Faktur_Pajak.Length > 3 Then
                    DTAInv.UpdateQuery(rowInv.No_Seri_Faktur_Pajak.Substring(0, 3), rowInv.Primary)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACust.Dispose()
            TBLCust.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Me.UpdateNoSeriFakturPajak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        MsgBox("Tax Series Update Finish", MsgBoxStyle.Information, "Information")
        Me.Text = "Update Utility"
    End Sub

    Private Sub UpdateNoPO()
        Dim DTA As New DseUpdatePOBTableAdapters.tbltrans_po1TableAdapter
        Dim TBL As New DseUpdatePOB.tbltrans_po1DataTable
        Dim row As DseUpdatePOB.tbltrans_po1Row
        Dim ObjNumber As New Core_Module.ClsMasterTransactionNumbering
        Dim PanjangAcuan As Integer = 0
        Dim ObjPO As New Purchase_Module.ClsTransaksiPO

        Try
            Me.TextBox1.Text = ""
            DTA.Fill(TBL)
            PanjangAcuan = ObjPO.GetNoBaru1(#1/1/2014#).Length

            For Each row In TBL.Rows
                If row.No_Transaksi.Length <> PanjangAcuan Then
                    If row.No_Transaksi.EndsWith("R") = False And _
                    row.No_Transaksi.EndsWith("R2") = False And _
                    row.No_Transaksi.EndsWith("R3") = False And _
                    row.No_Transaksi.EndsWith("R4") = False Then
                        If Me.TextBox1.Text = "" Then
                            Me.TextBox1.Text = row.No_Transaksi
                        Else
                            Me.TextBox1.Text = Me.TextBox1.Text + vbCrLf + row.No_Transaksi
                        End If
                    End If
                End If
            Next
            If Me.TextBox1.Text <> "" Then
                MsgBox("Please fix data", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Prog_UpdateNoPO.Value = 0
            Me.Prog_UpdateNoPO.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.Prog_UpdateNoPO.Value += 1
                Me.TextBox1.Text = "Update PO (data " + Me.Prog_UpdateNoPO.Value.ToString + " of " + Me.Prog_UpdateNoPO.Maximum.ToString + ") - " + row.No_Transaksi
                DTA.UpdateQuery(ObjNumber.GetUrutan(row.No_Transaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor, _
                row.Tanggal_Transaksi), row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.UpdateNoPO()
    End Sub

    Private Sub UpdateNoSO()
        Dim DTA As New DseUpdatepOBTableAdapters.tbltrans_SO1TableAdapter
        Dim TBL As New DseUpdatePOB.tbltrans_so1DataTable
        Dim row As DseUpdatePOB.tbltrans_so1Row
        Dim ObjNumber As New Core_Module.ClsMasterTransactionNumbering
        Dim PanjangAcuan As Integer = 0
        Dim ObjSO As New Sales_Module.ClsTransaksiSalesOrder

        Try
            Me.TextBox1.Text = ""
            DTA.Fill(TBL)
            PanjangAcuan = ObjSO.GetNoBaru1(#1/1/2014#).Length

            For Each row In TBL.Rows
                If row.No_Transaksi.Length <> PanjangAcuan Then
                    If row.No_Transaksi.EndsWith("R") = False And _
                    row.No_Transaksi.EndsWith("R2") = False And _
                    row.No_Transaksi.EndsWith("R3") = False And _
                    row.No_Transaksi.EndsWith("R4") = False Then
                        If Me.TextBox1.Text = "" Then
                            Me.TextBox1.Text = row.No_Transaksi
                        Else
                            Me.TextBox1.Text = Me.TextBox1.Text + vbCrLf + row.No_Transaksi
                        End If
                    End If
                End If
            Next
            If Me.TextBox1.Text <> "" Then
                MsgBox("Please fix data", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Prog_UpdateNoSO.Value = 0
            Me.Prog_UpdateNoSO.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.Prog_UpdateNoSO.Value += 1
                Me.TextBox1.Text = "Update SO (data " + Me.Prog_UpdateNoSO.Value.ToString + " of " + Me.Prog_UpdateNoSO.Maximum.ToString + ") - " + row.No_Transaksi
                DTA.UpdateQuery(ObjNumber.GetUrutan(row.No_Transaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, _
                row.Tanggal_Transaksi), row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.UpdateNoSO()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'Control.CheckForIllegalCrossThreadCalls = False
        'Me.BackgroundWorker5.RunWorkerAsync()
        Me.UpdateClosing()
        MsgBox("Finish")
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Me.UpdateClosing()
    End Sub

    Private Sub UpdateClosing()
        Dim DTA As New Purchase_Module.DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TBL As New Purchase_Module.DseTransaksiMonthlyClosing.tbltrans_purchase_closingDataTable
        Dim row As Purchase_Module.DseTransaksiMonthlyClosing.tbltrans_purchase_closingRow
        Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing
        Dim MinVal As Double = -1

        Try
            DTA.Fill(TBL)
            MinVal = Me.GetMinimumClosing(TBL)
            If TBL.Rows.Count > 1 Then
                Me.prgCost.Value = 0
                Me.prgCost.Maximum = TBL.Rows.Count
                For Each row In TBL.Rows
                    Me.prgCost.Value += 1
                    If row.Primary <> MinVal Then
                        ObjClosing.SimpanDataCost(row.Primary, row.Closing_Period)
                    End If
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetMinimumClosing(ByVal TBL As Purchase_Module.DseTransaksiMonthlyClosing.tbltrans_purchase_closingDataTable) As Double
        Dim row As Purchase_Module.DseTransaksiMonthlyClosing.tbltrans_purchase_closingRow
        Dim MinVal As Double = Double.MaxValue

        Try
            For Each row In TBL.Rows
                If row.Primary < MinVal Then
                    MinVal = row.Primary
                End If
            Next

            Return MinVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Function

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        MsgBox("Finish")
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim ij As New Purchase_Module.FrmTransaksiAOP_Quick2

        Try
            ij.UpdateAtAmount = True
            ij.UpdatePrg = Me.ProgressBar1
            ij.ShowDialog()
            MsgBox("Update Finish")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnItemPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnItemPR.Click

        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker6.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Dim DTA_PR As New Purchase_Module.DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TBL_PR As New Purchase_Module.DseTransaksiPurchaseReturn.tbltrans_purchasereturn1DataTable
        Dim row_PR As Purchase_Module.DseTransaksiPurchaseReturn.tbltrans_purchasereturn1Row
        Dim DTA_Item As New Purchase_Module.DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2TableAdapter
        Dim TBL_Item As New Purchase_Module.DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2DataTable
        Dim row_Item As Purchase_Module.DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2Row
        Dim DTA_Update As New Purchase_Module.DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter

        Dim IDRAmount As Decimal = 0
        Dim DefPrecIDR As Integer = 0
        Dim DefPrecCur As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim SubTotalIDR As Decimal = 0
        Dim SubTotal, SubTotal1, SubTotalIDRAmount As Decimal

        Try
            DefPrecIDR = ObjCur.GetDecimal_DefaultCurrency
            DTA_PR.Fill(TBL_PR)
            Me.prgItemPR.Value = 0
            Me.prgItemPR.Maximum = TBL_PR.Rows.Count

            For Each row_PR In TBL_PR.Rows
                Me.prgItemPR.Value += 1
                DTA_Item.FillBy_PrimaryTransaksi(TBL_Item, row_PR.Primary)
                DefPrecCur = ObjCur.GetDefaultPrecission(row_PR.Primary_Currency)
                SubTotal = 0
                SubTotal1 = 0
                SubTotalIDRAmount = 0

                For i As Integer = 0 To TBL_Item.Rows.Count - 1
                    row_Item = TBL_Item.Rows(i)
                    With row_Item
                        SubTotal1 += Math.Round(.Nett_Price * .Total_Qty * .Exchanged_Rate, DefPrecCur, MidpointRounding.AwayFromZero)
                    End With
                Next
                SubTotal = Math.Round(SubTotal1 * row_PR.Tax_Exchange_Rate, DefPrecIDR, MidpointRounding.AwayFromZero)
                For i As Integer = 0 To TBL_Item.Rows.Count - 1
                    row_Item = TBL_Item.Rows(i)
                    With row_Item
                        If i <> TBL_Item.Rows.Count - 1 Then
                            IDRAmount = Math.Round((.Nett_Price * .Total_Qty * .Exchanged_Rate * SubTotal) / SubTotal1, DefPrecIDR, MidpointRounding.AwayFromZero)
                            SubTotalIDRAmount += IDRAmount
                        Else
                            IDRAmount = SubTotal - SubTotalIDRAmount
                        End If
                        DTA_Update.UpdateQuery(IDRAmount, .Primary)
                    End With
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_PR.Dispose()
            TBL_PR.Dispose()
            DTA_Item.Dispose()
            TBL_Item.Dispose()
            DTA_Update.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            MsgBox("PR update completed", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_StockTaking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_StockTaking.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_StockTaking.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_StockTaking.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_StockTaking.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_StockTaking.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_StockTaking.Text))
                Me.TxtAccountDesc_StockTaking.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_StockTaking.Text)
                Me.TxtAccount_Adjustment.Focus()
            Else
                Me.TxtAccount_StockTaking.Clear()
                Me.TxtPrimaryAccount_StockTaking.Text = "-1"
                Me.TxtAccountDesc_StockTaking.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_StockTaking.Text)
                Me.TxtAccount_StockTaking.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_StockTaking_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_StockTaking.GotFocus
        Try
            If Me.TxtAccount_StockTaking.Enabled = True Then
                If Me.TxtAccount_StockTaking.ReadOnly = True Then
                    Me.TxtAccount_StockTaking.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_StockTaking.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_StockTaking_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_StockTaking.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_StockTaking.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_StockTaking.Text) = True Then
                        Me.TxtPrimaryAccount_StockTaking.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_StockTaking.Text))
                        Me.TxtAccount_StockTaking.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_StockTaking.Text))
                        Me.TxtAccount_Adjustment.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_StockTaking.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_StockTaking.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_StockTaking.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_StockTaking.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_StockTaking_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_StockTaking.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_StockTaking.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_StockTaking.Text) = True Then
                    Me.TxtPrimaryAccount_StockTaking.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_StockTaking.Text))
                    Me.TxtAccount_StockTaking.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_StockTaking.Text))
                Else
                    If Me.TxtAccount_StockTaking.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_StockTaking.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_StockTaking.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_StockTaking.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_StockTaking.Text))
                    Else
                        Me.TxtAccount_StockTaking.Clear()
                        Me.TxtPrimaryAccount_StockTaking.Text = "-1"
                        Me.TxtAccount_StockTaking.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_StockTaking.Clear()
                Me.TxtPrimaryAccount_StockTaking.Text = "-1"
            End If
            Me.TxtAccountDesc_StockTaking.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_StockTaking.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_StockTaking_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_StockTaking.LostFocus
        Try
            If Me.TxtAccount_StockTaking.Enabled = True Then
                If Me.TxtAccount_StockTaking.ReadOnly = True Then
                    Me.TxtAccount_StockTaking.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_StockTaking.BackColor = Me.TxtPrimaryAccount_StockTaking.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_StockTaking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_StockTaking.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_StockTaking.Text <> "" And Me.TxtPrimaryAccount_StockTaking.Text <> "-1" Then
                If MsgBox("Update stock taking account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_StockTakingAccount(Me.TxtPrimaryAccount_StockTaking.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_Adjustment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_Adjustment.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Adjustment.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_Adjustment.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_Adjustment.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_Adjustment.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Adjustment.Text))
                Me.TxtAccountDesc_Adjustment.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Adjustment.Text)
                'Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtAccount_Adjustment.Clear()
                Me.TxtPrimaryAccount_Adjustment.Text = "-1"
                Me.TxtAccountDesc_Adjustment.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Adjustment.Text)
                Me.TxtAccount_Adjustment.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Adjustment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Adjustment.GotFocus
        Try
            If Me.TxtAccount_Adjustment.Enabled = True Then
                If Me.TxtAccount_Adjustment.ReadOnly = True Then
                    Me.TxtAccount_Adjustment.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Adjustment.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Adjustment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_Adjustment.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_Adjustment.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Adjustment.Text) = True Then
                        Me.TxtPrimaryAccount_Adjustment.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Adjustment.Text))
                        Me.TxtAccount_Adjustment.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Adjustment.Text))
                        'Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_Adjustment.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_Adjustment.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_Adjustment.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Adjustment.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Adjustment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Adjustment.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Adjustment.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Adjustment.Text) = True Then
                    Me.TxtPrimaryAccount_Adjustment.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Adjustment.Text))
                    Me.TxtAccount_Adjustment.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Adjustment.Text))
                Else
                    If Me.TxtAccount_Adjustment.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_Adjustment.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_Adjustment.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_Adjustment.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Adjustment.Text))
                    Else
                        Me.TxtAccount_Adjustment.Clear()
                        Me.TxtPrimaryAccount_Adjustment.Text = "-1"
                        Me.TxtAccount_Adjustment.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_Adjustment.Clear()
                Me.TxtPrimaryAccount_Adjustment.Text = "-1"
            End If
            Me.TxtAccountDesc_Adjustment.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Adjustment.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Adjustment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Adjustment.LostFocus
        Try
            If Me.TxtAccount_Adjustment.Enabled = True Then
                If Me.TxtAccount_Adjustment.ReadOnly = True Then
                    Me.TxtAccount_Adjustment.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Adjustment.BackColor = Me.TxtPrimaryAccount_Adjustment.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_Adjustment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_Adjustment.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_Adjustment.Text <> "" And Me.TxtPrimaryAccount_Adjustment.Text <> "-1" Then
                If MsgBox("Update adjustment account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_AdjustmentAccount(Me.TxtPrimaryAccount_Adjustment.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_Inventory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_Inventory.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Inventory.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_Inventory.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_Inventory.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_Inventory.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Inventory.Text))
                Me.TxtAccountDesc_Inventory.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Inventory.Text)
                'Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtAccount_Inventory.Clear()
                Me.TxtPrimaryAccount_Inventory.Text = "-1"
                Me.TxtAccountDesc_Inventory.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Inventory.Text)
                Me.TxtAccount_Inventory.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Inventory_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Inventory.GotFocus
        Try
            If Me.TxtAccount_Inventory.Enabled = True Then
                If Me.TxtAccount_Inventory.ReadOnly = True Then
                    Me.TxtAccount_Inventory.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Inventory.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Inventory_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_Inventory.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_Inventory.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Inventory.Text) = True Then
                        Me.TxtPrimaryAccount_Inventory.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Inventory.Text))
                        Me.TxtAccount_Inventory.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Inventory.Text))
                        'Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_Inventory.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_Inventory.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_Inventory.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Inventory.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Inventory_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Inventory.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Inventory.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Inventory.Text) = True Then
                    Me.TxtPrimaryAccount_Inventory.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Inventory.Text))
                    Me.TxtAccount_Inventory.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Inventory.Text))
                Else
                    If Me.TxtAccount_Inventory.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_Inventory.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_Inventory.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_Inventory.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Inventory.Text))
                    Else
                        Me.TxtAccount_Inventory.Clear()
                        Me.TxtPrimaryAccount_Inventory.Text = "-1"
                        Me.TxtAccount_Inventory.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_Inventory.Clear()
                Me.TxtPrimaryAccount_Inventory.Text = "-1"
            End If
            Me.TxtAccountDesc_Inventory.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Inventory.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Inventory_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Inventory.LostFocus
        Try
            If Me.TxtAccount_Inventory.Enabled = True Then
                If Me.TxtAccount_Inventory.ReadOnly = True Then
                    Me.TxtAccount_Inventory.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Inventory.BackColor = Me.TxtPrimaryAccount_Inventory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_Inventory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_Inventory.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_Inventory.Text <> "" And Me.TxtPrimaryAccount_Inventory.Text <> "-1" Then
                If MsgBox("Update Inventory account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_Inventory(Me.TxtPrimaryAccount_Inventory.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_Sales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_Sales.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Sales.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_Sales.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_Sales.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_Sales.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Sales.Text))
                Me.TxtAccountDesc_Sales.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Sales.Text)
                'Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtAccount_Sales.Clear()
                Me.TxtPrimaryAccount_Sales.Text = "-1"
                Me.TxtAccountDesc_Sales.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Sales.Text)
                Me.TxtAccount_Sales.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Sales_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Sales.GotFocus
        Try
            If Me.TxtAccount_Sales.Enabled = True Then
                If Me.TxtAccount_Sales.ReadOnly = True Then
                    Me.TxtAccount_Sales.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Sales.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Sales_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_Sales.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_Sales.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Sales.Text) = True Then
                        Me.TxtPrimaryAccount_Sales.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Sales.Text))
                        Me.TxtAccount_Sales.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Sales.Text))
                        'Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_Sales.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_Sales.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_Sales.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Sales.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_Sales_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Sales.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_Sales.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_Sales.Text) = True Then
                    Me.TxtPrimaryAccount_Sales.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_Sales.Text))
                    Me.TxtAccount_Sales.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Sales.Text))
                Else
                    If Me.TxtAccount_Sales.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_Sales.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_Sales.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_Sales.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_Sales.Text))
                    Else
                        Me.TxtAccount_Sales.Clear()
                        Me.TxtPrimaryAccount_Sales.Text = "-1"
                        Me.TxtAccount_Sales.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_Sales.Clear()
                Me.TxtPrimaryAccount_Sales.Text = "-1"
            End If
            Me.TxtAccountDesc_Sales.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_Sales.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_Sales_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_Sales.LostFocus
        Try
            If Me.TxtAccount_Sales.Enabled = True Then
                If Me.TxtAccount_Sales.ReadOnly = True Then
                    Me.TxtAccount_Sales.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_Sales.BackColor = Me.TxtPrimaryAccount_Sales.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_Sales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_Sales.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_Sales.Text <> "" And Me.TxtPrimaryAccount_Sales.Text <> "-1" Then
                If MsgBox("Update Sales account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_Sales(Me.TxtPrimaryAccount_Sales.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_SalesReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_SalesReturn.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_SalesReturn.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_SalesReturn.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_SalesReturn.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_SalesReturn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_SalesReturn.Text))
                Me.TxtAccountDesc_SalesReturn.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_SalesReturn.Text)
                'Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtAccount_SalesReturn.Clear()
                Me.TxtPrimaryAccount_SalesReturn.Text = "-1"
                Me.TxtAccountDesc_SalesReturn.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_SalesReturn.Text)
                Me.TxtAccount_SalesReturn.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_SalesReturn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_SalesReturn.GotFocus
        Try
            If Me.TxtAccount_SalesReturn.Enabled = True Then
                If Me.TxtAccount_SalesReturn.ReadOnly = True Then
                    Me.TxtAccount_SalesReturn.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_SalesReturn.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_SalesReturn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_SalesReturn.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_SalesReturn.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_SalesReturn.Text) = True Then
                        Me.TxtPrimaryAccount_SalesReturn.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_SalesReturn.Text))
                        Me.TxtAccount_SalesReturn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_SalesReturn.Text))
                        'Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_SalesReturn.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_SalesReturn.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_SalesReturn.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_SalesReturn.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_SalesReturn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_SalesReturn.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_SalesReturn.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_SalesReturn.Text) = True Then
                    Me.TxtPrimaryAccount_SalesReturn.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_SalesReturn.Text))
                    Me.TxtAccount_SalesReturn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_SalesReturn.Text))
                Else
                    If Me.TxtAccount_SalesReturn.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_SalesReturn.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_SalesReturn.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_SalesReturn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_SalesReturn.Text))
                    Else
                        Me.TxtAccount_SalesReturn.Clear()
                        Me.TxtPrimaryAccount_SalesReturn.Text = "-1"
                        Me.TxtAccount_SalesReturn.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_SalesReturn.Clear()
                Me.TxtPrimaryAccount_SalesReturn.Text = "-1"
            End If
            Me.TxtAccountDesc_SalesReturn.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_SalesReturn.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_SalesReturn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_SalesReturn.LostFocus
        Try
            If Me.TxtAccount_SalesReturn.Enabled = True Then
                If Me.TxtAccount_SalesReturn.ReadOnly = True Then
                    Me.TxtAccount_SalesReturn.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_SalesReturn.BackColor = Me.TxtPrimaryAccount_SalesReturn.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_SalesReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_SalesReturn.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_SalesReturn.Text <> "" And Me.TxtPrimaryAccount_SalesReturn.Text <> "-1" Then
                If MsgBox("Update SalesReturn account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_SalesReturn(Me.TxtPrimaryAccount_SalesReturn.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindAccount_COGS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAccount_COGS.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_COGS.Text <> "" Then
                ij.txtCari.Text = Me.TxtAccount_COGS.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccount_COGS.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAccount_COGS.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_COGS.Text))
                Me.TxtAccountDesc_COGS.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_COGS.Text)
                'Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtAccount_COGS.Clear()
                Me.TxtPrimaryAccount_COGS.Text = "-1"
                Me.TxtAccountDesc_COGS.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_COGS.Text)
                Me.TxtAccount_COGS.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_COGS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_COGS.GotFocus
        Try
            If Me.TxtAccount_COGS.Enabled = True Then
                If Me.TxtAccount_COGS.ReadOnly = True Then
                    Me.TxtAccount_COGS.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_COGS.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_COGS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccount_COGS.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccount_COGS.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_COGS.Text) = True Then
                        Me.TxtPrimaryAccount_COGS.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_COGS.Text))
                        Me.TxtAccount_COGS.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_COGS.Text))
                        'Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAccount_COGS.PerformClick()
                    End If
                Else
                    Me.btnFindAccount_COGS.PerformClick()
                End If
            End If
            Me.TxtAccountDesc_COGS.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_COGS.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccount_COGS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_COGS.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAccount_COGS.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAccount_COGS.Text) = True Then
                    Me.TxtPrimaryAccount_COGS.Text = CStr(ObjAkun.GetPrimary(Me.TxtAccount_COGS.Text))
                    Me.TxtAccount_COGS.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_COGS.Text))
                Else
                    If Me.TxtAccount_COGS.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAccount_COGS.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccount_COGS.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAccount_COGS.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAccount_COGS.Text))
                    Else
                        Me.TxtAccount_COGS.Clear()
                        Me.TxtPrimaryAccount_COGS.Text = "-1"
                        Me.TxtAccount_COGS.Focus()
                    End If
                End If
            Else
                Me.TxtAccount_COGS.Clear()
                Me.TxtPrimaryAccount_COGS.Text = "-1"
            End If
            Me.TxtAccountDesc_COGS.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount_COGS.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAccount_COGS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccount_COGS.LostFocus
        Try
            If Me.TxtAccount_COGS.Enabled = True Then
                If Me.TxtAccount_COGS.ReadOnly = True Then
                    Me.TxtAccount_COGS.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtAccount_COGS.BackColor = Me.TxtPrimaryAccount_COGS.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateAccount_COGS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateAccount_COGS.Click
        Try
            'update stock taking account
            If Me.TxtPrimaryAccount_COGS.Text <> "" And Me.TxtPrimaryAccount_COGS.Text <> "-1" Then
                If MsgBox("Update COGS account?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.ObjProd.UpdateAccount_COGS(Me.TxtPrimaryAccount_COGS.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdatev160_SC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatev160_SC.Click
        Try
            If MsgBox("Update stock card?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker7.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker7_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker7.DoWork
        Me.UpdateStockCard()
    End Sub

    Private Sub BackgroundWorker7_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted
        MsgBox("Update finish")
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatev160_UpdateNumberingCL.Click
        Try
            If MsgBox("Update collection letter numbering?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.UpdateNumberingCollectionLetter(Me.LblProgress_UpdateNumberingColletionLetter)
                MsgBox("Update finish")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdatev160_UpdateNumberingMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatev160_UpdateNumberingMP.Click
        Try
            If MsgBox("Update material planning numbering?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.UpdateNumberingMaterialPlanning(Me.LblProgress_UpdateNumberingMaterialPlanning)
                MsgBox("Update finish")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteMPSF.Click
        'Dim DTAPO As New DseUpdate_OtherTableAdapters.tbltrans_poTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA1 As New DseUpdate_OtherTableAdapters.tbltrans_material_planningTableAdapter
        Dim TBL1 As New DseUpdate_Other.tbltrans_material_planningDataTable
        Dim row1 As DseUpdate_Other.tbltrans_material_planningRow
        Dim StartIndex As Integer = 0
        Dim DTA2 As New DseUpdate_OtherTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim TBL2 As New DseUpdate_Other.tbltrans_shipping_forecastDataTable
        Dim row2 As DseUpdate_Other.tbltrans_shipping_forecastRow


        Try
            Me.prgRefMP.Value = 0

            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring

            If MsgBox("Delete MP & SF data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Button14.Enabled = False
                If Me.dtpRefMP.Checked = True Then
                    DTA1.FillBy(TBL1, ObjCD.GetLasDateOfTheMonth(Me.dtpRefMP.Value))
                Else
                    DTA1.Fill(TBL1)
                End If
                If TBL1.Rows.Count > 0 Then
                    Me.prgRefMP.Maximum = TBL1.Rows.Count
                    StartIndex = 0

                    While StartIndex <= TBL1.Rows.Count - 1
                        row1 = TBL1.Rows(StartIndex)

                        If StartIndex = TBL1.Rows.Count - 1 Then
                            DTA1.DeleteQuery(row1.primary)
                        ElseIf Math.IEEERemainder(StartIndex, 5) = 0 Then
                            DTA1.DeleteQuery(row1.primary)
                        End If

                        StartIndex += 1
                        Me.prgRefMP.Value += 1
                    End While
                End If

                Me.prgRefMP.Value = 0
                If Me.dtpRefMP.Checked = True Then
                    DTA2.FillBy(TBL2, ObjCD.GetLasDateOfTheMonth(Me.dtpRefMP.Value))
                Else
                    DTA2.Fill(TBL2)
                End If
                If TBL2.Rows.Count > 0 Then
                    Me.prgRefMP.Maximum = TBL2.Rows.Count
                    StartIndex = 0

                    While StartIndex <= TBL2.Rows.Count - 1
                        row2 = TBL2.Rows(StartIndex)

                        If StartIndex = TBL2.Rows.Count - 1 Then
                            DTA2.DeleteQuery(row2.primary)
                        ElseIf Math.IEEERemainder(StartIndex, 5) = 0 Then
                            DTA2.DeleteQuery(row2.primary)
                        End If

                        StartIndex += 1
                        Me.prgRefMP.Value += 1
                    End While
                End If

                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                Me.Button14.Enabled = True
            End If
            'If Me.dtpRefMP.Checked = True Then
            '    If MsgBox("Delete MP & SF data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

            '        DTAPO.DeleteQuery_ByRefDate(ObjCD.GetLasDateOfTheMonth(Me.dtpRefMP.Value))
            '        DTAPO.DeleteQuery_MP_Grouping()
            '        DTAPO.DeleteQuerySF_ByRefDate(ObjCD.GetLasDateOfTheMonth(Me.dtpRefMP.Value))
            '        MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
            '    End If
            'Else
            '    If MsgBox("Delete MP & SF data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

            '        DTAPO.DeleteQuery_MP()
            '        DTAPO.DeleteQuery_MP_Grouping()
            '        DTAPO.DeleteQuery_SF()
            '        MsgBox("Data deleted.", MsgBoxStyle.Information, "Information")
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdatePenghasilanBruto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePenghasilanBruto.Click
        Dim DTA As New DseUpdate_OtherTableAdapters.tbltrans_salaryTableAdapter
        Dim TBL As New DseUpdate_Other.tbltrans_salaryDataTable
        Dim row As DseUpdate_Other.tbltrans_salaryRow
        Dim DTA1 As New DseUpdate_OtherTableAdapters.tbltrans_salary_itemTableAdapter
        Dim TBL1 As New DseUpdate_Other.tbltrans_salary_itemDataTable
        Dim row1 As DseUpdate_Other.tbltrans_salary_itemRow
        Dim DTA2 As New DseUpdate_OtherTableAdapters.tbltrans_salary_item_otherTableAdapter
        Dim TBL2 As New DseUpdate_Other.tbltrans_salary_item_otherDataTable
        Dim row2 As DseUpdate_Other.tbltrans_salary_item_otherRow
        Dim RetVal As Decimal = 0
        Dim Tipe1 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.gaji_pensiun
        Dim Tipe2 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.honorarium
        Dim Tipe3 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.penerimaan_dalam_bentuk_natura
        Dim Tipe4 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.premi_asuransi
        Dim Tipe5 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.Tanteiem_bonus_gratifikasi_jasa_produksi_thr
        Dim Tipe6 As Core_Module.ClsMasterSalaryFormula_item.EnumTax2 = Core_Module.ClsMasterSalaryFormula_item.EnumTax2.tunjangan_lainnya

        Dim TipeIncome1 As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Income
        Dim TipeIncome2 As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.CompanyExpense
        Dim TipeIncome3 As Core_Module.FrmMasterSalaryFormula_Component.EnumType = Core_Module.FrmMasterSalaryFormula_Component.EnumType.Deduction

        Try
            DTA.Fill(TBL)
            Me.prgUpdatePenghasilanBruto.Value = 0
            Me.prgUpdatePenghasilanBruto.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                DTA1.Fill(TBL1, row.Primary)
                DTA2.Fill(TBL2, row.Primary)

                RetVal = 0
                For Each row1 In TBL1.Rows
                    If row1.Amount <> 0 Then
                        If row1.Tax_Component = Tipe1 Or row1.Tax_Component = Tipe2 Or row1.Tax_Component = Tipe3 Or _
                        row1.Tax_Component = Tipe4 Or row1.Tax_Component = Tipe5 Or row1.Tax_Component = Tipe6 Then
                            If row1.Type = TipeIncome1 Or row1.Type = TipeIncome2 Then
                                RetVal += row1.Amount
                            ElseIf row1.Type = TipeIncome3 Then
                                RetVal -= row1.Amount
                            End If
                        End If

                    End If
                Next
                For Each row2 In TBL2.Rows
                    If row2.Amount <> 0 Then
                        If row2.Tax_Component = Tipe1 Or row2.Tax_Component = Tipe2 Or row2.Tax_Component = Tipe3 Or _
                        row2.Tax_Component = Tipe4 Or row2.Tax_Component = Tipe5 Or row2.Tax_Component = Tipe6 Then
                            If row2.Type = TipeIncome1 Or row2.Type = TipeIncome2 Then
                                RetVal += row2.Amount
                            ElseIf row2.Type = TipeIncome3 Then
                                RetVal -= row2.Amount
                            End If
                        End If

                    End If
                Next

                DTA.UpdateQuery(RetVal + row._02TunjanganPPh, row.Primary)
                Me.prgUpdatePenghasilanBruto.Value += 1
            Next

            MsgBox("Updated")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCekSalaryKomponen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCekSalaryKomponen.Click
        Dim DTAAkun As New DseUpdate_v180TableAdapters.tblmaster_akunTableAdapter
        Dim TBLAkun As New DseUpdate_v180.tblmaster_akunDataTable
        Dim DTASalaryComponen As New DseUpdate_v180TableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim TBLSalaryComponen As New DseUpdate_v180.tblmaster_salary_formula_salarycomponent1DataTable
        Dim rowSalaryComponen As DseUpdate_v180.tblmaster_salary_formula_salarycomponent1Row
        Dim ComponentTidakAdaAkun As New List(Of String)
        Dim RetVal As String = ""

        Try
            DTAAkun.Fill(TBLAkun)
            DTASalaryComponen.Fill(TBLSalaryComponen)

            For Each rowSalaryComponen In TBLSalaryComponen.Rows
                If rowSalaryComponen.Primary_Expense_Account <> -1 Then
                    TBLAkun.DefaultView.RowFilter = "Primary = " + rowSalaryComponen.Primary_Expense_Account.ToString
                    If TBLAkun.DefaultView.Count = 0 Then
                        If ComponentTidakAdaAkun.Contains(rowSalaryComponen.Description) = False Then
                            ComponentTidakAdaAkun.Add(rowSalaryComponen.Description)
                        End If
                    End If
                End If
                If rowSalaryComponen.primary_payment_account <> -1 Then
                    TBLAkun.DefaultView.RowFilter = "Primary = " + rowSalaryComponen.primary_payment_account.ToString
                    If TBLAkun.DefaultView.Count = 0 Then
                        If ComponentTidakAdaAkun.Contains(rowSalaryComponen.Description) = False Then
                            ComponentTidakAdaAkun.Add(rowSalaryComponen.Description)
                        End If
                    End If
                End If
            Next

            If ComponentTidakAdaAkun.Count > 0 Then
                For Each x As String In ComponentTidakAdaAkun
                    If RetVal = "" Then
                        RetVal = x
                    Else
                        RetVal = RetVal + ", " + x
                    End If
                Next
                MsgBox("Tidak ada akun: " + RetVal, MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Data Check")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePrimaryAkun.Click
        Dim DTASalaryComponen As New DseUpdate_v180TableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim TBLSalaryComponen As New DseUpdate_v180.tblmaster_salary_formula_salarycomponent1DataTable
        Dim rowSalaryComponen As DseUpdate_v180.tblmaster_salary_formula_salarycomponent1Row

        Try
            DTASalaryComponen.Fill(TBLSalaryComponen)

            For Each rowSalaryComponen In TBLSalaryComponen.Rows
                DTASalaryComponen.UpdateQuery(rowSalaryComponen.Primary_Expense_Account, _
                rowSalaryComponen.primary_payment_account, rowSalaryComponen.Identifier)
            Next

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New Core_Module.frmCariMasterProduct
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll


        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProd.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
                    Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDesc.Text = ""
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New Core_Module.frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProd.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProd.GetPrimary(Me.TxtCodeNo.Text)
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
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDesc.Text = ""
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtDesc.Clear()
            End If
            Me.IsikanSC_n_Qty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
           i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProd.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.TxtDesc.Text = Me.ObjProd.GetDescription(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtDesc.Text = ""
                Me.TxtCodeNo.Focus()
            End If
            Me.IsikanSC_n_Qty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub IsikanSC_n_Qty()
        Dim rowDGV As DataGridViewRow
        Dim DTASC As New DseUpdate_OtherTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New DseUpdate_Other.tbltransaksi_stock_cardDataTable
        'Dim rowSC As DseUpdate_Other.tbltransaksi_stock_cardRow
        Dim DTAQty As New DseUpdate_OtherTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLQty As New DseUpdate_Other.tblmaster_productqtyDataTable
        'Dim rowQty As DseUpdate_Other.tblmaster_productqtyRow

        Try
            'sc
            Me.dgvSC.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DTASC.Fill(TBLSC, Me.TxtPrimaryCodeNo.Text)
            Me.dgvSC.DataSource = TBLSC.DefaultView
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                If Me.dgvSC.RowCount > 0 Then
                    rowDGV = Me.dgvSC.Rows(0)
                    If rowDGV.Cells(0).Value = "" Then
                        Me.dgvSC.Columns(0).Visible = False
                    Else
                        Me.dgvSC.Columns(0).Visible = True
                    End If
                    If rowDGV.Cells(1).Value = "" Then
                        Me.dgvSC.Columns(1).Visible = False
                    Else
                        Me.dgvSC.Columns(1).Visible = True
                    End If
                    Me.dgvSC.Columns(2).DefaultCellStyle.Format = "N" + Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text).ToString
                    Me.dgvSC.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
                Me.dgvSC.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If

            'qty
            Me.dgvQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DTAQty.Fill(TBLQty, Me.TxtPrimaryCodeNo.Text)
            Me.dgvQty.DataSource = TBLQty.DefaultView
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                If Me.dgvQty.RowCount > 0 Then
                    rowDGV = Me.dgvQty.Rows(0)
                    If rowDGV.Cells(0).Value = "" Then
                        Me.dgvQty.Columns(0).Visible = False
                    Else
                        Me.dgvQty.Columns(0).Visible = True
                    End If
                    If rowDGV.Cells(1).Value = "" Then
                        Me.dgvQty.Columns(1).Visible = False
                    Else
                        Me.dgvQty.Columns(1).Visible = True
                    End If
                    Me.dgvQty.Columns(2).DefaultCellStyle.Format = "N" + Me.ObjProd.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text).ToString
                    Me.dgvQty.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
                Me.dgvQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASC.Dispose()
            TBLSC.Dispose()
            DTAQty.Dispose()
            TBLQty.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
               Me.IsikanSC_n_Qty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'menghapuskan stock card yang tipe transaksi-nya out sales dan tanggal transaksi-nya dibawah covertion date
    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim TipeSC As Integer = 0
        Dim TipeSC1 As Integer = 1
        Dim RefDate As Date = #1/30/2014#
        Dim DTA As New DseUpdate_v180TableAdapters.tbltransaksi_stock_cardTableAdapter

        Try
            DTA.DeleteQuery(RefDate, TipeSC)
            DTA.DeleteQuery(RefDate, TipeSC1)

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnCekUnbalanceAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCekUnbalanceAccount.Click
        Dim DTA_Unbalance_Journal As New AccModule.DseTransaksiJournalTableAdapters.UnBalanceJurnalTableAdapter
        Dim TBL_Unbalance_Journal As New AccModule.DseTransaksiJournal.UnBalanceJurnalDataTable
        Dim row_UnBalance_Journal As AccModule.DseTransaksiJournal.UnBalanceJurnalRow
        Dim counter As Integer = 0
        Dim Dummy As String = ""
        Dim Defprec1 As Integer = 0

        Try
            DTA_Unbalance_Journal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Unbalance_Journal.Fill(TBL_Unbalance_Journal)
            counter = 1
            Defprec1 = Me.ObjCur.GetDecimal_DefaultCurrency

            If TBL_Unbalance_Journal.Rows.Count > 0 Then
                For Each row_UnBalance_Journal In TBL_Unbalance_Journal.Rows
                    If Me.ObjCD.DigitGrouping(row_UnBalance_Journal.Debet, DefPrec1) <> Me.ObjCD.DigitGrouping(row_UnBalance_Journal.Kredit, DefPrec1) Then
                        If Dummy = "" Then
                            Dummy = counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                        Else
                            Dummy = Dummy + vbCrLf + counter.ToString + ". " + row_UnBalance_Journal.Ref_Number
                        End If
                        counter += 1
                    End If
                Next
                If Counter > 1 Then
                    MsgBox("The following journal is unbalance, please contact your vendor to solve this issue." + vbCrLf + Dummy, _
                     MsgBoxStyle.Information, "Confirm")
                Else
                    MsgBox("Data checked", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Data checked", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Unbalance_Journal.Dispose()
            TBL_Unbalance_Journal.Dispose()
        End Try
    End Sub

   
    Private Sub btnDeleteClosing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteClosing.Click
        Dim RefDate As Date = #1/1/1950#
        Dim DTA As New DseUpdate_OtherTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TBL As New DseUpdate_Other.tbltrans_purchase_closingDataTable
        Dim row As DseUpdate_Other.tbltrans_purchase_closingRow
        Dim SampeUjung As Boolean = False
        Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing
        Dim CurIdx As Integer = -1

        Try
            If MsgBox("Delete purchase closing data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                If Me.dtpRefClosing.Checked = True Then
                    RefDate = Me.dtpRefClosing.Value
                End If
                DTA.Fill(TBL, RefDate)
                Me.prgRefClosing.Value = 0
                Me.prgRefClosing.Maximum = TBL.Rows.Count

                For Each row In TBL.Rows
                    CurIdx += 1
                    Me.prgRefClosing.Value += 1

                    If ObjClosing.ApakahBolehHapus(row.Primary, False, "USER", False) = True Then
                        ObjClosing.Hapus(row.Primary, "USER", False)
                        If CurIdx = TBL.Rows.Count - 1 Then
                            SampeUjung = True
                        End If
                    Else
                        SampeUjung = False
                        Exit For
                    End If
                Next

                If SampeUjung = True Then
                    MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Data deletion not finish", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjClosing = Nothing
        End Try
    End Sub

    Private Sub btnBayarAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBayarAR.Click
        Dim DTA As New DseUpdate_OtherTableAdapters.tblmaster_customer1TableAdapter
        Dim TBL As New DseUpdate_Other.tblmaster_customer1DataTable
        Dim row As DseUpdate_Other.tblmaster_customer1Row

        Try
            If MsgBox("Update AR?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'DTA.Fill(TBL)
                'Me.prgBayarAR.Value = 0
                'Me.prgBayarAR.Maximum = TBL.Rows.Count

                'For Each row In TBL.Rows
                '    Me.prgBayarAR.Value += 1

                '    Me.LblProgressARPay.Text = "Data " + Me.prgBayarAR.Value.ToString + " of " + Me.prgBayarAR.Maximum.ToString + " - " + row.Customer_Code
                '    Dim i As New AccModule.FrmTransaksiJournal_Sales
                '    i.UserCode = "USER"
                '    i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
                '    ' Set the caption to be unique.
                '    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                '    i.vPrimaryCustomerBayar = row.Primary
                '    i.vTglBayar = Me.dtpARPay.Value.Date
                '    i.ShowDialog()
                '    i.Dispose()
                'Next

                'Me.prgBayarAR.Value = 0
                'Me.LblProgressARPay.Text = "Updated"
                'MsgBox("Data updated")
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwARPay.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwARPay_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwARPay.DoWork
        Dim DTA As New DseUpdate_OtherTableAdapters.tblmaster_customer1TableAdapter
        Dim TBL As New DseUpdate_Other.tblmaster_customer1DataTable
        Dim row As DseUpdate_Other.tblmaster_customer1Row
        Dim Lanjut As Boolean = True

        Try
            Me.ObjCD.UpdateThreading()

            ' Increment the caption counter.
            DTA.Fill(TBL)
            Me.prgBayarAR.Value = 0
            Me.prgBayarAR.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.prgBayarAR.Value += 1

                Me.LblProgressARPay.Text = "Data " + Me.prgBayarAR.Value.ToString + " of " + Me.prgBayarAR.Maximum.ToString + " - " + row.Customer_Code

                Lanjut = True
                While Lanjut = True
                    Dim i As New AccModule.FrmTransaksiJournal_Sales
                    i.UserCode = "USER"
                    i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
                    ' Set the caption to be unique.
                    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                    i.vPrimaryCustomerBayar = row.Primary
                    i.vTglBayar = Me.dtpARPay.Value.Date
                    i.ShowDialog()
                    Lanjut = i.vLanjut

                    i.Dispose()
                End While

            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwARPay_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwARPay.RunWorkerCompleted
        Try
            Me.prgBayarAR.Value = 0
            Me.LblProgressARPay.Text = "Updated"
            MsgBox("Data updated")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBalikJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBalikJurnal.Click
        Try
            If MsgBox("Balik Jurnal?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwBalikJurnal.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwBalikJurnal_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwBalikJurnal.DoWork
        Dim DTAJurnal As New DseUpdate_OtherTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBLJurnal As New DseUpdate_Other.tbltrans_jurnalDataTable
        Dim rowJurnal As DseUpdate_Other.tbltrans_jurnalRow
        Dim DTAItem As New DseUpdate_OtherTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TBLItem As New DseUpdate_Other.tbltrans_jurnal_itemDataTable
        Dim rowItem As DseUpdate_Other.tbltrans_jurnal_itemRow
        Dim ModificationString As String = "(SMESS REVERSAL - "
        Dim DV As DataRowView
        Dim counter As Integer = 0
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim DummyString As String = ""
        Dim DummyVal As Decimal = 0
        Dim PrimaryJurnal As Double = -1

        Try
            Me.ObjCD.UpdateThreading()
            DTAJurnal.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTAJurnal.Connection.ConnectionString

            '1. hapuskan data reversal yang sudah dibuat sebelumnya
            TBLJurnal.Rows.Clear()
            Me.lblProgressBalikJurnal.Text = "Phase 1 of 2: Deleting reversal journal"
            DTAJurnal.FillBy_Date(TBLJurnal, Me.dtpBalikJurnal.Value.Date)
            Me.prgBalikJurnal.Value = 0
            Me.prgBalikJurnal.Maximum = TBLJurnal.Rows.Count
            For Each rowJurnal In TBLJurnal.Rows
                Me.prgBalikJurnal.Value += 1
                Me.lblProgressBalikJurnal.Text = "Phase 1 of 2: Deleting reversal journal(data " + Me.prgBalikJurnal.Value.ToString + " of " + Me.prgBalikJurnal.Maximum.ToString + ")"
                If rowJurnal.Ref_Number.Contains(ModificationString) = True Then
                    DTAJurnal.DeleteQuery(rowJurnal.Primary)
                End If
            Next

            '2. buat reversal
            ModificationString = ModificationString + Format(Today, "yyyy-MM-dd") + ")"
            'load data tabel jurnal dan jurnal item
            TBLJurnal.Rows.Clear()
            Me.lblProgressBalikJurnal.Text = "Phase 2 of 2: Creating reversal journal"
            DTAJurnal.FillBy_Date(TBLJurnal, Me.dtpBalikJurnal.Value.Date)
            Me.prgBalikJurnal.Value = 0
            Me.prgBalikJurnal.Maximum = TBLJurnal.Rows.Count
            TBLItem.Rows.Clear()
            DTAItem.FillBy_Date(TBLItem, Me.dtpBalikJurnal.Value.Date)

            'untuk setiap jurnal:
            For Each rowJurnal In TBLJurnal.Rows
                Me.prgBalikJurnal.Value += 1
                'If Me.prgBalikJurnal.Value > 500 Then
                '    Exit Sub
                'End If
                Me.lblProgressBalikJurnal.Text = "Phase 2 of 2: Creating reversal journal(data " + Me.prgBalikJurnal.Value.ToString + " of " + Me.prgBalikJurnal.Maximum.ToString + ")"
                '1. ubah tipenya menjadi general jurnal
                rowJurnal.Tipe_Journal = 0
                rowJurnal.Tipe_Jurnal_2 = 0
                '2. modifikasi nomor jurnalnya
                rowJurnal.Ref_Number = rowJurnal.Ref_Number + ModificationString
                rowJurnal.Created_By = "USER"
                rowJurnal.Modified_By = "USER"

                counter = 0
                DummyString = rowJurnal.Ref_Number
                While ObjJurnal.ApakahSudahAda(rowJurnal.Ref_Number) = True
                    counter += 1
                    rowJurnal.Ref_Number = DummyString + "(" + counter.ToString + ")"
                End While

                '3. urutan = 0 -> untuk kalo nanti pake autonumber
                rowJurnal.Urutan = 0

                '4. balik debet - credit pada item
                TBLItem.DefaultView.RowFilter = "[Primary Jurnal] = " + rowJurnal.Primary.ToString
                For Each DV In TBLItem.DefaultView
                    rowItem = DV.Row
                    DummyVal = rowItem.Kredit
                    rowItem.Kredit = rowItem.Debet
                    rowItem.Debet = DummyVal

                    DummyVal = rowItem.Credit_Original
                    rowItem.Credit_Original = rowItem.Debet_Original
                    rowItem.Debet_Original = DummyVal
                Next

                '5. simpan header + item
                With rowJurnal
                    DTAJurnal.Insert(.Ref_Number, ._Date, .Urutan, .Tipe_Journal, .Primary_Document_Acuan, .Primary_Pihak_Acuan, .Desc, _
                    .Created_By, Now, .Modified_By, Now, Now, .primary_akun_pay_to, .primary_akun_difference, .paid, .exchange_rate, .paid_IDR, _
                    .AR_paid_IDR, .difference, .exchange_rate_gain_loss, .Tipe_Jurnal_2, .Primary_Currency, .Type_InOut, .Primary_Akun_Asset, .Primary_Asset, _
                    .Reference_Party, .payment_method, .primary_exhange_gain_loss)

                    PrimaryJurnal = DTAJurnal.ScalarQuery_GetPrimaryJurnal(.Ref_Number).Value

                    For Each DV In TBLItem.DefaultView
                        rowItem = DV.Row
                        With rowItem
                            DTAItem.Insert(PrimaryJurnal, .Primary_Akun, .Debet, .Kredit, .Remarks, .Debet_Original, .Credit_Original, .Ratio, .primary_dept, .primary_job)
                        End With
                    Next
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwBalikJurnal_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwBalikJurnal.RunWorkerCompleted
        Try
            Me.prgBalikJurnal.Value = 0
            Me.lblProgressBalikJurnal.Text = "Updated"
            MsgBox("Data updated")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Try
            MsgBox("Buat fungsi kecil untuk update item ot pada salary trans. Ini harus di update supaya overtime analisisnya bisa berjalan dengan baik.")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Try
            Dim i As New Purchase_Module.Form2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorALL = Core_Module.GetSMESSBackColorAll
            'i.MdiParent = Me
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Try
            Dim ij As New Inventory_Module.Form2
            'ij.MdiParent = Me
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate_ChkInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_ChkInv.Click 
        Try
            If MsgBox("Perform invoice data check vs ar card vs sales balance?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwChkInv.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwChkInv_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwChkInv.DoWork
        Dim TipeBalance As Sales_Module.ClsTransaksiSalesInvoice.EnumTipeSalesBalance = Sales_Module.ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim KartuAR_JenisTransaksiAsal As Sales_Module.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = Sales_Module.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
        Dim KartuAR_JenisKartu As Sales_Module.ClsTransaksi_Kartu_APAR.EnumJenisStockCard = Sales_Module.ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR
        Dim DTAInv As New DseUpdate_v2018TableAdapters.tbltrans_salesinvoiceTableAdapter
       Dim rowInv As DseUpdate_v2018.tbltrans_salesinvoiceRow
        Dim DTABalance As New DseUpdate_v2018TableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBLBalance As New DseUpdate_v2018.tbltrans_balance_salesDataTable
        Dim rowBalance As DseUpdate_v2018.tbltrans_balance_salesRow
        Dim DTAKartu As New DseUpdate_v2018TableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Dim TBLKartu As New DseUpdate_v2018.tbltransaksi_kartu_ar_apDataTable
        Dim rowKartu As DseUpdate_v2018.tbltransaksi_kartu_ar_apRow
        Dim ij As Double = 0
        Dim Ketemu1, Ketemu2 As Boolean

        Try
            Me.ObjCD.UpdateThreading()

            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTABalance.Connection.ConnectionString = Core_Module.GetConstring
            DTAKartu.Connection.ConnectionString = Core_Module.GetConstring

            TBLInv.Rows.Clear()
            DTAInv.Fill(TBLInv)
            DTABalance.Fill(TBLBalance, CInt(TipeBalance))
            DTAKartu.Fill(TBLKartu, CInt(KartuAR_JenisKartu), CInt(KartuAR_JenisTransaksiAsal))

            Me.PrgUpdate_ChkInv.Value = 0
            Me.PrgUpdate_ChkInv.Maximum = TBLInv.Rows.Count

            While ij <= TBLInv.Rows.Count - 1
                rowInv = TBLInv.Rows(ij)
                Me.PrgUpdate_ChkInv.Value += 1
                Me.lblProgress_ChkInv.Text = "Checking data " + Me.PrgUpdate_ChkInv.Value.ToString + _
                " of " + Me.PrgUpdate_ChkInv.Maximum.ToString + " - " + rowInv.No_Transaksi

                Ketemu1 = False
                Ketemu2 = False

                For jk As Double = 0 To TBLKartu.Rows.Count - 1
                    rowKartu = TBLKartu.Rows(jk)
                    If rowKartu.Primary_Transaksi_Asal = rowInv.Primary Then
                        TBLKartu.Rows.Remove(rowKartu)
                        Ketemu1 = True
                        Exit For
                    ElseIf rowKartu.Primary_Transaksi_Asal > rowInv.Primary Then
                        Exit For
                    End If
                Next

                If Ketemu1 = True Then
                    For jk As Double = 0 To TBLBalance.Rows.Count - 1
                        rowBalance = TBLBalance.Rows(jk)
                        If rowBalance.Primary_Invoice = rowInv.Primary Then
                            Ketemu2 = True
                            TBLBalance.Rows.Remove(rowBalance)
                            Exit For
                        ElseIf rowBalance.Primary_Invoice > rowInv.Primary Then
                            Exit For
                        End If
                    Next
                End If

                If Ketemu1 = True And Ketemu2 = True Then
                    'aman
                    TBLInv.Rows.Remove(rowInv)
                Else
                    'bermasalah karena kartu atau balance tidak lengkap
                    ij += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message + "ij: " + ij.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInv.Dispose()
            DTABalance.Dispose()
            TBLBalance.Dispose()
            DTAKartu.Dispose()
            TBLKartu.Dispose()
        End Try
    End Sub

    Private Sub BgwChkInv_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwChkInv.RunWorkerCompleted
        Try
            If TBLInv.Rows.Count > 0 Then
                Dim FrmInv As New FrmUpdate_ChkInv
                FrmInv.TBLInv = TBLInv
                FrmInv.ShowDialog()
            Else
                MsgBox("Everything is ok", MsgBoxStyle.Information, "Information")
            End If

            Me.PrgUpdate_ChkInv.Value = 0
            Me.lblProgress_ChkInv.Text = "Done"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim ij As New FrmVersionControl

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim ij As New Inventory_Module.FrmTrans_ProductionSchedule_Reorder_LMB

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwClearTable_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwClearTable.DoWork
        Try
            Me.ClearTable()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwClearTable_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwClearTable.RunWorkerCompleted
        Try
            Me.LoadTableName()
            Me.LblClearTable_Progress.Text = ""
            MsgBox("Table cleared.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Dim ij As New FrmUpdate_RepairAPAR

        Try
            ij.ShowDialog()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdate2018_PrimarySalesP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate2018_PrimarySalesP.Click
        Try
            If MsgBox("Update sales person data in sales invoice & sales return?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.BgwUpdPrimarySalesPerson.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUpdPrimarySalesPerson_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUpdPrimarySalesPerson.DoWork
        Dim dtaSO As New DseUpdate_v2018TableAdapters.tbltrans_soTableAdapter
        Dim tblSO As New DseUpdate_v2018.tbltrans_soDataTable
        Dim rowSO As DseUpdate_v2018.tbltrans_soRow
        Dim dtaSR_Quick As New DseUpdate_v2018TableAdapters.tbltrans_salesreturn_quickTableAdapter
        Dim tblSR_Quick As New DseUpdate_v2018.tbltrans_salesreturn_quickDataTable
        Dim rowSR_Quick As DseUpdate_v2018.tbltrans_salesreturn_quickRow

        Try
            Me.Cursor = Cursors.WaitCursor
            dtaSO.Connection.ConnectionString = Core_Module.GetConstring
            dtaSR_Quick.Connection.ConnectionString = Core_Module.GetConstring
            dtaSO.Fill(tblSO)
            dtaSR_Quick.Fill(tblSR_Quick)

            Me.ProgressBar17.Value = 0
            Me.ProgressBar17.Maximum = tblSO.Rows.Count
            For i As Integer = 0 To tblSO.Rows.Count - 1
                Me.ProgressBar17.Value += 1
                Me.Label80.Text = "Update: " & i & " of " & tblSO.Rows.Count - 1
                rowSO = tblSO.Rows(i)
                dtaSO.UpdateQuery_InvoiceItem_PrimarySalesPersonByPrimarySO(rowSO.primary_sales_person, rowSO.Primary)
                dtaSO.UpdateQuery_SalesReturnItem_PrimarySalesPersonByPrimarySO(rowSO.primary_sales_person, rowSO.Primary)
            Next

            For i As Integer = 0 To tblSR_Quick.Rows.Count - 1
                rowSR_Quick = tblSR_Quick.Rows(i)
                dtaSR_Quick.UpdateQuery_PrimarySalesPerson(rowSR_Quick.Primary_Sales_Person, rowSR_Quick.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUpdPrimarySalesPerson_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUpdPrimarySalesPerson.RunWorkerCompleted

        Try
            Me.Cursor = Cursors.Default
            Me.Label80.Text = "Update."
            MsgBox("Data updated.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdateSetVersion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateSetVersion.Click
        Dim dtaAutoUpdate As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        Dim DirPath As String = ""
        Dim ProductName As String = ""
        Dim ObjVersion As New Core_Module.Clstrans_VersionControl

        Try
            If MsgBox("Set dummy data to my.settings?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                dtaAutoUpdate.Connection.ConnectionString = Core_Module.GetConstring

                'update my setting
                DirPath = My.Application.Info.DirectoryPath
                ProductName = My.Application.Info.ProductName
                My.Settings.update_link_dummy = Me.GetServerPath()
                My.Settings.version_dummy = ObjVersion.GetOperatingVersion()
                My.Settings.Save()
                'update di db
                dtaAutoUpdate.Delete()
                dtaAutoUpdate.Insert(DirPath, "", ProductName)
                ObjVersion.SetOperatingVersion("1.6.40")

                MsgBox("data updated", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaAutoUpdate.Dispose()
            ObjVersion = Nothing
        End Try
    End Sub

    Private Sub btnUpdateGetVersion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateGetVersion.Click
        Dim dtaAutoUpdate As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        Dim DirPath As String = ""
        Dim ProductName As String = ""
        Dim ObjVersion As New Core_Module.Clstrans_VersionControl

        Try
            If MsgBox("Set dummy data from my.settings?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                dtaAutoUpdate.Connection.ConnectionString = Core_Module.GetConstring
                'update di db
                DirPath = My.Application.Info.DirectoryPath
                ProductName = My.Application.Info.ProductName
                dtaAutoUpdate.Delete()
                dtaAutoUpdate.Insert(DirPath, My.Settings.update_link_dummy, ProductName)
                ObjVersion.SetOperatingVersion(My.Settings.version_dummy)

                'update my setting
                My.Settings.update_link_dummy = ""
                My.Settings.version_dummy = ""
                My.Settings.Save()

                MsgBox("data updated", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaAutoUpdate.Dispose()
            ObjVersion = Nothing
        End Try
    End Sub

    Private Function GetServerPath() As String
        Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        Dim Dummy As String = ""
        Try
            Dummy = DTA.ScalarQuery_GetServerPath()
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub btnUpdateSalaryOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateSalaryOverride.Click
        Try
            Dim ij As New HR.FrmTrans_Gaji_Override_Create
            ij.UserCode = FrmLogin.UsernameTextBox.Text
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUpdatePresentasiPayroll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdatePresentasiPayroll.Click
        Dim DTAMasterKaryawan As New DseUpdatePayrollTableAdapters.tblmaster_karyawanTableAdapter
        Dim TBLMasterKaryawan As New DseUpdatePayroll.tblmaster_karyawanDataTable
        Dim TBLMasterKaryawanD As New DseUpdatePayroll.tblmaster_karyawanDataTable
        Dim rowMasterKaryawan, rowMasterKaryawanD As DseUpdatePayroll.tblmaster_karyawanRow
        Dim Idx As Double = 0
        Dim DTAKeluarga As New DseUpdatePayrollTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim DTACompanyData As New DseUpdatePayrollTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBLCompanyData As New DseUpdatePayroll.tbltrans_company_dataDataTable
        Dim rowCompanyData As DseUpdatePayroll.tbltrans_company_dataRow
        Dim DTACompanyDataItem As New DseUpdatePayrollTableAdapters.tbltrans_company_data_itemTableAdapter
        Dim TBLCompanyDataItem As New DseUpdatePayroll.tbltrans_company_data_itemDataTable
        Dim rowCompanyDataItem As DseUpdatePayroll.tbltrans_company_data_itemRow
        Dim CompanyName As String = "PT. Demo"
        Dim CompanyAdd As String = "Jl. Dahlia IV Blok D1 no. "
        Dim CompanyNPWPNo As String = "09.435.036.0-428.000"

        Try
            If MsgBox("Start data update. Warning you can't reverse this process.", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTAMasterKaryawan.Connection.ConnectionString = Core_Module.GetConstring
                DTAKeluarga.Connection.ConnectionString = Core_Module.GetConstring
                DTACompanyData.Connection.ConnectionString = DTAKeluarga.Connection.ConnectionString
                DTACompanyDataItem.Connection.ConnectionString = DTAKeluarga.Connection.ConnectionString

                '1. Nama karyawan
                DTAMasterKaryawan.Fill(TBLMasterKaryawan)
                Me.LoadDataKaryawan(TBLMasterKaryawanD)
                If TBLMasterKaryawanD.Rows.Count < TBLMasterKaryawan.Rows.Count Then
                    MsgBox("Please update 'Load Data Karyawan' function, not enough employee is registered.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    While Idx <= TBLMasterKaryawan.Rows.Count - 1
                        rowMasterKaryawan = TBLMasterKaryawan.Rows(Idx)
                        rowMasterKaryawanD = TBLMasterKaryawanD.Rows(Idx)

                        With rowMasterKaryawanD
                            DTAMasterKaryawan.UpdateQuery1(.Full_Name, .npwp_no, .npwp_address, .current_address, .email, .kpa_astek_no, .kpj_no, _
                            .Bank_Name, .Bank_Branch, .Bank_Account_No, .Bank_Account_Holder, .Nama, .id_card_no, _
                            .id_card_expiry, .id_card_address, .picture_path, rowMasterKaryawan.Primary)

                            DTAMasterKaryawan.UpdateQuery4(CompanyNPWPNo, CompanyName, .npwp_no, .Nama, .npwp_address, rowMasterKaryawan.Primary)
                        End With

                        Idx += 1
                    End While
                End If

                '2. company data
                DTACompanyData.Fill(TBLCompanyData)
                rowCompanyData = TBLCompanyData.Rows(0)
                With rowCompanyData
                    DTACompanyData.Update(CompanyName, CompanyAdd + "1", "", "", "", "", "Bandung", CompanyNPWPNo, #1/1/2008#, CompanyAdd + "1", "", .Primary)
                End With
                DTACompanyDataItem.Fill(TBLCompanyDataItem)
                Idx = 0
                For Each rowCompanyDataItem In TBLCompanyDataItem.Rows
                    Idx += 1
                    With rowCompanyDataItem
                        DTACompanyDataItem.Update(CompanyAdd + CInt(Idx).ToString, "", "", "", "", "", .Primary)
                    End With
                Next

                '3. nama anggota keluarga, alamat, dan data pribadi lainnya
                DTAKeluarga.Delete()
                MsgBox("blm coding: update keluarga karyawan")

                '4. operating version + auto update path
                DTAMasterKaryawan.UpdateQuery2()
                DTAMasterKaryawan.UpdateQuery3()

                MsgBox("Data updated", MsgBoxStyle.Information, "Information")

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAMasterKaryawan.Dispose()
            TBLMasterKaryawan.Dispose()
            TBLMasterKaryawanD.Dispose()
        End Try
    End Sub

    Private Sub AddBogusRow(ByVal vTBLMasterKaryawan As DseUpdatePayroll.tblmaster_karyawanDataTable, ByVal vNama As String)
        Dim rowMasterKaryawan As DseUpdatePayroll.tblmaster_karyawanRow
        Dim Dummy As String = ""
        Dim PicPath As String = "C:\Users\Yulimus-ASUS\Dropbox (Artha Jina)\Presentasi dan Demo\support data\foto karyawan"
        Dim DummyAdd As String = "Jl. Abdulrahman Saleh no. "

        Try
            rowMasterKaryawan = vTBLMasterKaryawan.Newtblmaster_karyawanRow

            With rowMasterKaryawan
                .Bank_Account_Holder = vNama
                .Bank_Account_No = 1138016827 + vTBLMasterKaryawan.Rows.Count
                .Bank_Branch = "Bandung"
                .Bank_Name = "PERMATA"
                .current_address = DummyAdd + vTBLMasterKaryawan.Rows.Count.ToString
                .email = vNama.ToLower + "@ptdemo.com"
                .First_Name = vNama
                .Full_Name = vNama
                .id_card_address = DummyAdd + vTBLMasterKaryawan.Rows.Count.ToString
                .id_card_expiry = #1/1/2050#
                .id_card_no = (3275050101700000 + vTBLMasterKaryawan.Rows.Count).ToString
                .kpa_astek_no = .id_card_no
                .kpj_no = .id_card_no
                .Last_Name = ""
                .Middle_Name = ""
                .Nama = vNama
                .npwp_address = .current_address
                Dummy = vTBLMasterKaryawan.Rows.Count.ToString
                While Dummy.Length < 3
                    Dummy = "0" + Dummy
                End While
                .npwp_no = "00.636.426.1-017." + Dummy
                If Math.IEEERemainder(vTBLMasterKaryawan.Rows.Count, 10) < 0 Then
                    Dummy = Math.IEEERemainder(vTBLMasterKaryawan.Rows.Count, 10) + 10
                Else
                    Dummy = Math.IEEERemainder(vTBLMasterKaryawan.Rows.Count, 10)
                End If
                If Dummy = "0" Then
                    Dummy = "1"
                End If
                While Dummy.Length < 2
                    Dummy = "0" + Dummy
                End While
                .picture_path = PicPath + "\E-" + Dummy + ".jpg"
                .Prefix = ""
                .Primary = -1
                .Suffix = "'"
            End With

            vTBLMasterKaryawan.Rows.Add(rowMasterKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataKaryawan(ByVal vTBLMasterKaryawan As DseUpdatePayroll.tblmaster_karyawanDataTable)
        Try
            vTBLMasterKaryawan.Rows.Clear()
            Me.AddBogusRow(vTBLMasterKaryawan, "Aang Suparman")
            Me.AddBogusRow(vTBLMasterKaryawan, "Achmad Jufrianto")
            Me.AddBogusRow(vTBLMasterKaryawan, "Adjat Sudradjat")
            Me.AddBogusRow(vTBLMasterKaryawan, "Adolf Kabo")
            Me.AddBogusRow(vTBLMasterKaryawan, "Agung Setyabudi")
            Me.AddBogusRow(vTBLMasterKaryawan, "Agus Indra Kurniawan")
            Me.AddBogusRow(vTBLMasterKaryawan, "Agus Salim")
            Me.AddBogusRow(vTBLMasterKaryawan, "Agusman Riyadi")
            Me.AddBogusRow(vTBLMasterKaryawan, "Agus Yuwono")
            Me.AddBogusRow(vTBLMasterKaryawan, "Ahmad Bustomi")
            Me.AddBogusRow(vTBLMasterKaryawan, "Airlangga Sucipto")
            Me.AddBogusRow(vTBLMasterKaryawan, "Alan Martha")
            Me.AddBogusRow(vTBLMasterKaryawan, "Alexander Pulalo")
            Me.AddBogusRow(vTBLMasterKaryawan, "Alexander Saununu")
            Me.AddBogusRow(vTBLMasterKaryawan, "Al Hadad")
            Me.AddBogusRow(vTBLMasterKaryawan, "Andik Vermansyah")
            Me.AddBogusRow(vTBLMasterKaryawan, "Anjas Asmara")
            Me.AddBogusRow(vTBLMasterKaryawan, "Aji Santoso")
            Me.AddBogusRow(vTBLMasterKaryawan, "Ardan Aras")
            Me.AddBogusRow(vTBLMasterKaryawan, "Aris Budi Prasetyo")
            Me.AddBogusRow(vTBLMasterKaryawan, "Aris Indarto")
            Me.AddBogusRow(vTBLMasterKaryawan, "Arief Hidayat")
            Me.AddBogusRow(vTBLMasterKaryawan, "Arif Suyono")
            Me.AddBogusRow(vTBLMasterKaryawan, "Arthur Irawan")
            Me.AddBogusRow(vTBLMasterKaryawan, "Asri Akbar")
            Me.AddBogusRow(vTBLMasterKaryawan, "Atep")
            Me.AddBogusRow(vTBLMasterKaryawan, "Bambang Nurdiansyah")
            Me.AddBogusRow(vTBLMasterKaryawan, "Bambang Pamungkas")
            Me.AddBogusRow(vTBLMasterKaryawan, "Benyamin Van Breukelen")
            Me.AddBogusRow(vTBLMasterKaryawan, "Bima Sakti")
            Me.AddBogusRow(vTBLMasterKaryawan, "Budi Sudarsono")
            Me.AddBogusRow(vTBLMasterKaryawan, "Boaz Solossa")
            Me.AddBogusRow(vTBLMasterKaryawan, "Boy Jati Asmara")
            Me.AddBogusRow(vTBLMasterKaryawan, "Cecep Supriatna")
            Me.AddBogusRow(vTBLMasterKaryawan, "Christian Gonzales")
            Me.AddBogusRow(vTBLMasterKaryawan, "Christian Worabay")
            Me.AddBogusRow(vTBLMasterKaryawan, "David Laly")
            Me.AddBogusRow(vTBLMasterKaryawan, "Dede Sulaiman")
            Me.AddBogusRow(vTBLMasterKaryawan, "Didik Darmadi")
            Me.AddBogusRow(vTBLMasterKaryawan, "Diego Michiels")
            Me.AddBogusRow(vTBLMasterKaryawan, "Dominggus Nowenik")
            Me.AddBogusRow(vTBLMasterKaryawan, "Doni Latuperisa")
            Me.AddBogusRow(vTBLMasterKaryawan, "Edi Harto")
            Me.AddBogusRow(vTBLMasterKaryawan, "Eduard Ivakdalam")
            Me.AddBogusRow(vTBLMasterKaryawan, "Elie Aiboy")
            Me.AddBogusRow(vTBLMasterKaryawan, "Endang Witarsa")
            Me.AddBogusRow(vTBLMasterKaryawan, "Eri Eriawan")
            Me.AddBogusRow(vTBLMasterKaryawan, "Eric Ardhyanto")
            Me.AddBogusRow(vTBLMasterKaryawan, "Erol FX Iba")
            Me.AddBogusRow(vTBLMasterKaryawan, "Evan Dimas D")
            Me.AddBogusRow(vTBLMasterKaryawan, "Fachry Husaini")
            Me.AddBogusRow(vTBLMasterKaryawan, "Felix Arifin Nugroho")
            Me.AddBogusRow(vTBLMasterKaryawan, "Ferrel Raymond Hattu")
            Me.AddBogusRow(vTBLMasterKaryawan, "Ferry Rotinsulu")
            Me.AddBogusRow(vTBLMasterKaryawan, "Firman Utina")
            Me.AddBogusRow(vTBLMasterKaryawan, "Greg Nwokolo")
            Me.AddBogusRow(vTBLMasterKaryawan, "Budi Sudarsono")
            Me.AddBogusRow(vTBLMasterKaryawan, "Gunaryo")
            Me.AddBogusRow(vTBLMasterKaryawan, "Hadi Surento")
            Me.AddBogusRow(vTBLMasterKaryawan, "Hamka Hamzah")
            Me.AddBogusRow(vTBLMasterKaryawan, "Hariono")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub  

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim i As New Inventory_Module.FrmTransaksiStockCard2_UpdateMonthlyClosing
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim i As New Inventory_Module.FrmTrans_ManufacturingOrder_SetFinishWOStatus
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim i As New Inventory_Module.FrmTransaksiStockCard2_CheckProductQtyMonthlyClosing
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnv2021_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnv2021.Click
        Dim ij As New Core_Module.Frm2021

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class