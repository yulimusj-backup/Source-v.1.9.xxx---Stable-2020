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

    Public BackColorAll As Color = Core_Module.GetSMESSColor

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
End Class