Option Explicit On
Option Strict Off

Partial Public Class ClsTransaksiJournal_Item
    Public Function SimpanItem_Asset(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_Acquisition_Revaluation_Other, ByVal vDGV_Pay_Detail As DataGridView) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA_Payment_Detail As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_payment_detailTableAdapter
        Dim DefaultPrecision As Integer = 0

        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
                     CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebet_OriginalCurr.Index).Value), _
                    CDec(row.Cells(Frm.ColCredit_OriginalCurrency.Index).Value), CDec(row.Cells(Frm.ColExchangeRate.Index).Value), _
                    row.Cells(Frm.ColPrimaryDept.Index).Value, row.Cells(Frm.ColPrimaryJob.Index).Value)
                End If
            Next

            For Each row In vDGV_Pay_Detail.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.COlItemPay_Remarks.Index).Value) = True Then
                    row.Cells(Frm.COlItemPay_Remarks.Index).Value = ""
                End If
                If row.Cells(Frm.COlItemPay_Account.Index).Value <> "" Then
                    DTA_Payment_Detail.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.COlItemPay_AmountOriginal.Index).Value), row.Cells(Frm.COlItemPay_Remarks.Index).Value, _
                    CDec(row.Cells(Frm.ColItemPay_Amount.Index).Value), CDbl(row.Cells(Frm.COlItemPay_PrimaryDept.Index).Value), CDbl(row.Cells(Frm.COlItemPay_PrimaryJob.Index).Value), _
                    CDbl(row.Cells(Frm.COlItemPay_PrimaryAkun.Index).Value), CDec(row.Cells(Frm.ColItemPay_ExRate.Index).Value))
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            DTA_Payment_Detail.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem_Other(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_Acquisition_Revaluation_Other, ByVal vDGVItem_payment As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim DTA_Pay_detail As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_payment_detail1TableAdapter
        Dim TBL_Pay_Detail As New DseTransaksiJournal.tbltrans_jurnal_item_payment_detail1DataTable
        Dim row_Pay_Detail As DseTransaksiJournal.tbltrans_jurnal_item_payment_detail1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjAcc As New ClsMasterAkun
        Dim objDept As New Core_Module.ClsMasterDepartement
        Dim objJob As New Core_Module.ClsMasterJob

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVItem_payment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem_payment.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        DefPrecBase = ObjCur.GetDefaultPrecission(ObjAcc.GetPrimaryCurrency(.Primary_Akun))
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColPrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(Frm.ColAccount.Index).Value = .Account_Code
                        rowDGV.Cells(Frm.ColDescAccount.Index).Value = .Description
                        rowDGV.Cells(Frm.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet, DefPrecBase)
                        rowDGV.Cells(Frm.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit, DefPrecBase)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(.Debet_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(.Credit_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColPrimaryCurrency.Index).Value = .primary_currency
                        rowDGV.Cells(Frm.ColOriginalCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Frm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio, Core_Module.DefPrecExchangeRate)
                        rowDGV.Cells(Frm.ColDept.Index).Value = .Dept_Code
                        rowDGV.Cells(Frm.ColPrimaryDept.Index).Value = .primary_dept
                        rowDGV.Cells(Frm.ColPrimaryJob.Index).Value = .primary_job
                        rowDGV.Cells(Frm.ColJob.Index).Value = .Job_Code
                    End With
                Next
            End If

            If DTA_Pay_detail.Fill(TBL_Pay_Detail, vPrimaryTransaksi) > 0 Then
                Dim DefPrec As Integer = 0
                For Each row_Pay_Detail In TBL_Pay_Detail.Rows
                    With row_Pay_Detail
                        vDGVItem_payment.Rows.Add()
                        rowDGV = vDGVItem_payment.Rows(vDGVItem_payment.Rows.Count - 1)

                        With rowDGV
                            .Cells(Frm.COlItemPay_Account.Index).Value = row_Pay_Detail.Account_Code
                            .Cells(Frm.ColItemPay_Amount.Index).Value = ObjCD.DigitGrouping(row_Pay_Detail.Amount, DefPrecBase)
                            .Cells(Frm.COlItemPay_AmountOriginal.Index).Value = ObjCD.DigitGrouping(row_Pay_Detail.Amount_original, _
                            row_Pay_Detail.Default_Precision)
                            .Cells(Frm.COlItemPay_DeptCode.Index).Value = row_Pay_Detail.Dept_Code
                            .Cells(Frm.COlItemPay_Desc.Index).Value = row_Pay_Detail.Description
                            .Cells(Frm.ColItemPay_ExRate.Index).Value = ObjCD.DigitGrouping(row_Pay_Detail.exchange_rate, Core_Module.DefPrecExchangeRate)
                            .Cells(Frm.COlItemPay_JobCode.Index).Value = row_Pay_Detail.Jo_Code
                            .Cells(Frm.COlItemPay_PrimaryAkun.Index).Value = row_Pay_Detail.Primary_Akun
                            .Cells(Frm.COlItemPay_PrimaryDept.Index).Value = row_Pay_Detail.Primary_Dept
                            .Cells(Frm.COlItemPay_PrimaryJob.Index).Value = row_Pay_Detail.Primary_Job
                            .Cells(Frm.COlItemPay_Remarks.Index).Value = row_Pay_Detail.Remarks
                            .Cells(Frm.ColItemPay_Curr.Index).Value = row_Pay_Detail.Currency_Code
                        End With
                    End With
                Next
            End If

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem_payment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            DTA_Pay_detail.Dispose()
            TBL_Pay_Detail = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            ObjAcc = Nothing
            objDept = Nothing
            objJob = Nothing
        End Try
    End Function
End Class
