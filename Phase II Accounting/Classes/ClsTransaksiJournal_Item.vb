Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal_Item
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
  ByVal Frm As FrmTransaksiJournal_CashBank, ByVal vDGVItem_payment As DataGridView) As Boolean
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

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal_AssetDepreciation, ByVal vDGVAsset As DataGridView) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA_Payment_Detail As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_payment_detailTableAdapter
        Dim DefaultPrecision As Integer = 0

        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColAsset_Remarks.Index).Value) = True Then
                    row.Cells(Frm.ColAsset_Remarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColAsset_Account.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColAsset_PrimaryAkun.Index).Value), _
                     CDec(row.Cells(Frm.ColAsset_Debet.Index).Value), CDec(row.Cells(Frm.ColAsset_Kredit.Index).Value), _
                    row.Cells(Frm.ColAsset_Remarks.Index).Value, CDec(row.Cells(Frm.ColAsset_Debet_OriginalCurr.Index).Value), _
                    CDec(row.Cells(Frm.ColAsset_Credit_OriginalCurrency.Index).Value), CDec(row.Cells(Frm.ColAsset_ExchangeRate.Index).Value), _
                    row.Cells(Frm.ColAsset_PrimaryDept.Index).Value, row.Cells(Frm.ColAsset_PrimaryJob.Index).Value)
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

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_CashBank, ByVal vDGV_Pay_Detail As DataGridView) As Boolean
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

    '' untuk sales journal
    ''
    Public Function IsiDGVItemSalesJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_Sales, ByVal vDGVItemOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecBase As Integer = 0
        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_other1TableAdapter
        Dim TBL1 As New DseTransaksiJournal.tbltrans_jurnal_item_other1DataTable
        Dim row1 As DseTransaksiJournal.tbltrans_jurnal_item_other1Row

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVItemOther.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemOther.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, CDec(vPrimaryTransaksi)) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColPrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(Frm.ColAccount.Index).Value = .Account_Code
                        rowDGV.Cells(Frm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(Frm.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet, DefPrecBase)
                        rowDGV.Cells(Frm.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit, DefPrecBase)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColDebetOriginal.Index).Value = ObjCD.DigitGrouping(.Debet_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColCreditOriginal.Index).Value = ObjCD.DigitGrouping(.Credit_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColPrimaryCurrency.Index).Value = .primary_currency
                        rowDGV.Cells(Frm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Frm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio, Core_Module.DefPrecExchangeRate)
                    End With
                Next
            End If

            'isikan tabel item other
            DTA1.FillBy_PrimaryTrans(TBL1, vPrimaryTransaksi)
            For Each row1 In TBL1.Rows
                vDGVItemOther.Rows.Add()
                rowDGV = vDGVItemOther.Rows(vDGVItemOther.Rows.Count - 1)
                With row1
                    rowDGV.Cells(Frm.ColAccountDiff_AccountCode.Index).Value = row1.Account_Code
                    rowDGV.Cells(Frm.ColAccountDiff_Desc.Index).Value = row1.Description
                    rowDGV.Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value = ObjCD.DigitGrouping(row1.exchange_rate, Core_Module.DefPrecExchangeRate)
                    rowDGV.Cells(Frm.ColAccountDiff_OriginalCurrency.Index).Value = row1.Currency_Code
                    rowDGV.Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value = .primary_akun
                    rowDGV.Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value = .primary_currency
                    rowDGV.Cells(Frm.ColAccountDiff_Remarks.Index).Value = .remarks
                    rowDGV.Cells(Frm.ColAccountDiff_Value.Index).Value = ObjCD.DigitGrouping(.value, DefPrecBase)
                    rowDGV.Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value = ObjCD.DigitGrouping(.value_original, .Default_Precision)
                End With
            Next

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemOther.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    '' untuk purchase journal
    ''
    Public Function IsiDGVItemPurchaseJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
     ByVal Frm As FrmTransaksiJournal_Purchase, ByVal vDGVItemOther As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_other1TableAdapter
        Dim TBL1 As New DseTransaksiJournal.tbltrans_jurnal_item_other1DataTable
        Dim row1 As DseTransaksiJournal.tbltrans_jurnal_item_other1Row

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVItemOther.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemOther.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColPrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(Frm.ColAccount.Index).Value = .Account_Code
                        rowDGV.Cells(Frm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(Frm.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet, DefPrecBase)
                        rowDGV.Cells(Frm.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit, DefPrecBase)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColDebetOriginal.Index).Value = ObjCD.DigitGrouping(.Debet_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColCreditOriginal.Index).Value = ObjCD.DigitGrouping(.Credit_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColPrimaryCurrency.Index).Value = .primary_currency
                        rowDGV.Cells(Frm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Frm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio, Core_Module.DefPrecExchangeRate)
                    End With
                Next
            End If

            'isikan tabel item other
            DTA1.FillBy_PrimaryTrans(TBL1, vPrimaryTransaksi)
            For Each row1 In TBL1.Rows
                vDGVItemOther.Rows.Add()
                rowDGV = vDGVItemOther.Rows(vDGVItemOther.Rows.Count - 1)
                With row1
                    rowDGV.Cells(Frm.ColAccountDiff_AccountCode.Index).Value = row1.Account_Code
                    rowDGV.Cells(Frm.ColAccountDiff_Desc.Index).Value = row1.Description
                    rowDGV.Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value = ObjCD.DigitGrouping(row1.exchange_rate, Core_Module.DefPrecExchangeRate)
                    rowDGV.Cells(Frm.ColAccountDiff_OriginalCurrency.Index).Value = row1.Currency_Code
                    rowDGV.Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value = .primary_akun
                    rowDGV.Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value = .primary_currency
                    rowDGV.Cells(Frm.ColAccountDiff_Remarks.Index).Value = .remarks
                    rowDGV.Cells(Frm.ColAccountDiff_Value.Index).Value = ObjCD.DigitGrouping(.value, DefPrecBase)
                    rowDGV.Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value = ObjCD.DigitGrouping(.value_original, .Default_Precision)
                End With
            Next

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemOther.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Function GetPayment(ByVal vPrimaryJurnal As Double, ByVal vPrimaryDocumentAcuan As Double, _
    ByVal vTipeDokumentAcuan As ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vTipePayment As ClsTransaksiSalesBalance.EnumTipePayment) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter

        Try
            Return DTA.ScalarQuery_GetPayment(vPrimaryJurnal, vTipePayment, _
            vPrimaryDocumentAcuan, vTipeDokumentAcuan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).HasValue Then
                Return DTA.ScalarQuery_GetJmlJournalByPrimaryAkun(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem_EndOfMonth(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal_EndOfMonth, ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, _
    ByVal vDGVItemAP As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi1TableAdapter
        Dim TBL1 As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi1DataTable
        Dim row1 As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi1Row

        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_arTableAdapter
        Dim TBL2 As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arDataTable
        Dim row2 As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_arRow

        Dim DTA3 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_apTableAdapter
        Dim TBL3 As New DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apDataTable
        Dim row3 As DseTransaksiJournal.tbltrans_jurnal_item_revaluasi_apRow

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemAR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVCurrency.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVCurrency.Rows.Clear()
            vDGVItemAR.Rows.Clear()

            'isikan tabel
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
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
                    End With
                Next
            End If

            DTA1.Fill(TBL1, vPrimaryTransaksi)
            For Each row1 In TBL1.Rows
                vDGVCurrency.Rows.Add()
                rowDGV = vDGVCurrency.Rows(vDGVCurrency.Rows.Count - 1)

                With rowDGV
                    .Cells(Frm.ColItemExc_Currency.Index).Value = row1.Currency_Code
                    .Cells(Frm.ColItemExc_ExchangeRate.Index).Value = ObjCD.DigitGrouping(row1.exchange_rate, Core_Module.DefPrecExchangeRate)
                    .Cells(Frm.ColItemExc_PrimaryCurrency.Index).Value = row1.primary_currency
                    .Cells(Frm.ColItemExc_Remarks.Index).Value = row1.remarks
                End With
            Next

            vDGVItemAR.Rows.Clear()
            DTA2.FillBy_PrimaryJurnal(TBL2, vPrimaryTransaksi)
            For Each row2 In TBL2.Rows
                vDGVItemAR.Rows.Add()
                rowDGV = vDGVItemAR.Rows(vDGVItemAR.Rows.Count - 1)

                With rowDGV
                    .Cells(Frm.ColItemAR_ExchangeRateLama.Index).Value = row2.exchange_rate_lama
                    .Cells(Frm.ColItemAR_PrimaryAkunAR.Index).Value = row2.primary_akun_ar
                    .Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value = row2.primary_currency
                    .Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value = row2.primary_transaksi_asal
                    .Cells(Frm.ColItemAR_Saldo.Index).Value = row2.saldo
                    .Cells(Frm.ColItemAR_SaldoOriginal.Index).Value = row2.saldo_original
                    .Cells(Frm.ColItemAR_Selisih.Index).Value = row2.selisih
                    .Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value = row2.tipe_transaksi_asal
                    .Cells(Frm.ColItemAR_Tanggal.Index).Value = row2.tanggal_acuan
                    .Cells(Frm.ColItemAR_PaymentTermAcuan.Index).Value = row2.payment_term_acuan
                    .Cells(Frm.ColItemAR_PrimaryPihakAcuan.Index).Value = row2.primary_pihak_acuan
                End With
            Next

            vDGVItemAP.Rows.Clear()
            DTA3.FillBy_PrimaryTransaksi(TBL3, vPrimaryTransaksi)
            For Each row3 In TBL3.Rows
                vDGVItemAP.Rows.Add()
                rowDGV = vDGVItemAP.Rows(vDGVItemAP.Rows.Count - 1)

                With rowDGV
                    .Cells(Frm.ColItemAP_ExchangeRateLama.Index).Value = row3.exchange_rate_lama
                    .Cells(Frm.ColItemAP_PrimaryAkunAP.Index).Value = row3.primary_akun_ap
                    .Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value = row3.primary_currency
                    .Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value = row3.primary_transaksi_asal
                    .Cells(Frm.ColItemAP_Saldo.Index).Value = row3.saldo
                    .Cells(Frm.ColItemAP_SaldoOriginal.Index).Value = row3.saldo_original
                    .Cells(Frm.ColItemAP_Selisih.Index).Value = row3.selisih
                    .Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value = row3.tipe_transaksi_asal
                    .Cells(Frm.ColItemAP_Tanggal.Index).Value = row3.tanggal_acuan
                    .Cells(Frm.ColItemAP_PaymentTermAcuan.Index).Value = row3.payment_term_acuan
                    .Cells(Frm.ColItemAP_PrimaryPihakAcuan.Index).Value = row3.primary_pihak_acuan
                End With
            Next

            'set autosize
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVCurrency.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            DTA3.Dispose()
            TBL3.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmTransaksiJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
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

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function IsiDGVItem_Asset(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal_AssetDepreciation, ByVal vDGVItemAsset As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item1TableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_item1DataTable
        Dim row As DseTransaksiJournal.tbltrans_jurnal_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColAsset_PrimaryAkun.Index).Value = .Primary_Akun
                        rowDGV.Cells(Frm.ColAsset_Account.Index).Value = .Account_Code
                        rowDGV.Cells(Frm.ColAsset_DescAccount.Index).Value = .Description
                        rowDGV.Cells(Frm.ColAsset_Debet.Index).Value = ObjCD.DigitGrouping(.Debet, DefPrecBase)
                        rowDGV.Cells(Frm.ColAsset_Kredit.Index).Value = ObjCD.DigitGrouping(.Kredit, DefPrecBase)
                        rowDGV.Cells(Frm.ColAsset_Remarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColAsset_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(.Debet_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColAsset_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(.Credit_Original, .Default_Precision)
                        rowDGV.Cells(Frm.ColAsset_PrimaryCurrency.Index).Value = .primary_currency
                        rowDGV.Cells(Frm.ColAsset_OriginalCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Frm.ColAsset_ExchangeRate.Index).Value = ObjCD.DigitGrouping(.Ratio, Core_Module.DefPrecExchangeRate)
                        rowDGV.Cells(Frm.ColAsset_Dept.Index).Value = .Dept_Code
                        rowDGV.Cells(Frm.ColAsset_PrimaryDept.Index).Value = .primary_dept
                        rowDGV.Cells(Frm.ColAsset_PrimaryJob.Index).Value = .primary_job
                        rowDGV.Cells(Frm.ColAsset_Job.Index).Value = .Job_Code
                    End With
                Next
            End If

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function SimpanItem_EndOfMonth(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
  ByVal Frm As FrmTransaksiJournal_EndOfMonth, ByVal vDGVCurrency As DataGridView, ByVal vDGVItemAR As DataGridView, _
    ByVal vDGVItemAP As DataGridView) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasiTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_arTableAdapter
        Dim DTA3 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_apTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            'simpan data item jurnal
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
                    CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebet_OriginalCurr.Index).Value), _
                    CDec(row.Cells(Frm.ColCredit_OriginalCurrency.Index).Value), CDec(row.Cells(Frm.ColExchangeRate.Index).Value), -1, -1)
                End If
            Next

            'simpan data currency
            For Each row In vDGVCurrency.Rows
                DTA1.InsertQuery(vPrimaryTransaksi, row.Cells(Frm.ColItemExc_PrimaryCurrency.Index).Value, _
                row.Cells(Frm.ColItemExc_ExchangeRate.Index).Value, row.Cells(Frm.ColItemExc_Remarks.Index).Value)
            Next

            'simpan data revaluasi ar
            For Each row In vDGVItemAR.Rows
                DTA2.Insert(vPrimaryTransaksi, row.Cells(Frm.ColItemAR_PrimaryCurrency.Index).Value, _
                CDec(row.Cells(Frm.ColItemAR_ExchangeRateLama.Index).Value), CDec(row.Cells(Frm.ColItemAR_Saldo.Index).Value), _
                CDec(row.Cells(Frm.ColItemAR_SaldoOriginal.Index).Value), CDec(row.Cells(Frm.ColItemAR_TipeTransaksiAsal.Index).Value), _
                CDbl(row.Cells(Frm.ColItemAR_PrimaryTransaksiAsal.Index).Value), CDec(row.Cells(Frm.ColItemAR_PrimaryAkunAR.Index).Value), _
                CDec(row.Cells(Frm.ColItemAR_Selisih.Index).Value), CDec(row.Cells(Frm.ColItemAR_PrimaryPihakAcuan.Index).Value), _
                CDec(row.Cells(Frm.ColItemAR_PaymentTermAcuan.Index).Value), row.Cells(Frm.ColItemAR_Tanggal.Index).Value)
            Next

            'simpan data revaluasi ap
            For Each row In vDGVItemAP.Rows
                DTA3.Insert(vPrimaryTransaksi, row.Cells(Frm.ColItemAP_PrimaryCurrency.Index).Value, _
                CDec(row.Cells(Frm.ColItemAP_ExchangeRateLama.Index).Value), CDec(row.Cells(Frm.ColItemAP_Saldo.Index).Value), _
                CDec(row.Cells(Frm.ColItemAP_SaldoOriginal.Index).Value), CDec(row.Cells(Frm.ColItemAP_TipeTransaksiAsal.Index).Value), _
                CDbl(row.Cells(Frm.ColItemAP_PrimaryTransaksiAsal.Index).Value), CDec(row.Cells(Frm.ColItemAP_PrimaryAkunAP.Index).Value), _
                CDec(row.Cells(Frm.ColItemAP_Selisih.Index).Value), CDec(row.Cells(Frm.ColItemAP_PrimaryPihakAcuan.Index).Value), _
                CDec(row.Cells(Frm.ColItemAP_PaymentTermAcuan.Index).Value), row.Cells(Frm.ColItemAP_Tanggal.Index).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
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
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA1 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasiTableAdapter
        Dim DTA2 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_arTableAdapter
        Dim DTA3 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_revaluasi_apTableAdapter
        Dim DTA4 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_payment_detailTableAdapter
        Dim DTA5 As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

        Try
            'hapus data item
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Delete(vPrimaryTransaksi)
            DTA1.DeleteQuery(vPrimaryTransaksi)
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.Delete(vPrimaryTransaksi)
            DTA4.Delete(vPrimaryTransaksi)
            DTA5.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            DTA4.Dispose()
            DTA5.Dispose()
        End Try
    End Function

    ''Untuk Sales Journal
    Public Function SimpanItemSalesJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
       ByVal Frm As FrmTransaksiJournal_Sales, ByVal vDGVItemPayment As DataGridView, _
       ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, ByVal vJurnalNumber As String, _
       ByVal vPrimaryCust As Double, ByVal vTanggalJurnal As Date, ByVal vDGVItemAkunOther As DataGridView) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTABalance_SR As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

        Try
            'simpan data item
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColAccount.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryAkun.Index).Value), _
                    CDec(row.Cells(Frm.ColDebet.Index).Value), CDec(row.Cells(Frm.ColKredit.Index).Value), _
                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebetOriginal.Index).Value), CDec(row.Cells(Frm.ColCreditOriginal.Index).Value), _
                    CDec(row.Cells(Frm.ColExchangeRate.Index).Value), -1, -1)
                End If
            Next

            If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesInvoice_SalesInvoiceDP
            ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.SalesReturn
            End If

            For Each row In vDGVItemPayment.Rows
                If CDec(row.Cells(Frm.ColReceivement.Index).Value) <> 0 Then
                    If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                        'simpan data receivement
                        DTA_Payment.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.ColReceivement.Index).Value), TipePayment, _
                        row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value, row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value, _
                        row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, row.Cells(Frm.ColExhangeRateBalance.Index).Value, _
                        row.Cells(Frm.ColReceivementOriginal.Index).Value, row.Cells(Frm.ColPrimaryAkunBalance.Index).Value, _
                        row.Cells(Frm.ColInvoiceNo.Index).Value, "")

                        'update data balance dan received
                        DTABalance.UpdateQuery(row.Cells(Frm.ColBalanceSales.Index).Value, _
                        CDec(row.Cells(Frm.ColReceivement.Index).Value) + CDec(row.Cells(Frm.ColReceived.Index).Value), _
                        CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value) + CDec(row.Cells(Frm.ColReceivedOriginal.Index).Value), _
                        row.Cells(Frm.ColBalanceOriginal.Index).Value, _
                        row.Cells(Frm.ColPrimarySalesBalance.Index).Value, CDbl(row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value), _
                        row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value)

                        'simpan data kartu
                        If CDec(row.Cells(Frm.ColReceivement.Index).Value) > 0 Then
                            'untuk pembayaran sales invoice
                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                            vJurnalNumber, vPrimaryCust, 0, CDec(row.Cells(Frm.ColReceivement.Index).Value), _
                            row.Cells(Frm.ColInvoiceNo.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
                            CDbl(row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value), CDate(row.Cells(Frm.ColInvoiceDate.Index).Value), _
                            Math.Abs((CDate(row.Cells(Frm.ColInvoiceDate.Index).Value) - CDate(row.Cells(Frm.ColDueDateInvoice.Index).Value)).Days), _
                            Me.GetTipeTransaksiAsalSalesReceivementJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
                            0, CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value), CDec(row.Cells(Frm.ColExhangeRateBalance.Index).Value), _
                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
                        Else
                            'untuk pembayaran sales return
                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                            vJurnalNumber, vPrimaryCust, Math.Abs(CDec(row.Cells(Frm.ColReceivement.Index).Value)), 0, _
                            row.Cells(Frm.ColInvoiceNo.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, vTanggalJurnal, _
                            CDbl(row.Cells(Frm.ColPrimaryDocumenAcuan_Sales.Index).Value), CDate(row.Cells(Frm.ColInvoiceDate.Index).Value), _
                            Math.Abs((CDate(row.Cells(Frm.ColInvoiceDate.Index).Value) - CDate(row.Cells(Frm.ColDueDateInvoice.Index).Value)).Days), _
                            Me.GetTipeTransaksiAsalSalesReceivementJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Sales.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
                            Math.Abs(CDec(row.Cells(Frm.ColReceivementOriginal.Index).Value)), 0, CDec(row.Cells(Frm.ColExhangeRateBalance.Index).Value), _
                             CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
                        End If
                    ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next

            'simpan akun other
            For Each row In vDGVItemAkunOther.Rows
                With row
                    DTAOther.Insert(vPrimaryTransaksi, .Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Frm.ColAccountDiff_Value.Index).Value, _
                    .Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value, .Cells(Frm.ColAccountDiff_Remarks.Index).Value, _
                    .Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value, .Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Payment.Dispose()
            DTABalance.Dispose()
            DTABalance_SR.Dispose()
            ObjCD = Nothing
            ObjKartu = Nothing
            DTAOther.Dispose()
        End Try
    End Function

    Public Function HapusItem_SalesJurnal(ByVal vPrimaryTransaksi As Double, ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBL_Payment As New DseTransaksiJournal.tbltrans_jurnal_item_paymentDataTable
        Dim row_Payment As DseTransaksiJournal.tbltrans_jurnal_item_paymentRow
        Dim DTABalance As New DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        'Dim DTABalance_SR As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim Balance_Lama, Received_Lama, BalanceOriginal_Lama, ReceivedOriginal_Lama As Decimal
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

        Try
            'update data balance sales
            DTA_Payment.FillBy_PrimaryJurnal(TBL_Payment, vPrimaryTransaksi)

            For Each row_Payment In TBL_Payment.Rows
                If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                    If DTABalance.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Balance_Lama = DTABalance.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Balance_Lama = 0
                    End If

                    If DTABalance.ScalarQuery_GetReceived(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Received_Lama = DTABalance.ScalarQuery_GetReceived(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Received_Lama = 0
                    End If

                    If DTABalance.ScalarQuery_GetBalanceOriginalLama(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        BalanceOriginal_Lama = DTABalance.ScalarQuery_GetBalanceOriginalLama(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        BalanceOriginal_Lama = 0
                    End If

                    If DTABalance.ScalarQuery_GetReceivedOriginalLama(row_Payment.Primary_Document_Acuan, _
                   row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        ReceivedOriginal_Lama = DTABalance.ScalarQuery_GetReceivedOriginalLama(row_Payment.Primary_Document_Acuan, _
                   row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        ReceivedOriginal_Lama = 0
                    End If

                    DTABalance.UpdateQuery1(Balance_Lama + row_Payment.Payment, Received_Lama - _
                    row_Payment.Payment, BalanceOriginal_Lama + row_Payment.payment_original, ReceivedOriginal_Lama - row_Payment.payment_original, _
                    row_Payment.Primary_Document_Acuan, row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency)

                    'hapus data kartu
                    ObjKartu.Hapus(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement)
                ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                    MsgBox("Logic error, contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                End If
            Next

            'hapus data item payment + jurnal + other
            DTA_Payment.Delete(vPrimaryTransaksi)
            DTA.Delete(vPrimaryTransaksi)
            DTAOther.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Payment.Dispose()
            TBL_Payment.Dispose()
            ObjKartu = Nothing
            DTAOther.Dispose()
        End Try
    End Function

    ''Untuk Purchase Journal
    ''
    Public Function SimpanItemPurchaseJournal(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiJournal_Purchase, ByVal vDGVItemPayment As DataGridView, _
    ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal, ByVal vJurnalNumber As String, _
    ByVal vPrimarySupplier As Double, ByVal vTanggalJurnal As Date, ByVal vDGVItemAkunOther As DataGridView) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TipePayment As ClsTransaksiSalesBalance.EnumTipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
        Dim DTABalance As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DTABalance_SR As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

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
                    row.Cells(Frm.ColRemarks.Index).Value, CDec(row.Cells(Frm.ColDebetOriginal.Index).Value), CDec(row.Cells(Frm.ColCreditOriginal.Index).Value), _
                    CDec(row.Cells(Frm.ColExchangeRate.Index).Value), -1, -1)
                End If
            Next

            If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.AOP_AOPDP
            ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                TipePayment = ClsTransaksiSalesBalance.EnumTipePayment.Purchase_Return
            End If

            For Each row In vDGVItemPayment.Rows
                If CDec(row.Cells(Frm.ColPayment.Index).Value) <> 0 Then
                    If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                        'simpan data payment
                        DTA_Payment.Insert(vPrimaryTransaksi, CDec(row.Cells(Frm.ColPayment.Index).Value), TipePayment, _
                        row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value, row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value, _
                        row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, row.Cells(Frm.ColExchangeRateBalance.Index).Value, _
                        row.Cells(Frm.ColPaymentOriginal.Index).Value, row.Cells(Frm.ColPrimaryAkunBalance.Index).Value, _
                        row.Cells(Frm.ColAOPNumber.Index).Value, row.Cells(Frm.ColSupplierInvoiceNumber.Index).Value)

                        'update data balance dan received
                        DTABalance.UpdateQuery(row.Cells(Frm.ColBalancePurchase.Index).Value, _
                        CDec(row.Cells(Frm.ColPayment.Index).Value) + CDec(row.Cells(Frm.ColPaid.Index).Value), _
                        CDec(row.Cells(Frm.ColBalanceOriginal.Index).Value), _
                        CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value + CDec(row.Cells(Frm.ColPaidOriginal.Index).Value)), _
                        row.Cells(Frm.ColPrimaryPurchaseBalance.Index).Value, CDbl(row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value), _
                        row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value)

                        'simpan data kartu
                        If CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value) > 0 Then
                            'pembayaran aop, didebet
                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
                            vJurnalNumber, vPrimarySupplier, CDec(row.Cells(Frm.ColPayment.Index).Value), 0, _
                            row.Cells(Frm.ColAOPNumber.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalJurnal, CDbl(row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value), CDate(row.Cells(Frm.ColAOPDate.Index).Value), _
                            Math.Abs((CDate(row.Cells(Frm.ColDueDatePurchase.Index).Value) - CDate(row.Cells(Frm.ColAOPDate.Index).Value)).Days), _
                            Me.GetTipeTransaksiAsalPurchasePaymentJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
                            Math.Abs(CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value)), 0, CDec(row.Cells(Frm.ColExchangeRateBalance.Index).Value), _
                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
                        Else
                            'pembayaran pd dikredit
                            ObjKartu.Simpan(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
                            vJurnalNumber, vPrimarySupplier, 0, Math.Abs(CDec(row.Cells(Frm.ColPayment.Index).Value)), _
                            row.Cells(Frm.ColAOPNumber.Index).Value, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalJurnal, CDbl(row.Cells(Frm.ColPrimaryDocumentAcuan_Purchase.Index).Value), CDate(row.Cells(Frm.ColAOPDate.Index).Value), _
                            Math.Abs((CDate(row.Cells(Frm.ColDueDatePurchase.Index).Value) - CDate(row.Cells(Frm.ColAOPDate.Index).Value)).Days), _
                            Me.GetTipeTransaksiAsalPurchasePaymentJurnal(row.Cells(Frm.ColTipeDocumentAcuan_Purchase.Index).Value), row.Cells(Frm.ColPrimaryCurrencyBalance.Index).Value, _
                            0, Math.Abs(CDec(row.Cells(Frm.ColPaymentOriginal.Index).Value)), CDec(row.Cells(Frm.ColExchangeRateBalance.Index).Value), _
                            CDec(row.Cells(Frm.ColPrimaryAkunBalance.Index).Value))
                        End If
                    ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                        MsgBox("logic error please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next

            'simpan akun other
            For Each row In vDGVItemAkunOther.Rows
                With row
                    DTAOther.Insert(vPrimaryTransaksi, .Cells(Frm.ColAccountDiff_PrimaryAkun.Index).Value, .Cells(Frm.ColAccountDiff_Value.Index).Value, _
                    .Cells(Frm.ColAccountDiff_ValueOriginalCurrency.Index).Value, .Cells(Frm.ColAccountDiff_Remarks.Index).Value, _
                    .Cells(Frm.ColAccountDiff_PrimaryCurrency.Index).Value, .Cells(Frm.ColAccountDiff_ExchangeRate.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Payment.Dispose()
            DTABalance.Dispose()
            DTABalance_SR.Dispose()
            ObjCD = Nothing
            ObjKartu = Nothing
            DTAOther.Dispose()
        End Try
    End Function

    Public Function GetTipeTransaksiAsalPurchasePaymentJurnal(ByVal vTipeBalance As _
    ClsTransaksiSalesBalance.EnumTipeSalesBalance) As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
        Try
            If vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetTipeTransaksiAsalSalesReceivementJurnal(ByVal vTipeBalance As _
    ClsTransaksiSalesBalance.EnumTipeSalesBalance) As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
        Try
            If vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_DP Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function HapusItem_PurchaseJurnal(ByVal vPrimaryTransaksi As Double, ByVal vTipeJurnal As ClsTransaksiJournal.EnumTipeJournal) As Boolean
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim DTA_Payment As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBL_Payment As New DseTransaksiJournal.tbltrans_jurnal_item_paymentDataTable
        Dim row_Payment As DseTransaksiJournal.tbltrans_jurnal_item_paymentRow
        Dim DTABalance_Purchase As New DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        'Dim DTABalance_PR As New DseTransaskiPurchaseReturnBalanceTableAdapters.tbltrans_balance_purchasereturnTableAdapter
        Dim Balance_Lama, Received_Lama, Balance_Lama_Original, Received_Lama_Original As Decimal
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim DTAPurchase As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter
        Dim DTAOther As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_item_otherTableAdapter

        Try
            'update data balance sales
            DTA_Payment.FillBy_PrimaryJurnal(TBL_Payment, vPrimaryTransaksi)
            For Each row_Payment In TBL_Payment.Rows
                If vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                    If DTABalance_Purchase.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Balance_Lama = DTABalance_Purchase.ScalarQuery_GetBalance(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Balance_Lama = 0
                    End If
                    If DTABalance_Purchase.ScalarQuery_GetPaid(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Received_Lama = DTABalance_Purchase.ScalarQuery_GetPaid(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Received_Lama = 0
                    End If

                    If DTABalance_Purchase.ScalarQuery_GetBalanceOriginal2(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Balance_Lama_Original = DTABalance_Purchase.ScalarQuery_GetBalanceOriginal2(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Balance_Lama_Original = 0
                    End If
                    If DTABalance_Purchase.ScalarQuery_GetPaidOriginal(row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).HasValue = True Then
                        Received_Lama_Original = DTABalance_Purchase.ScalarQuery_GetPaidOriginal(row_Payment.Primary_Document_Acuan, _
                        row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency).Value
                    Else
                        Received_Lama_Original = 0
                    End If

                    DTABalance_Purchase.UpdateQuery1(Balance_Lama + row_Payment.Payment, Received_Lama - _
                    row_Payment.Payment, Balance_Lama_Original + row_Payment.payment_original, _
                    Received_Lama_Original - row_Payment.payment_original, row_Payment.Primary_Document_Acuan, _
                    row_Payment.Tipe_Document_Acuan, row_Payment.primary_currency)

                    ObjKartu.Hapus(vPrimaryTransaksi, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment)
                ElseIf vTipeJurnal = ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                    MsgBox("Logic error, contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                End If
            Next

            'hapus data item payment + jurnal
            DTA_Payment.Delete(vPrimaryTransaksi)
            DTA.Delete(vPrimaryTransaksi)
            DTAPurchase.Delete(vPrimaryTransaksi)
            DTAOther.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Payment.Dispose()
            TBL_Payment.Dispose()
            DTAPurchase.Dispose()
            ObjKartu = Nothing
            DTAOther.Dispose()
        End Try
    End Function
End Class
