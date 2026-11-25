Option Explicit On
Option Strict Off
Public Class FrmTransaksiCreditNote
    Dim BatasBalanceMaksimum As Double = 0
    Dim batasBalanceMaksimum2 As Double = 0
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
    Public TipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE
    Public IsBeginingBalance As Integer = 0
    Dim TipeNoTrans As String
    Dim objSetting As New Core_Module.ClsTransaksiSetting
    Dim baru As Boolean = True
    Dim objCd As New Core_Module.ClsCekData
    Dim objTransCN As New ClsTransaksiCreditNote
    Dim objCust As New Core_Module.clsMasterCustomer
    Dim objSupp As New Core_Module.clsMasterSupplier
    Dim objCurr As New Core_Module.clsMasterCurrency
    Dim objAkun As New ClsMasterAkun
    Dim objSalesInv As New ClsTransaksiSalesInvoice
    Dim objAop As New ClsTransaksiAOP
    Dim TblArAccount As New DseTransaksiCreditNote.tblisi_cbo_ar_akunDataTable
    Dim TblArAccount2 As New DseTransaksiCreditNote.tblmaster_akunDataTable
    Dim TblArAccountByPrimary As New DseTransaksiCreditNote.tblisi_cbo_ar_akun2DataTable
    Dim TBlCurrency As New DseTransaksiCreditNote.tblmaster_currencyDataTable
    Dim TBlCurrencyReff As New DseTransaksiCreditNote.tblmaster_currencyDataTable
    Dim RowArAccount As DseTransaksiCreditNote.tblisi_cbo_ar_akunRow
    Dim PrimaryCurrReffAcc As Double = -1
    Dim CurrCodeReffAcc As String = ""
    Dim DeffPrecX As Decimal = 4
    Dim ObjClosing As New ClsTransaksiClosing
    Public ObjJurnal As New ClsTransaksiJournal
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim SedangLeave As Integer = 0
    Dim PrimaryBaseCurr As Integer = Core_Module.PrimaryBaseCurrency
    Dim DefPrec_BaseCurr As Integer = 0
    Dim DefaultPrecision As Integer = 0
    Dim Active As Boolean = False
    Dim Lanjut As Boolean = True
    Dim EraseRowIdx As Integer = -1
    Dim Buffer As String = ""
    Dim DefPrecEx As Integer = 0
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJob As New Core_Module.ClsMasterJob
    Dim PrimaryAccAr_Lama As String = ""

    Private Sub CekExpense_Account()
        Dim Dummy As String = ""

        Try
            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                    'expense
                    Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Customer)
                Else
                    'income
                    Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Supplier)
                End If
            Else
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                    'expense
                    Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Supplier)
                Else
                    'income
                    Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Customer)
                End If
            End If

            If Dummy = "" Then
                Me.TxtPrimaryExpenseAcc.Text = "-1"
            Else
                Me.TxtPrimaryExpenseAcc.Text = Dummy
            End If
            Me.TxtExpenseAccDesc.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryExpenseAcc.Text)
            Me.TxtExpenseAcc.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryExpenseAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        
        Try
            Me.KosongkanField()
            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                TipeNoTrans = "CN"
            Else
                TipeNoTrans = "DN"
            End If
            Me.baru = True
            Me.CekExpense_Account()
            Me.dgvItem.Rows.Add()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiDgv_AddItemDefault()
            Me.AktifkanDetail()
            Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, Me.TipeTransaksi)
            Me.DtpDate.Focus()
            Me.txtNoTransaksi.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgv_AddItemDefault()
        Dim R As New DataGridViewRow
        Try
            If Me.TblArAccount.Rows.Count > 0 And Me.TxtPrimaryArAccount.Text <> -1 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                R = Me.dgvItem.Rows(0)
                With R
                    .Cells(Me.ColPrimaryAkun.Index).Value = Me.TxtPrimaryArAccount.Text
                    .Cells(Me.ColAccount.Index).Value = objAkun.GetAccountCode(Me.TxtPrimaryArAccount.Text)
                    .Cells(Me.ColDescAccount.Index).Value = objAkun.GetDescription(.Cells(Me.ColAccount.Index).Value)
                    .Cells(Me.ColDebet.Index).Value = 0
                    .Cells(Me.ColKredit.Index).Value = 0
                    .Cells(Me.ColDebet_OriginalCurr.Index).Value = 0
                    .Cells(Me.ColCredit_OriginalCurrency.Index).Value = 0
                    .Cells(Me.ColRemarks.Index).Value = ""
                    .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency.Text
                    .Cells(Me.ColOriginalCurrency.Index).Value = Me.cboCurrency.Text
                    .Cells(Me.ColExchangeRate.Index).Value = Me.TxtExchangeRate.Text
                    .Cells(Me.ColDept.Index).Value = ""
                    .Cells(Me.ColPrimaryDept.Index).Value = -1
                    .Cells(Me.ColPrimaryJob.Index).Value = -1
                    .Cells(Me.ColJob.Index).Value = ""
                    .Cells(Me.ColDefPrec.Index).Value = Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                    Me.DefaultPrecision = .Cells(Me.ColDefPrec.Index).Value
                    Me.DefPrec_BaseCurr = Me.objCurr.GetDefaultPrecission(Me.PrimaryBaseCurr)
                    Me.TxtBalance.Text = Me.objCd.DigitGrouping(0, Me.DefaultPrecision)
                End With
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If

            Me.UpdateBalance()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetMulticurrency()
        Dim rowDGV As DataGridViewRow
        Dim Multicurrency As Boolean = False

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index = Me.dgvItem.RowCount - 1 Then
                    If Me.dgvItem.AllowUserToAddRows = True Then
                        'tidak dicek
                    Else
                        'dicek
                        If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurr Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurr Then
                        Multicurrency = True
                        Exit For
                    End If
                End If
            Next

            If Multicurrency = True Then
                Me.ColExchangeRate.Visible = True
                Me.ColDebet_OriginalCurr.Visible = True
                Me.ColCredit_OriginalCurrency.Visible = True
                Me.ColOriginalCurrency.Visible = True
            Else
                Me.ColOriginalCurrency.Visible = False
                Me.ColExchangeRate.Visible = False
                Me.ColDebet_OriginalCurr.Visible = False
                Me.ColCredit_OriginalCurrency.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateBalance()
        Dim dR As New DataGridViewRow
        Try
            If Me.Active = True Then
                dR = Me.dgvItem.Rows(0)
                If TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE Then
                    dR.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), Me.DefPrec_BaseCurr, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                    dR.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                    dR.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.objCd.DigitGrouping(Me.TxtBalance.Text, dR.Cells(Me.ColDefPrec.Index).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                    dR.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.objCd.DigitGrouping(0, dR.Cells(Me.ColDefPrec.Index).Value)
                Else
                    dR.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                    dR.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), Me.DefPrec_BaseCurr, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                    dR.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.objCd.DigitGrouping(0, dR.Cells(Me.ColDefPrec.Index).Value)
                    dR.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.objCd.DigitGrouping(Me.TxtBalance.Text, dR.Cells(Me.ColDefPrec.Index).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                End If

                dR.Cells(Me.ColExchangeRate.Index).Value = Me.TxtExchangeRate.Text
                dR.Cells(Me.ColOriginalCurrency.Index).Value = Me.objCurr.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                'Me.HitungOriginalValues(0)
                'Me.HitungUnbalance()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOriginalValues(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.objCurr.GetDefaultPrecission(Me.PrimaryBaseCurr)
            DefPrecOriginal = Me.objCurr.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = objCd.DigitGrouping(0, DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = objCd.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = objCd.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = objCd.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HitungUnbalance(ByVal vdebet As Decimal, ByVal vkredit As Decimal) As Decimal
        Dim Unbalance As Decimal = 0
        Try
            Unbalance = Math.Round(vdebet - vkredit, Me.DefaultPrecision, MidpointRounding.AwayFromZero)
            Return Unbalance
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HitungUnbalance()
        Dim row As DataGridViewRow
        Try
            ''hitung debet dan kredit
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColAccount.Index).Value) = False Then
                    TotalDebet += CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value)
                    TotalKredit += CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value)
                End If
            Next
            Me.TxtUnbalance.Text = CStr(ObjCD.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note

        Try
            Me.dgvItem.AllowUserToAddRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.Active = True
            objCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            objCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            objCd.KendalikanControl(Me.gboDetail.Controls, False, False, True, True, False, False, False, False)
            objCd.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            'Me.BackColorReset()
            Me.FindExpenseAccount.Enabled = True
            Me.LinkAddress.Enabled = True
            'Me.LblExchangeRate.Enabled = True
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            'Me.txtNoTransaksi.ReadOnly = True
            'Me.txtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
            Me.cboPrimaryReffAcc.Enabled = False

            If baru = True Then
                Me.TxtBalance.Text = objCd.DigitGrouping(TxtBalance.Text, objCurr.GetDefaultPrecission(TxtPrimaryCurrency.Text))
                objCd.Number_Click(TxtBalance)
            End If

            If baru = False Then
                Me.CboCustomerSuuplier.Enabled = False
            End If

            Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
            If Not String.IsNullOrEmpty(Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)) Then
                CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
            Else
                CurrCodeReffAcc = Me.cboCurrency.Text
            End If

            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, True)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimary.BackColor, True)
            Me.objCd.SetBackColor(Me.groupDetail.Controls, Me.txtPrimary.BackColor, True)

            If Me.cboCurrency.Text = "IDR" Then
                Me.TxtExchangeRate.Text = "1"
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.TxtExchangeRate.ReadOnly = True
                'Me.TxtExchangeRate.ReadOnly = True
                'Me.TxtExchangeRate.TabStop = False

                If CurrCodeReffAcc = Me.cboCurrency.Text Then
                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Text = Me.TxtBalance.Text
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    Me.cboCurrencyReff.Enabled = True
                    Me.txtBalanceReff.Text = 0
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                End If
            Else
                If CurrCodeReffAcc = Me.cboCurrency.Text Then
                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Text = Me.TxtBalance.Text
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    Me.cboCurrencyReff.Enabled = True
                    Me.txtBalanceReff.Text = 0
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                End If
                Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                'Me.TxtExchangeRate.ReadOnly = False
                'Me.TxtExchangeRate.TabStop = True
                Me.TxtExchangeRate.Enabled = True
                Me.TxtExchangeRate.Clear()
            End If

            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If
            If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                Me.txtNoTransaksi.ReadOnly = True
                Me.txtNoTransaksi.BackColor = Me.txtReadOnly.BackColor
                Me.txtNoTransaksi.TabStop = False
                Me.DtpDate.Focus()
            Else
                Me.txtNoTransaksi.ReadOnly = False
                Me.txtNoTransaksi.BackColor = Me.txtPrimary.BackColor
                Me.txtNoTransaksi.TabStop = True
                'Me.txtNoTransaksi.Focus()
            End If
            Me.txtNoJurnal.Focus()
            Me.txtNama.BackColor = Me.txtReadOnly.BackColor
            Me.TxtExpenseAccDesc.BackColor = Me.txtReadOnly.BackColor
            Me.TxtUnbalance.BackColor = Me.txtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.txtReadOnly.ReadOnly = True
            Me.objCd.SetBackColor(Me.gboDetail.Controls, Me.txtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.Active = False

            objCd.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            objCd.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            objCd.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            objCd.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            'Me.LblExchangeRate.Enabled = False
            Me.FindExpenseAccount.Enabled = False
            Me.LinkAddress.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.TxtExchangeRate.ReadOnly = True
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtReadOnly.BackColor, True)
            Me.objCd.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtReadOnly.BackColor, True)
            Me.objCd.SetBackColor(Me.groupDetail.Controls, Me.txtReadOnly.BackColor, True)

            'Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.EraseRowIdx = -1
            Me.dgvItem.Rows.Clear()
            Me.CboCustomerSuuplier.SelectedIndex = 0
            Me.CboAddress.SelectedIndex = -1
            Me.CboArAccount.SelectedIndex = -1
            Me.cboPrimaryReffAcc.SelectedIndex = -1
            Me.cboCurrency.SelectedIndex = -1
            Me.cboCurrencyReff.SelectedIndex = -1

            Me.txtPrimary.Text = "-1"
            Me.TxtPrimaryArAccount.Text = "-1"
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.TxtPrimaryExpenseAcc.Text = "-1"
            Me.TxtPrimaryReferedInvoice.Text = "-1"
            Me.txtPrimaryCustSupp.Text = "-1"
            Me.txtPrimaryJournal.Text = "-1"
            Me.DtpDate.Value = Today
            Me.DtpDueDate.Value = Today
            Me.dtpTransaksi.Value = Today
            Me.txtNoTransaksi.Text = ""
            Me.txtCode.Clear()
            Me.txtNama.Clear()
            Me.TxtContactPerson.Clear()
            Me.TxtCustomerDebitNote.Clear()
            Me.TxtReferedInvoice.Clear()
            Me.TxtExpenseAcc.Text = ""
            Me.TxtExpenseAccDesc.Text = ""
            Me.TxtBalance.Clear()
            Me.TxtExchangeRate.Text = ""
            Me.txtRemarks.Clear()
            Me.CurrCodeReffAcc = ""
            Me.TxtUnbalance.Clear()
            Me.PrimaryCurrReffAcc = -1
            Me.txtDecJurnal.Clear()
            Me.txtNoJurnal.Clear()

            Me.baru = True
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColRemarks.Index).Value) = False Then
                    If row.Cells(Me.ColRemarks.Index).Value.ToString.Contains("'") = True Or row.Cells(Me.ColRemarks.Index).Value.ToString.Contains("*") = True Or row.Cells(Me.ColRemarks.Index).Value.ToString.Contains("%") = True Then
                        MsgBox("" + "Data item remarks" _
                            + " can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Error")
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim KoleksiNamaData2 As New Collection
        Dim KoleksiTextBox2 As New Collection
        Dim objAkun As New ClsMasterAkun
        'Dim DTAInvoice As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_salesinvoiceTableAdapter
        'Dim DTAAOP As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_aopTableAdapter
        Dim DTAInvoice As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTAAOP As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note

        Try
            'cek data yang harus diisi
            Me.txtNoTransaksi.Text = Me.txtNoTransaksi.Text.Trim
            Me.TxtCustomerDebitNote.Text = Me.TxtCustomerDebitNote.Text.Trim
            Me.txtDecJurnal.Text = Me.txtDecJurnal.Text.Trim
            Me.TxtContactPerson.Text = Me.TxtContactPerson.Text.Trim
            KoleksiNamaData.Add("No Journal")
            KoleksiTextBox.Add(Me.txtNoJurnal)
            KoleksiNamaData.Add("Document no")
            KoleksiTextBox.Add(Me.txtNoTransaksi)


            KoleksiNamaData.Add("Customer or Supplier code")
            KoleksiTextBox.Add(Me.txtCode)
            KoleksiNamaData.Add("Balance")
            KoleksiTextBox.Add(Me.TxtBalance)
            KoleksiNamaData.Add("Exchange rate")
            KoleksiTextBox.Add(Me.TxtExchangeRate)
            KoleksiNamaData.Add("Reason for claim")
            KoleksiTextBox.Add(Me.txtRemarks)
            If TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                KoleksiNamaData.Add("AR Account")
                KoleksiTextBox.Add(Me.CboArAccount)
            ElseIf TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                KoleksiNamaData.Add("AP Account")
                KoleksiTextBox.Add(Me.CboArAccount)
            End If
            KoleksiNamaData.Add("Currency")
            KoleksiTextBox.Add(Me.cboCurrency)
            If objCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '
            KoleksiNamaData2.Add("No Journal")
            KoleksiTextBox2.Add(Me.txtNoJurnal)
            KoleksiNamaData2.Add("Description Journal")
            KoleksiTextBox2.Add(Me.txtDecJurnal)
            KoleksiNamaData2.Add("Reason for claim")
            KoleksiTextBox2.Add(Me.txtRemarks)
            KoleksiNamaData2.Add("Referred Document No")
            KoleksiTextBox2.Add(Me.TxtCustomerDebitNote)
            If Me.CekDataKosong2(KoleksiTextBox2, KoleksiNamaData2) = False Then
                Return False
            End If
            KoleksiNamaData2.Clear()
            KoleksiTextBox2.Clear()
            'KoleksiNamaData.Add("Currency Reff")
            'KoleksiTextBox.Add(Me.cboCurrencyReff)
            'If objCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
            '    Return False
            'End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam nomor transaksi atau kosong
            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If
            If Me.baru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    If Me.objTransCN.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                        Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value.Date, Me.TipeTransaksi)
                    End If
                Else
                    If Me.objTransCN.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                'ditransaksinya
                If Me.objTransCN.ApakahNoJurnalAda(Me.txtNoJurnal.Text) = True Then
                    MsgBox("Duplication in no journal: " + Me.txtNoJurnal.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtNoJurnal.Focus()
                    Return False
                End If

                'ditbljurnalnya
                If Me.objTransCN.ApakahSudahAdaNoJurnal(Me.txtNoJurnal.Text) = True Then
                    MsgBox("Duplication in no journal: " + Me.txtNoJurnal.Text + ", please change your number", _
                                        MsgBoxStyle.Information, "Information")
                    Me.txtNoJurnal.Focus()
                    Return False
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    If Me.objTransCN.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                        If Me.objTransCN.GetPrimary(Me.txtNoTransaksi.Text) <> Me.txtPrimary.Text Then
                            Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value.Date, Me.TipeTransaksi)
                        End If
                    End If
                Else
                    If Me.objTransCN.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                        If Me.objTransCN.GetPrimary(Me.txtNoTransaksi.Text) <> Me.txtPrimary.Text Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If

                Dim TipeCodeLama As String = Me.objTransCN.GetNoTrans(Me.txtPrimary.Text)
                If TipeCodeLama <> Me.txtNoTransaksi.Text Then
                    If Me.objTransCN.ApakahSudahAda(Me.txtNoTransaksi.Text) = True Then
                        MsgBox("Duplication in Document No: " + Me.txtNoTransaksi.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If


                Dim NoJurLama As String = Me.objTransCN.GetNoJurnal(Me.txtPrimary.Text)
                'di transaksisnya
                If NoJurLama <> Me.txtNoJurnal.Text Then
                    If Me.objTransCN.ApakahNoJurnalAda(Me.txtNoJurnal.Text) = True Then
                        MsgBox("Duplication in no journal: " + Me.txtNoJurnal.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtNoJurnal.Focus()
                        Return False
                    End If
                    'di jurnalnya
                    If Me.objTransCN.ApakahSudahAdaNoJurnal(Me.txtNoJurnal.Text) = True Then
                        MsgBox("Duplication in no journal: " + Me.txtNoJurnal.Text + ", please change your number", _
                                            MsgBoxStyle.Information, "Information")
                        Me.txtNoJurnal.Focus()
                        Return False
                    End If
                End If





                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objTransCN.GetTimeStamp(Me.txtPrimary.Text)
                If Me.objTransCN.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objTransCN.CariDetail(Me.objTransCN.Primary, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Not IsNumeric(Me.TxtBalance.Text) OrElse CDec(Me.TxtBalance.Text) <= 0 Then
                MsgBox("Please specify a positive number for amount data.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If


            'If Not IsNumeric(Me.txtBalanceReff.Text) OrElse CDec(Me.txtBalanceReff.Text) <= 0 Then
            '    MsgBox("Please specify a positive number for ref. amount data.", MsgBoxStyle.Information, "Information")
            '    Me.txtBalanceReff.Focus()
            '    Return False
            'End If

            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count < 2 Then
                MsgBox("Minimum item = 2, please input item again", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Focus()
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If
            Me.dgvItem.AllowUserToAddRows = True

            If Fix(Me.TxtBalance.Text).ToString.Length > 15 Then
                MsgBox("Maximum value for AP/AP amount is 999,999,999,999.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If

            If Fix(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text)).ToString.Length > 15 Then
                MsgBox("Maximum value for AP/AP amount in IDR is 999,999,999,999.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
                Return False
            End If

            If Fix(Me.txtBalanceReff.Text).ToString.Length > 15 Then
                MsgBox("Maximum value for ref. amount is 999,999,999,999.", MsgBoxStyle.Information, "Information")
                Me.txtBalanceReff.Focus()
                Return False
            End If

            If Not String.IsNullOrEmpty(Me.TxtPrimaryReferedInvoice.Text) AndAlso Me.TxtPrimaryReferedInvoice.Text > -1 Then
                If (Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE AndAlso _
                   Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR) OrElse _
                   (Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE AndAlso _
                   Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP) Then
                    Dim decTotalInvoice As Decimal = 0
                    Dim decTotalCDNote As Decimal = 0
                    Dim decTotalCDNoteB4 As Decimal = 0

                    If DTA.ScalarQuery_GetAmountPerTransaksi(Me.txtPrimary.Text).HasValue AndAlso _
                       DTA.ScalarQuery_GetAmountPerTransaksi(Me.txtPrimary.Text).Value > 0 Then
                        decTotalCDNoteB4 = DTA.ScalarQuery_GetAmountPerTransaksi(Me.txtPrimary.Text).Value
                    End If

                    If DTA.ScalarQuery_GetAmountByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text, Me.TipeKartu, Me.txtPrimary.Text).HasValue AndAlso _
                       DTA.ScalarQuery_GetAmountByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text, Me.TipeKartu, Me.txtPrimary.Text).Value > 0 Then
                        decTotalCDNote = DTA.ScalarQuery_GetAmountByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text, Me.TipeKartu, Me.txtPrimary.Text).Value
                    End If

                    If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                        If DTAInvoice.ScalarQuery_GetBalanceByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text).HasValue AndAlso _
                            DTAInvoice.ScalarQuery_GetBalanceByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text).Value > 0 Then
                            decTotalInvoice = DTAInvoice.ScalarQuery_GetBalanceByPrimaryInvoice(Me.TxtPrimaryReferedInvoice.Text).Value
                        End If
                    ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                        If DTAAOP.ScalarQuery_GetBalanceByPrimaryAOP(Me.TxtPrimaryReferedInvoice.Text).HasValue AndAlso _
                                  DTAAOP.ScalarQuery_GetBalanceByPrimaryAOP(Me.TxtPrimaryReferedInvoice.Text).Value > 0 Then
                            decTotalInvoice = DTAAOP.ScalarQuery_GetBalanceByPrimaryAOP(Me.TxtPrimaryReferedInvoice.Text).Value
                        End If
                    End If
                    'decTotalCDNote += CDec(Me.TxtBalance.Text)

                    If (decTotalInvoice) - (decTotalCDNote) < CDec(Me.TxtBalance.Text) Then
                        If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                            MsgBox("Your balance value of debit note for AOP number (AP balance - total debet note): " & Me.TxtReferedInvoice.Text & _
                            " is " & objCd.DigitGrouping((decTotalInvoice) - (decTotalCDNote), Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)) & ". AP balance for this AOP number will be minus. " _
                              , MsgBoxStyle.Information, "Information")
                            Me.TxtReferedInvoice.Focus()
                            Return False
                        ElseIf Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                            MsgBox("Your balance value of credit note for invoice number (AR balance - total credit note): " & Me.TxtReferedInvoice.Text & _
                                   " is " & objCd.DigitGrouping((decTotalInvoice) - (decTotalCDNote), Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)) & ". AR balance for this invoice number will be minus. " _
                                   , MsgBoxStyle.Information, "Information")
                            Me.TxtReferedInvoice.Focus()
                            Return False
                        End If
                    End If
                End If
            End If

            'cek perubahan nomor transaksi
            If Me.baru = False Then
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    If Me.ObjNumbering.GetResetType(TipeNumbering) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                        If Me.objTransCN.Tgl.Year = Me.DtpDate.Value.Year Then
                            Me.txtNoTransaksi.Text = Me.objTransCN.NoTrans
                        Else
                            If MsgBox(" You just changed transaction date, this will change your transaction " _
                            + " number and there's a possibility your transaction number cannot be changed back, " _
                                + "continue and accept this change?", _
                                MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.objTransCN.Tgl
                                Me.txtNoTransaksi.Text = Me.objTransCN.NoTrans
                                Return False
                            End If
                        End If
                    Else
                        If Me.objTransCN.Tgl.Month = Me.DtpDate.Value.Month And _
                        Me.objTransCN.Tgl.Year = Me.DtpDate.Value.Year Then
                            Me.txtNoTransaksi.Text = Me.objTransCN.NoTrans
                        Else
                            If MsgBox(" You just changed transaction date, this will change your transaction " _
                            + " number and there's a possibility your transaction number cannot be changed back, " _
                                + "continue and accept this change?", _
                                MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.objTransCN.Tgl
                                Me.txtNoTransaksi.Text = Me.objTransCN.NoTrans
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColAccount.Index).Value) = False Then
                    If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
                                    Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
                        MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            If Me.TxtUnbalance.Text <> 0 Then
                MsgBox("This is an unbalance account. Please add numbers in debet or credit column so that a balanced calculation", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dim ObjJournal As New ClsTransaksiJournal
            If Me.TxtPrimaryCurrency.Text <> Core_Module.PrimaryBaseCurrency Then
                If ObjJournal.ApakahSudahRevaluasi(Me.DtpDate.Value.Date) = True Then
                    MsgBox("There's end of month journal for period " + Format(Me.DtpDate.Value.Date, "MMM yyyy") + _
                    ", you can't create/edit credit note/debit note with non-IDR currency.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim TipeBalance As Integer = 0
            Dim Acuan As String = ""
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
                Else
                    TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
                End If
            Else
                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                Else
                    TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                End If
            End If

            If Me.baru = False Then
                Dim DTAXSales As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
                Dim DTAXPurchase As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
                DTAXSales.Connection.ConnectionString = DTA.Connection.ConnectionString
                DTAXPurchase.Connection.ConnectionString = DTA.Connection.ConnectionString

                If DTAXSales.ScalarQuery_getReceived(Me.txtPrimary.Text, TipeBalance).HasValue AndAlso CDec(DTAXSales.ScalarQuery_getReceived(Me.txtPrimary.Text, TipeBalance).Value) <> 0 Then
                    'cari sales receivement journal yang merujuk pada transaksi ini
                    Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, Me.txtPrimary.Text)
                    MsgBox("You cannot edit this data because transaction journal already received on sales receivement journal: " + Acuan, MsgBoxStyle.Information, "Information")
                    Me.objTransCN.CariDetail(Me.txtPrimary.Text, Me)
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Return False
                End If

                If DTAXPurchase.ScalarQuery_GetPaid(Me.txtPrimary.Text, TipeBalance).HasValue AndAlso CDec(DTAXPurchase.ScalarQuery_GetPaid(Me.txtPrimary.Text, TipeBalance).Value) <> 0 Then
                    Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, Me.txtPrimary.Text)
                    MsgBox("You cannot edit this data because transaction journal already payed on purchase payment journal: " + Acuan, MsgBoxStyle.Information, "Information")
                    Me.objTransCN.CariDetail(Me.txtPrimary.Text, Me)
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Return False
                End If
            End If

            If Me.ApakahDataItemBenar = False Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTAInvoice.Dispose()
            DTAAOP.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function CekDataKosong2(ByVal KoleksiTextBox As Collection, ByVal KoleksiNamaData As Collection) As Boolean
        Dim i As Int16
        Dim data As String
        Dim Txt As Windows.Forms.TextBox

        Try
            Do While i < KoleksiTextBox.Count
                'trim kanan kiri
                Txt = TryCast(KoleksiTextBox.Item(i + 1), Windows.Forms.TextBox)
                If IsNothing(Txt) = False Then
                    Txt.Text = Txt.Text.Trim
                End If
                'salin data
                data = KoleksiTextBox.Item(i + 1).text
                'data = HapusSpasi(data)

                ''cek data
                'If data.Length = 0 Then
                '    MsgBox(Convert.ToString(KoleksiNamaData.Item(i + 1)) + " cannot be an empty string.", MsgBoxStyle.Information, "Error")
                '    KoleksiTextBox.Item(i + 1).focus()
                '    Return False
                'End If

                'cek berkaitan dengan karakter yang tidak diingikan --> ', *, %
                If data.Contains("'") = True Or data.Contains("*") = True Or data.Contains("%") = True Then
                    MsgBox("" + KoleksiNamaData.Item(i + 1) _
                        + " data can not contain the following characters: ' , % , or * .", MsgBoxStyle.Information, "Error")
                    KoleksiTextBox.Item(i + 1).focus()
                    Return False
                End If
                i += 1
            Loop

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.objTransCN
                .Primary = Me.txtPrimary.Text
                .PrimaryAkunAr = Me.TxtPrimaryArAccount.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .PrimaryExpenseAcc = Me.TxtPrimaryExpenseAcc.Text
                .PrimaryInvoice = Me.TxtPrimaryReferedInvoice.Text
                .PrimaryCustomer = Me.txtPrimaryCustSupp.Text
                .PrimaryAkunAr = Me.TxtPrimaryArAccount.Text
                .PrimaryJournal = Me.txtPrimaryJournal.Text
                .NoTrans = Me.txtNoTransaksi.Text
                .Tgl = Me.DtpDate.Value
                .TglTransaksi = Me.DtpDate.Value 'sengaja disamaain karna utk smess beda dgn IGI
                .CustomerDebitNote = Me.TxtCustomerDebitNote.Text
                .DueDate = Me.DtpDueDate.Value
                .Reason = Me.txtRemarks.Text
                .ContactPerson = Me.TxtContactPerson.Text
                .IdentifierAddCust = Me.txtCustomerAddIdentifier.Text
                .Amount = Me.TxtBalance.Text
                .AmountReff = Me.txtBalanceReff.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .TipeTransaksi = TipeTransaksi
                '- jika customer
                If Me.CboCustomerSuuplier.SelectedIndex = 0 Then
                    .TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
                Else
                    '- jika supplier
                    .TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
                End If
                .NoJurnal = Me.txtNoJurnal.Text
                .DescJurnal = Me.txtDecJurnal.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim Dummy As String = ""

        Try
            Me.Lanjut = False
            With Me.objTransCN
                Me.TipeTransaksi = .TipeTransaksi
                Me.TipeKartu = .TipeKartu
                Me.txtPrimary.Text = .Primary

                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeNoTrans = "CN"
                Else
                    TipeNoTrans = "DN"
                End If

                '- jika customer
                If TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                    Me.CboCustomerSuuplier.SelectedIndex = 0
                    Me.DtpDate.Value = .Tgl
                    Me.txtNoTransaksi.Text = .NoTrans
                    Me.txtPrimaryCustSupp.Text = .PrimaryCustomer
                    Me.txtCustomerAddIdentifier.Text = .IdentifierAddCust
                    Me.TxtPrimaryArAccount.Text = .PrimaryAkunAr
                    Me.txtSupplierInvoiceNo.Text = ""
                    Me.TxtPrimaryCurrency.Text = .PrimaryCurrency
                    Me.TxtPrimaryReferedInvoice.Text = .PrimaryInvoice
                    Me.TxtPrimaryExpenseAcc.Text = .PrimaryExpenseAcc
                    Me.txtCode.Text = objCust.GetCustomerCode(Me.txtPrimaryCustSupp.Text)
                    Me.txtNama.Text = objCust.GetCustName(Me.txtPrimaryCustSupp.Text)

                    Me.objCust.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                    Dummy = Me.objCust.GetKodeAddress(.IdentifierAddCust)
                    While Me.CboAddress.Text <> Dummy
                        If Me.CboAddress.SelectedIndex + 1 <= Me.CboAddress.Items.Count - 1 Then
                            Me.CboAddress.SelectedIndex += 1
                        End If
                    End While
                    Me.TxtContactPerson.Text = .ContactPerson

                    Me.TxtCustomerDebitNote.Text = .CustomerDebitNote
                    Me.CboArAccount.Text = Me.objAkun.GetAccountCode(.PrimaryAkunAr)
                    Me.cboCurrency.Text = Me.objCurr.GetCurrencyCode(.PrimaryCurrency)
                    Me.TxtReferedInvoice.Text = Me.objSalesInv.GetInvoiceNumber(.PrimaryInvoice)
                    Me.TxtExpenseAccDesc.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryExpenseAcc.Text)
                    Me.TxtExpenseAcc.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryExpenseAcc.Text)
                    Me.TxtBalance.Text = Me.objCd.DigitGrouping(.Amount, objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                    Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                    Me.CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                    Me.TxtExchangeRate.Text = Me.objCd.DigitGrouping(.ExchangeRate, DeffPrecX)
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(.AmountReff, objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc))
                    Me.batasBalanceMaksimum2 = Me.TxtBalance.Text
                    Me.txtPrimaryJournal.Text = .PrimaryJournal
                    Me.DtpDueDate.Value = .DueDate
                    Me.dtpTransaksi.Value = .Tgl 'sengaja disamaain dengan tgl/date document
                    Me.txtRemarks.Text = .Reason
                    Me.txtNoJurnal.Text = .NoJurnal
                    Me.txtDecJurnal.Text = .DescJurnal
                    '  Me.cboPrimaryReffAcc.Text = Me.CurrCodeReffAcc
                ElseIf TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                    '- jika supplier
                    Me.CboCustomerSuuplier.SelectedIndex = 1
                    Me.DtpDate.Value = .Tgl
                    Me.txtNoTransaksi.Text = .NoTrans
                    Me.txtPrimaryCustSupp.Text = .PrimaryCustomer
                    Me.TxtPrimaryArAccount.Text = .PrimaryAkunAr
                    Me.TxtPrimaryCurrency.Text = .PrimaryCurrency
                    Me.TxtPrimaryReferedInvoice.Text = .PrimaryInvoice
                    If .PrimaryInvoice <> -1 Then
                        Me.txtSupplierInvoiceNo.Text = Me.objTransCN.GetSuppInvoiceNoByPrimaryAOP(.PrimaryInvoice)
                    Else
                        Me.txtSupplierInvoiceNo.Text = "-"
                    End If
                    Me.TxtPrimaryExpenseAcc.Text = .PrimaryExpenseAcc
                    Me.txtCode.Text = objSupp.GetSupplierCode(Me.txtPrimaryCustSupp.Text)
                    Me.txtNama.Text = objSupp.GetSupplierName(Me.txtPrimaryCustSupp.Text)
                    Me.TxtCustomerDebitNote.Text = .CustomerDebitNote
                    Me.txtCustomerAddIdentifier.Text = .IdentifierAddCust
                    Me.CboArAccount.Text = Me.objAkun.GetAccountCode(.PrimaryAkunAr)
                    Me.cboCurrency.Text = Me.objCurr.GetCurrencyCode(.PrimaryCurrency)
                    'Me.cboCurrency.Text = Me.objTransCN.GetsCurrCode(.PrimaryCurrency)
                    Me.objSupp.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                    Dummy = Me.objSupp.GetKodeAddress(.IdentifierAddCust)
                    While Me.CboAddress.Text <> Dummy
                        If Me.CboAddress.SelectedIndex + 1 <= Me.CboAddress.Items.Count - 1 Then
                            Me.CboAddress.SelectedIndex += 1
                        End If
                    End While
                    Me.TxtContactPerson.Text = .ContactPerson

                    Me.TxtReferedInvoice.Text = Me.objAop.GetNoTransaksi(.PrimaryInvoice)
                    Me.TxtExpenseAccDesc.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryExpenseAcc.Text)
                    Me.TxtExpenseAcc.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryExpenseAcc.Text)
                    Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                    Me.CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                    Me.TxtBalance.Text = Me.objCd.DigitGrouping(.Amount, objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                    Me.TxtExchangeRate.Text = Me.objCd.DigitGrouping(.ExchangeRate, DeffPrecX)
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(.AmountReff, objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc))
                    batasBalanceMaksimum2 = Me.TxtBalance.Text
                    Me.txtPrimaryJournal.Text = .PrimaryJournal
                    Me.DtpDueDate.Value = .DueDate
                    Me.dtpTransaksi.Value = .Tgl
                    Me.txtRemarks.Text = .Reason
                    Me.txtNoJurnal.Text = .NoJurnal
                    Me.txtDecJurnal.Text = .DescJurnal
                    '  Me.cboPrimaryReffAcc.Text = Me.CurrCodeReffAcc
                End If
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.Lanjut = True
            Me.HitungUnbalance()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCreditNote_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        '- cek nilai balance tdk boleh negativ
                        If objCd.CekNonNegatif(TxtBalance.Text) = True Then
                            Me.SalinKeObjek()
                            If baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.objTransCN.SimpanBaru(Me.UserCode, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.objTransCN.SimpanPerubahan(Me.UserCode, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnEdit.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            End If
                        Else
                            MsgBox("Please specify a positive number for AR/AP amount.", MsgBoxStyle.Information, "Information")
                            Me.TxtBalance.Focus()
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            '- cek nilai balance tdk blh negative
            If objCd.CekNonNegatif(TxtBalance.Text) = True Then
                If Me.ApakahDataBenar = True Then
                    Me.SalinKeObjek()
                    If baru = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                        If hasil = MsgBoxResult.Yes Then
                            '- simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objTransCN.SimpanBaru(Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                    Else
                        '- simpan perubahaan
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar = True Then
                                If Me.objTransCN.SimpanPerubahan(Me.UserCode, Me) Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnEdit.Focus()
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox("Please specify a positive number for AR/AP amount.", MsgBoxStyle.Information, "Information")
                Me.TxtBalance.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtCode.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.TxtPrimaryExpenseAcc.Text = "-1"
                        Me.TxtPrimaryArAccount.Text = "-1"
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.txtCustomerAddIdentifier.Text = ""
                    Else
                        Me.objTransCN.CariDetail(Me.objTransCN.Primary, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtNoTransaksi.Text <> "" Then
                Dim TipeBalance As Integer = 0
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                    If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                        TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
                    Else
                        TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
                    End If
                Else
                    If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                        TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                    Else
                        TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                    End If
                End If
                If Me.objTransCN.ApakahBolehEdit(Me.objTransCN.Primary, UserCode, True, Me.objTransCN.PrimaryJournal, TipeBalance, True) = True Then
                    Me.baru = False
                    Me.objTransCN.CariDetail(Me.objTransCN.Primary, Me)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtNoTransaksi.Text <> "" Then
                Dim TipeBalance As Integer = 0
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                    If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                        TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
                    Else
                        TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
                    End If
                Else
                    If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                        TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                    Else
                        TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                    End If
                End If

                If Me.objTransCN.ApakahBolehHapus(Me.txtPrimary.Text, True, Me.UserCode, Me.txtPrimaryJournal.Text, TipeBalance, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objTransCN.Hapus(Me.txtPrimary.Text, Me.TipeTransaksi, Me.TipeKartu, Me.txtPrimaryJournal.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Click
        Try
            If IsNumeric(Me.TxtExchangeRate.Text) = True Then
                objCd.Number_Click(TxtExchangeRate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            If IsNumeric(Me.TxtExchangeRate.Text) = True Then
                objCd.Number_Enter(TxtExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.GotFocus, _
      txtNama.GotFocus, txtCode.GotFocus, TxtContactPerson.GotFocus, TxtCustomerDebitNote.GotFocus, _
      TxtReferedInvoice.GotFocus, TxtExpenseAcc.GotFocus, TxtExpenseAccDesc.GotFocus, TxtBalance.GotFocus, _
      DtpDueDate.GotFocus, txtRemarks.GotFocus, txtBalanceReff.GotFocus, TxtUnbalance.GotFocus, TxtExchangeRate.GotFocus, _
      txtNoJurnal.GotFocus, txtDecJurnal.GotFocus, txtNoTransaksi.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnbalance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnbalance.Leave
        Try
            objCd.Number_Leave(TxtUnbalance, sender, e, DefaultPrecision, objCd)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.LostFocus, _
    txtNama.LostFocus, txtCode.LostFocus, TxtContactPerson.LostFocus, TxtCustomerDebitNote.LostFocus, _
    TxtReferedInvoice.LostFocus, TxtExpenseAcc.LostFocus, TxtExpenseAccDesc.LostFocus, TxtBalance.LostFocus, _
    DtpDueDate.LostFocus, txtRemarks.LostFocus, txtBalanceReff.LostFocus, TxtUnbalance.LostFocus, _
    TxtExchangeRate.LostFocus, txtNoJurnal.LostFocus, txtDecJurnal.LostFocus, txtNoTransaksi.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimary.BackColor
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim TTC As New FrmTransaksiCreditNote_cari
        Try
            TTC.BackColorAll = Me.BackColorAll
            TTC.TipeTransaksi = Me.TipeTransaksi

            TTC.ShowDialog()
            If TTC.PrimaryTransCN <> -1 Then
                If Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.ObjJurnal.ApakahMemilikiLedgerAccess(Me.UserCode, Me.objTransCN.GetCreatedBy(TTC.PrimaryTransCN)) = False Then
                        MsgBox("You are not authorized to view this data.", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Exit Sub
                    End If
                End If

                If Me.objTransCN.CariDetail(TTC.PrimaryTransCN, Me) = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TTC.Dispose()
        End Try
    End Sub

    Private Sub dtpdate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.ValueChanged
        Try
            Me.BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustSupp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustSupp.Click
        Try
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Dim hj As New Core_Module.FrmCariMasterCustomer
                hj.BackColorAll = Me.BackColorAll
                If Me.txtCode.Text <> "" Then
                    hj.txtCari.Text = Me.txtCode.Text
                End If
                hj.txtCari.Text = Me.txtCode.Text
                hj.ShowDialog()
                If hj.PrimaryCustomer <> -1 Then
                    Me.txtPrimaryCustSupp.Text = hj.PrimaryCustomer
                    Me.txtCode.Text = Me.objCust.GetCustomerCode(hj.PrimaryCustomer)
                    Me.txtNama.Text = Me.objCust.GetCustName(hj.PrimaryCustomer)
                    Me.TxtContactPerson.Text = objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
                    IsikanContactPerson_Address_Cust()
                    Me.TxtCustomerDebitNote.Focus()
                Else
                    Me.txtPrimaryCustSupp.Text = "-1"
                    Me.txtCode.Text = ""
                    Me.txtNama.Clear()
                    Me.txtCustomerAddIdentifier.Text = ""
                    Me.TxtContactPerson.Text = ""
                    Me.CboAddress.Items.Clear()
                End If
                If Me.txtPrimaryCustSupp.Text <> -1 Then
                    IsikanContactPerson_Address_Cust()
                End If
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                Dim hj As New Core_Module.FrmCariMasterSupplier
                hj.BackColorAll = Me.BackColorAll
                If Me.txtCode.Text <> "" Then
                    hj.txtCari.Text = Me.txtCode.Text
                End If
                hj.txtCari.Text = Me.txtCode.Text
                hj.ShowDialog()
                If hj.PrimarySupplier <> -1 Then
                    Me.txtPrimaryCustSupp.Text = hj.PrimarySupplier
                    Me.txtCode.Text = Me.objSupp.GetSupplierCode(hj.PrimarySupplier)
                    Me.txtNama.Text = Me.objSupp.GetSupplierName(hj.PrimarySupplier)
                    Me.TxtContactPerson.Text = objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
                    IsikanContactPerson_Address_Supp()
                    Me.TxtCustomerDebitNote.Focus()
                Else
                    Me.txtPrimaryCustSupp.Text = "-1"
                    Me.txtCode.Text = ""
                    Me.txtNama.Clear()
                    Me.txtCustomerAddIdentifier.Text = ""
                    Me.TxtContactPerson.Text = ""
                    Me.CboAddress.Items.Clear()
                End If
                If Me.txtPrimaryCustSupp.Text <> -1 Then
                    Me.IsikanContactPerson_Address_Supp()
                End If
            End If

            If Me.txtCode.Text <> "" AndAlso Me.TxtReferedInvoice.Text <> "" Then
                Me.CekCustomerSupplier()
                Me.CekNoinvoice()
                Me.cekTextInvoiceNumber()
                'Me.CekExpense_Account()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.Leave
        Try
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Dim hj As New Core_Module.FrmCariMasterCustomer
                hj.BackColorAll = Me.BackColorAll
                Try
                    If Me.txtCode.Text <> "" Then
                        If objCust.ApakahSudahAda(Me.txtCode.Text) = True Then
                            Me.txtPrimaryCustSupp.Text = objCust.GetPrimary(Me.txtCode.Text)
                            Me.txtCode.Text = objCust.GetCustomerCode(Me.txtPrimaryCustSupp.Text)
                            Me.txtNama.Text = objCust.GetCustName(Me.txtPrimaryCustSupp.Text)
                            Me.TxtContactPerson.Text = objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
                            Me.IsikanContactPerson_Address_Cust()
                        Else
                            Me.btnFindCustSupp.PerformClick()
                        End If
                    Else
                        Me.txtPrimaryCustSupp.Text = "-1"
                        Me.txtCode.Clear()
                        Me.txtNama.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.TxtContactPerson.Text = ""
                        Me.CboAddress.Items.Clear()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                Dim hj As New Core_Module.FrmCariMasterSupplier
                hj.BackColorAll = Me.BackColorAll
                Try
                    If Me.txtCode.Text <> "" Then
                        If objSupp.ApakahSudahAda(Me.txtCode.Text) = True Then
                            Me.txtPrimaryCustSupp.Text = objSupp.GetPrimary(Me.txtCode.Text)
                            Me.txtCode.Text = objSupp.GetSupplierCode(Me.txtPrimaryCustSupp.Text)
                            Me.txtNama.Text = objSupp.GetSupplierName(Me.txtPrimaryCustSupp.Text)
                            Me.TxtContactPerson.Text = objSupp.Get_ContactPerson(Me.txtPrimaryCustSupp.Text)
                            IsikanContactPerson_Address_Supp()
                        Else
                            Me.btnFindCustSupp.PerformClick()
                        End If
                    Else
                        Me.txtPrimaryCustSupp.Text = "-1"
                        Me.txtCode.Clear()
                        Me.txtNama.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.TxtContactPerson.Text = ""
                        Me.CboAddress.Items.Clear()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            End If
            If Me.txtCode.Text <> "" AndAlso Me.TxtReferedInvoice.Text <> "" Then
                Me.CekCustomerSupplier()
                Me.CekNoinvoice()
                Me.cekTextInvoiceNumber()
                'Me.CekExpense_Account()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyDown
        Try
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If e.KeyValue = 13 Then
                    If Me.txtCode.Text <> "" Then
                        If objCust.ApakahSudahAda(Me.txtCode.Text) = True Then
                            Me.txtPrimaryCustSupp.Text = objCust.GetPrimary(Me.txtCode.Text)
                            Me.txtCode.Text = objCust.GetCustomerCode(Me.txtPrimaryCustSupp.Text)
                            Me.txtNama.Text = objCust.GetCustName(Me.txtPrimaryCustSupp.Text)
                            Me.TxtContactPerson.Text = objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
                            IsikanContactPerson_Address_Cust()
                            Me.TxtCustomerDebitNote.Focus()
                        Else
                            Me.btnFindCustSupp.PerformClick()
                        End If
                    Else
                        Me.btnFindCustSupp.PerformClick()
                    End If
                End If
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If e.KeyValue = 13 Then
                    If Me.txtCode.Text <> "" Then
                        If objSupp.ApakahSudahAda(Me.txtCode.Text) = True Then
                            Me.txtPrimaryCustSupp.Text = objSupp.GetPrimary(Me.txtCode.Text)
                            Me.txtCode.Text = objSupp.GetSupplierCode(Me.txtPrimaryCustSupp.Text)
                            Me.txtNama.Text = objSupp.GetSupplierName(Me.txtPrimaryCustSupp.Text)
                            Me.TxtContactPerson.Text = objSupp.Get_ContactPerson(Me.txtPrimaryCustSupp.Text)
                            IsikanContactPerson_Address_Supp()
                            Me.TxtCustomerDebitNote.Focus()
                        Else
                            Me.btnFindCustSupp.PerformClick()
                        End If
                    Else
                        Me.btnFindCustSupp.PerformClick()
                    End If
                End If
            End If

            'If Me.txtCode.Text <> "" AndAlso Me.TxtReferedInvoice.Text <> "" Then
            '    Me.CekCustomerSupplier()
            '    Me.CekNoinvoice()
            '    Me.cekTextInvoiceNumber()
            '    Me.CekExpense_Account()
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboCustomerSuuplier_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboCustomerSuuplier.SelectedIndexChanged
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note

        Try
            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Me.txtSupplierInvoiceNo.Visible = False
                Me.lblSuppInvoiceNo.Visible = False
                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    Me.LblNote.Text = "Customer's Credit Note"
                Else
                    Me.LblNote.Text = "Customer's Debit Note"
                End If
                Me.LblRefered.Text = "Referred Invoice"
                Me.LblAcc.Text = "AR Account"
                Me.LblARAmount.Text = "AR Amount"
                Me.txtPrimaryCustSupp.Text = "-1"
                Me.TxtPrimaryReferedInvoice.Text = "-1"
                Me.txtSupplierInvoiceNo.Text = ""
                Me.TxtBalance.Text = 0
                Me.TxtReferedInvoice.Text = ""
                Me.TxtContactPerson.Text = ""
                Me.CboAddress.Items.Clear()
                Me.txtCode.Text = ""
                Me.txtRemarks.Clear()
                Me.txtNama.Text = ""
                Me.txtCode.Focus()
                Me.IsiCBOArAkunByTipePerkiraan1()
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, TipeTransaksi)
                End If
                Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
                Me.CekExpense_Account()
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    Me.LblNote.Text = "Supplier's Credit Note"
                Else
                    Me.LblNote.Text = "Supplier's Debit Note"
                End If
                Me.txtSupplierInvoiceNo.Visible = True
                Me.lblSuppInvoiceNo.Visible = True
                Me.LblRefered.Text = "Referred AOP"
                Me.LblAcc.Text = "AP Account"
                Me.LblARAmount.Text = "AP Amount"
                Me.txtPrimaryCustSupp.Text = "-1"
                Me.TxtPrimaryReferedInvoice.Text = "-1"
                Me.TxtReferedInvoice.Text = ""
                Me.TxtContactPerson.Text = ""
                Me.txtSupplierInvoiceNo.Text = ""
                Me.TxtBalance.Text = 0
                Me.CboAddress.Items.Clear()
                Me.txtCode.Text = ""
                Me.txtNama.Text = ""
                Me.txtRemarks.Clear()
                Me.txtCode.Focus()
                Me.IsiCBOAOPAkunByTipePerkiraan2()
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, TipeTransaksi)
                End If
                Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
                Me.CekExpense_Account()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FindExpenseAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindExpenseAccount.Click
        Dim hj As New FrmMasterAkun_Cari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCode.Text <> "" Then
                hj.txtCari.Text = Me.txtCode.Text
            End If
            hj.txtCari.Text = Me.TxtExpenseAcc.Text
            hj.ShowDialog()
            If hj.mPrimaryAkun <> -1 Then
                Dim Level As Integer = 0
                Level = objAkun.GetHeader(hj.mPrimaryAkun)
                If Level = 0 Then
                    Me.TxtPrimaryExpenseAcc.Text = hj.mPrimaryAkun
                    Me.TxtExpenseAcc.Text = Me.objAkun.GetAccountCode(hj.mPrimaryAkun)
                    Me.TxtExpenseAccDesc.Text = Me.objAkun.GetDescription_ByPrimaryAcc(hj.mPrimaryAkun)
                Else
                    MsgBox("You cannot entry header account for refered account.", MsgBoxStyle.Information, "Information")
                    Me.TxtPrimaryExpenseAcc.Text = -1
                    Me.TxtExpenseAcc.Text = ""
                    Me.TxtExpenseAccDesc.Text = ""
                    Exit Sub
                End If
            Else
                Me.CekExpense_Account()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExpenseAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExpenseAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtExpenseAcc.Text <> "" Then
                    If objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpenseAcc.Text) = True Then
                        Me.TxtPrimaryExpenseAcc.Text = objAkun.GetPrimary(Me.TxtExpenseAcc.Text)
                        Dim Level As Integer = 0
                        Level = objAkun.GetHeader(Me.TxtPrimaryExpenseAcc.Text)
                        If Level = 0 Then
                            Me.TxtExpenseAcc.Text = objAkun.GetAccountCode(Me.TxtPrimaryExpenseAcc.Text)
                            Me.TxtExpenseAccDesc.Text = objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryExpenseAcc.Text)
                        Else
                            MsgBox("You cannot entry header account for refered account.", MsgBoxStyle.Information, "Information")
                            'MsgBox("You should entry referred account with level = 4. ", MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryExpenseAcc.Text = -1
                            Me.TxtExpenseAcc.Text = ""
                            Me.TxtExpenseAccDesc.Text = ""
                            Exit Sub
                        End If
                         Me.cboCurrency.Focus()
                    Else
                        Me.FindExpenseAccount.PerformClick()
                    End If
                Else
                    Me.FindExpenseAccount.PerformClick()
                    '' Me.TxtExpenseAcc.Text = ""
                    ' Me.TxtExpenseAccDesc.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanContactPerson_Address_Supp()
        Try
            If Me.txtPrimaryCustSupp.Text <> "" And Me.txtPrimaryCustSupp.Text <> "-1" Then
                Me.objSupp.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                Me.CboAddress.SelectedIndex = 0
                Me.TxtContactPerson.Text = Me.objSupp.Get_ContactPerson(Me.txtPrimaryCustSupp.Text)
            Else
                Me.CboAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.TxtContactPerson.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanContactPerson_Address_Cust()
        Try
            If Me.txtPrimaryCustSupp.Text <> "" And Me.txtPrimaryCustSupp.Text <> "-1" Then
                Me.objCust.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                Me.CboAddress.SelectedIndex = 0
                Me.TxtContactPerson.Text = Me.objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
            Else
                Me.CboAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.TxtContactPerson.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboAddress.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                    If Me.CboAddress.SelectedIndex <> -1 Then
                        Me.txtCustomerAddIdentifier.Text = Me.objCust.GetIdentifierAdd(Me.txtPrimaryCustSupp.Text, _
                        Me.CboAddress.Text)
                    Else
                        Me.txtCustomerAddIdentifier.Text = ""
                    End If
                Else
                    If Me.CboAddress.SelectedIndex <> -1 Then
                        Me.txtCustomerAddIdentifier.Text = Me.objSupp.GetIdentifierAdd(Me.txtPrimaryCustSupp.Text, _
                        Me.CboAddress.Text)
                    Else
                        Me.txtCustomerAddIdentifier.Text = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkAddress.LinkClicked
        Try
            If Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                'supplier
                LLB_Address_Supp()
            Else
                'customer
                LLB_Address_Cust()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub LLB_Address_Supp()
        Dim ij As New Core_Module.FrmDetailAddressSupplier

        Try
            If Me.txtPrimaryCustSupp.Text = "" Or Me.txtPrimaryCustSupp.Text = "-1" Then
                If Me.txtRemarks.Enabled = True Then
                    MsgBox("Please specify supplier first.", MsgBoxStyle.Information, "Information")
                    Me.txtCode.Focus()
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If

            ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
            ij.PrimarySup = Me.txtPrimaryCustSupp.Text
            If Me.txtRemarks.Enabled = True Then
                ij.btnOK.BringToFront()
            Else
                ij.cboCustAddCode.Enabled = False
            End If
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.objSupp.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                Me.CboAddress.Text = Me.objSupp.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LLB_Address_Cust()
        Try
            If Me.txtPrimaryCustSupp.Text = "" Or Me.txtPrimaryCustSupp.Text = "-1" Then
                MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                Me.txtCode.Focus()
                Exit Sub
            End If
            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
            ij.PrimaryCust = Me.txtPrimaryCustSupp.Text
            ij.btnOK.BringToFront()
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.objCust.IsiCboAddress(Me.CboAddress, Me.txtPrimaryCustSupp.Text)
                Me.CboAddress.Text = Me.objCust.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindReferedInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindReferedInvoice.Click
        If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
            Try
                If Me.txtCode.Text <> "" Then
                    Dim hj As New FrmCariTransaksiSalesInvoice
                    hj.BackColorAll = Me.BackColorAll
                    If Me.TxtReferedInvoice.Text <> "" Then
                        hj.txtCari.Text = Me.TxtReferedInvoice.Text
                    End If
                    hj.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice
                    hj.ShowDialog()
                    If hj.PrimaryTransaksi > 0 Then
                        Me.TxtPrimaryReferedInvoice.Text = hj.PrimaryTransaksi
                        Me.TxtReferedInvoice.Text = Me.objSalesInv.GetInvoiceNumber(hj.PrimaryTransaksi)
                        Me.txtSupplierInvoiceNo.Text = ""
                        If Me.objSalesInv.GetPrimaryAkunAR(hj.PrimaryTransaksi) > -1 Then
                            Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryAkunAR(hj.PrimaryTransaksi)
                        Else
                            Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                        End If
                        Me.CekNoinvoice()
                        Me.CekCustomerSupplier()
                        Me.TxtExpenseAcc.Focus()
                    Else
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.TxtReferedInvoice.Text = ""
                        Me.txtSupplierInvoiceNo.Text = ""
                    End If
                    cekTextInvoiceNumber()
                Else
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.TxtReferedInvoice.Clear()
                    Me.TxtPrimaryReferedInvoice.Text = "-1"
                    Me.txtSupplierInvoiceNo.Text = ""
                    Me.txtCode.Focus()
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
            Try
                If Me.txtCode.Text <> "" Then
                    Dim hj As New FrmCariTransaksiAOP
                    hj.TipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
                    hj.BackColorAll = Me.BackColorAll
                    hj.FindCriteria = EnumSearchCriteria.ByNoTransaksi
                    If Me.TxtReferedInvoice.Text <> "" Then
                        hj.txtCari.Text = Me.TxtReferedInvoice.Text
                    End If
                    hj.ShowDialog()
                    If hj.PrimaryTransaksi > 0 Then
                        'If objAop.ApakahSudahAda(TxtReferedInvoice.Text) = True Then
                        Me.TxtPrimaryReferedInvoice.Text = hj.PrimaryTransaksi
                        Me.TxtReferedInvoice.Text = Me.objAop.GetNoTransaksi(hj.PrimaryTransaksi)
                        If hj.PrimaryTransaksi <> -1 Then
                            Me.txtSupplierInvoiceNo.Text = Me.objTransCN.GetSuppInvoiceNoByPrimaryAOP(hj.PrimaryTransaksi)
                        Else
                            Me.txtSupplierInvoiceNo.Text = "-"
                        End If
                        If Me.objAop.GetPrimaryAkunAP(hj.PrimaryTransaksi) > -1 Then
                            Me.TxtPrimaryArAccount.Text = Me.objAop.GetPrimaryAkunAP(hj.PrimaryTransaksi)
                        Else
                            Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                        End If
                        'Me.TxtPrimaryArAccount.Text = Me.objTransCN.GetPrimaryakunByNoInvoice(Me.TxtReferedInvoice.Text)
                        Me.CekNoinvoice()
                        Me.CekCustomerSupplier()
                        Me.TxtExpenseAcc.Focus()
                    Else
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.TxtReferedInvoice.Text = ""
                        Me.txtSupplierInvoiceNo.Text = ""
                    End If
                    ''End If
                    cekTextInvoiceNumber()
                Else
                    MsgBox("Please specify supplier first.", MsgBoxStyle.Information, "Information")
                    Me.TxtReferedInvoice.Clear()
                    Me.TxtPrimaryReferedInvoice.Text = "-1"
                    Me.txtSupplierInvoiceNo.Text = ""
                    Me.txtCode.Focus()
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
            End Try
        End If
    End Sub

    Private Sub TxtReferedInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtReferedInvoice.KeyDown
        Try
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                If e.KeyValue = 13 Then
                    If Me.TxtReferedInvoice.Text <> "" Then
                        If objSalesInv.ApakahSudahAda(Me.TxtReferedInvoice.Text) = True Then
                            Me.TxtPrimaryReferedInvoice.Text = objSalesInv.GetPrimary(Me.TxtReferedInvoice.Text)
                            Me.txtSupplierInvoiceNo.Text = ""
                            If Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text) > -1 Then
                                Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text)
                            Else
                                Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                            End If
                            Me.CekNoinvoice()
                            Me.CekCustomerSupplier()
                            Me.TxtExpenseAcc.Focus()
                        Else
                            Me.BtnFindReferedInvoice_Click(Nothing, Nothing)
                            ' Me.BtnFindReferedInvoice.PerformClick()
                            ' Me.txtSupplierInvoiceNo.Text = ""
                            ' Me.TxtBalance.Text = 0
                        End If
                    Else
                        Me.BtnFindReferedInvoice_Click(Nothing, Nothing)
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtBalance.Text = 0
                    End If
                    cekTextInvoiceNumber()
                End If
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                If e.KeyValue = 13 Then
                    If Me.TxtReferedInvoice.Text <> "" Then
                        If objAop.ApakahSudahAda(Me.TxtReferedInvoice.Text) = True Then
                            Me.TxtPrimaryReferedInvoice.Text = objAop.GetPrimary(Me.TxtReferedInvoice.Text)
                            If Me.TxtPrimaryReferedInvoice.Text <> -1 Then
                                Me.txtSupplierInvoiceNo.Text = Me.objTransCN.GetSuppInvoiceNoByPrimaryAOP(Me.TxtPrimaryReferedInvoice.Text)
                            Else
                                Me.txtSupplierInvoiceNo.Text = "-"
                            End If
                            If Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text) > -1 Then
                                Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text)
                            Else
                                Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                            End If
                            'Me.TxtPrimaryArAccount.Text = Me.objTransCN.GetPrimaryakunByNoInvoice(Me.TxtReferedInvoice.Text)
                            Me.CekNoinvoice()
                            Me.CekCustomerSupplier()
                            Me.TxtExpenseAcc.Focus()
                        Else
                            Me.BtnFindReferedInvoice.PerformClick()
                            ' Me.txtSupplierInvoiceNo.Text = ""
                            ' Me.TxtBalance.Text = 0
                        End If
                    Else
                        Me.BtnFindReferedInvoice.PerformClick()
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtBalance.Text = 0
                    End If
                    cekTextInvoiceNumber()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReferedInvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReferedInvoice.Leave
        Try
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Dim hj As New FrmCariTransaksiSalesInvoice
                hj.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice
                hj.BackColorAll = Me.BackColorAll
                Try
                    If Me.TxtReferedInvoice.Text <> "" Then
                        If Me.objSalesInv.ApakahSudahAda(Me.TxtReferedInvoice.Text) = True Then
                            Me.TxtPrimaryReferedInvoice.Text = objSalesInv.GetPrimary(Me.TxtReferedInvoice.Text)
                            Me.txtSupplierInvoiceNo.Text = ""
                            If Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text) > -1 Then
                                Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text)
                            Else
                                Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                            End If
                            'Me.TxtReferedInvoice.Text = objSalesInv.GetNoInvoice(Me.TxtPrimaryReferedInvoice.Text)
                            Me.CekNoinvoice()
                            Me.CekCustomerSupplier()
                        Else
                            Me.BtnFindReferedInvoice.PerformClick()
                            Me.txtSupplierInvoiceNo.Text = ""
                        End If
                    Else
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtReferedInvoice.Clear()
                    End If
                    cekTextInvoiceNumber()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                Dim hj As New FrmCariTransaksiAOP
                hj.TipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
                hj.BackColorAll = Me.BackColorAll
                Try
                    If Me.TxtReferedInvoice.Text <> "" Then
                        If Me.objAop.ApakahSudahAda(Me.TxtReferedInvoice.Text) = True Then
                            Me.TxtPrimaryReferedInvoice.Text = objAop.GetPrimary(Me.TxtReferedInvoice.Text)

                            If hj.PrimaryTransaksi <> -1 Then
                                Me.txtSupplierInvoiceNo.Text = Me.objTransCN.GetSuppInvoiceNoByPrimaryAOP(hj.PrimaryTransaksi)
                            Else
                                Me.txtSupplierInvoiceNo.Text = "-"
                            End If

                            If Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text) > -1 Then
                                Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryCurrency(Me.TxtPrimaryReferedInvoice.Text)
                            Else
                                Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text
                            End If
                            'Me.TxtReferedInvoice.Text = objAop.GetNoTransaksi(Me.TxtPrimaryReferedInvoice.Text)
                            Me.CekNoinvoice()
                            Me.CekCustomerSupplier()
                        Else
                            Me.BtnFindReferedInvoice.PerformClick()
                        End If
                    Else
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtBalance.Text = 0
                        Me.TxtReferedInvoice.Clear()
                    End If
                    cekTextInvoiceNumber()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New ClsMasterAkun.EnumTipeAkun
        Dim tp2 As New ClsMasterAkun.EnumTipeAkun2
        tp1 = ClsMasterAkun.EnumTipeAkun.Asset
        tp2 = ClsMasterAkun.EnumTipeAkun2.Account_Receivable
        Try
            If Me.txtNoTransaksi.Text <> "-1" Then
                Me.CboArAccount.DataSource = Nothing
                Me.CboArAccount.Items.Clear()
                Me.cboPrimaryReffAcc.DataSource = Nothing
                Me.cboPrimaryReffAcc.Items.Clear()
                Me.objTransCN.IsiTBLArAcount(tp1, tp2, TblArAccount)
                Me.CboArAccount.DataSource = Me.TblArAccount.DefaultView
                Me.CboArAccount.DisplayMember = "Account Code"
                Me.cboPrimaryReffAcc.DataSource = Me.TblArAccount.DefaultView
                Me.cboPrimaryReffAcc.DisplayMember = "Primary"
                Me.TxtPrimaryArAccount.Text = Me.cboPrimaryReffAcc.Text

                'Me.IsiCBOCurrency()
                Me.CboArAccount_SelectedIndexChanged(Nothing, Nothing)

                'Me.TxtPrimaryCurrency.Text = Me.ComboBox1.Text
                'Me.TxtPrimaryCurrency.Text = Me.TblArAccount(Me.CboArAccount.SelectedIndex).primary_currency
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOAOPAkunByTipePerkiraan2()
        '- isicbo supplier
        Dim tp1 As New ClsMasterAkun.EnumTipeAkun
        Dim tp2 As New ClsMasterAkun.EnumTipeAkun2
        tp1 = ClsMasterAkun.EnumTipeAkun.Liabilities
        tp2 = ClsMasterAkun.EnumTipeAkun2.Account_Payable
        Try
            Me.CboArAccount.DataSource = Nothing
            Me.CboArAccount.Items.Clear()
            Me.objTransCN.IsiTBLArAcount(tp1, tp2, TblArAccount)
            Me.CboArAccount.DataSource = Me.TblArAccount.DefaultView
            Me.CboArAccount.DisplayMember = "Account Code"
            Me.TxtPrimaryCurrency.Text = Me.TblArAccount(Me.CboArAccount.SelectedIndex).primary_currency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboArAccount_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboArAccount.SelectedIndexChanged
        Try
            If Me.CboArAccount.Items.Count > 0 AndAlso Me.cboPrimaryReffAcc.Items.Count > 0 Then
                Me.cboPrimaryReffAcc.SelectedIndex = Me.CboArAccount.SelectedIndex
                Dim PrimaryAkun As Double = -1
                Dim strKodeAccount As String = ""
                If IsNumeric(Me.cboPrimaryReffAcc.Text) Then
                    PrimaryAkun = CDbl(Me.cboPrimaryReffAcc.Text)
                End If
                strKodeAccount = Me.objAkun.GetAccountCode(PrimaryAkun)
                Me.TxtPrimaryArAccount.Text = PrimaryAkun
                Me.TxtPrimaryCurrency.Text = Me.objAkun.GetPrimaryCurrency(PrimaryAkun)
                Me.IsiCBOCurrency()
                If Me.TxtPrimaryCurrency.Text <> "-1" Then
                    If Me.TxtPrimaryArAccount.Text <> Me.PrimaryAccAr_Lama Then
                        Me.TxtExchangeRate.Text = Me.objCd.DigitGrouping(Me.objTransCN.Get_DefaultExchangerateByPrimary(Me.TxtPrimaryCurrency.Text, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE))
                    End If
                End If

                If IsNumeric(Me.TxtPrimaryExpenseAcc.Text) AndAlso CDbl(Me.TxtPrimaryExpenseAcc.Text) Then
                    Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                    If Not String.IsNullOrEmpty(Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)) Then
                        CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                    Else
                        CurrCodeReffAcc = Me.cboCurrency.Text
                    End If


                    Me.IsiCBOCurrencyReff()
                    If Me.cboCurrency.Text = "IDR" Then
                        Me.TxtExchangeRate.Text = "1"
                        Me.TxtExchangeRate.ReadOnly = True
                        Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                        Me.TxtExchangeRate.TabStop = False
                    Else
                        Me.TxtExchangeRate.ReadOnly = False
                        Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                        Me.TxtExchangeRate.TabStop = True
                    End If

                    Dim deffPrecReff As Integer = 0
                    deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

                    If Me.cboCurrencyReff.Text = "IDR" Then
                        If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                        Else
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        End If
                        Me.cboCurrencyReff.Enabled = False
                        Me.txtBalanceReff.Enabled = False
                        Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                    Else
                        If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                            If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                                Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                            Else
                                Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                            End If

                            Me.cboCurrencyReff.Enabled = False
                            Me.txtBalanceReff.Enabled = False
                            Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                        Else
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                            Me.cboCurrencyReff.Enabled = True
                            Me.txtBalanceReff.Enabled = True
                            Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor

                        End If
                    End If

                End If

                If Me.Active = True And Me.TxtPrimaryArAccount.Text <> -1 And Me.Lanjut = True Then
                    If Me.TxtPrimaryArAccount.Text <> Me.PrimaryAccAr_Lama Then
                        Me.dgvItem.Rows.Clear()
                        Me.dgvItem.Rows.Add()
                        Me.IsiDgv_AddItemDefault()
                    End If
                End If
                'Me.TxtPrimaryCurrency.Text = Me.objTransCN.GetPrimarCurrencyByKodeAkunAr(Me.ComboBox1.Text)
                'Me.TxtPrimaryArAccount.Text = Me.objAkun.GetPrimary(Me.ComboBox1.Text)
                'Me.IsiCBOCurrency()
                'If Me.TxtPrimaryCurrency.Text <> "-1" Then
                '    Me.TxtExchangeRate.Text = Me.objTransCN.Get_DefaultExchangerateByPrimary(Me.TxtPrimaryCurrency.Text, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE)
                'End If
                'If Me.cboCurrency.Text = "IDR" Then
                '    Me.TxtExchangeRate.Text = "1"
                '    Me.TxtExchangeRate.ReadOnly = True
                '    Me.TxtExchangeRate.TabStop = False
                'Else
                '    Me.TxtExchangeRate.TabStop = True
                '    Me.TxtExchangeRate.ReadOnly = False
                'End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.PrimaryAccAr_Lama = Me.TxtPrimaryArAccount.Text
        End Try
    End Sub
    Private Sub IsiCBOCurrency()
        Try
            Me.cboCurrency.DataSource = Nothing
            Me.cboCurrency.Items.Clear()
            If Me.TxtPrimaryCurrency.Text <> "-1" Then
                Me.objTransCN.IsiTBLCurrency(Me.TxtPrimaryCurrency.Text, TBlCurrency)
                Me.cboCurrency.DataSource = Me.TBlCurrency.DefaultView
                Me.TblArAccount.DefaultView.Sort = "Primary"
                Me.cboCurrency.DisplayMember = "Currency Code"
            Else
                Me.cboCurrency.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOCurrencyReff()
        Try
            Me.cboCurrencyReff.DataSource = Nothing
            Me.cboCurrencyReff.Items.Clear()
            If Me.PrimaryCurrReffAcc <> -1 Then
                Me.objTransCN.IsiTBLCurrency(PrimaryCurrReffAcc, TBlCurrencyReff)
                Me.cboCurrencyReff.DataSource = Me.TBlCurrencyReff.DefaultView
                Me.cboCurrencyReff.DisplayMember = "Currency Code"
            Else
                Me.cboCurrencyReff.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekNoinvoice()
        Try
            If Me.TxtPrimaryReferedInvoice.Text = "-1" Then
                'Me.TxtExchangeRate.Enabled = True
                Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                Me.TxtExchangeRate.ReadOnly = False
                Me.TxtExchangeRate.TabStop = True
                Exit Sub
            Else
                '- isi balance
                If baru = True Then
                    '- jika customer isi balance by sales
                    If CboCustomerSuuplier.SelectedIndex = 0 Then
                        TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
                        Me.TxtBalance.Text = Me.objCd.DigitGrouping(Me.objTransCN.Get_TotalBySales(Me.TxtPrimaryReferedInvoice.Text, Me.TipeTransaksi), Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.txtSupplierInvoiceNo.Visible = False
                    ElseIf CboCustomerSuuplier.SelectedIndex = 1 Then
                        '- jika supplier isi balance by AOP
                        TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
                        Me.TxtBalance.Text = Me.objCd.DigitGrouping(Me.objTransCN.Get_TotalAmountByAOP(Me.TxtPrimaryReferedInvoice.Text, Me.TipeTransaksi), Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        If Me.TxtPrimaryReferedInvoice.Text <> -1 Then
                            Me.txtSupplierInvoiceNo.Text = Me.objAop.GetSuppInvoiceNo(Me.TxtPrimaryReferedInvoice.Text)
                        Else
                            Me.txtSupplierInvoiceNo.Text = "-"
                        End If
                        Me.txtSupplierInvoiceNo.Visible = True
                    End If
                    '- balance < batas max
                    BatasBalanceMaksimum = CDbl(Me.TxtBalance.Text)
                    Me.TxtExchangeRate.Text = Me.objTransCN.Get_DefaultExchangerateByPrimary(Me.TxtPrimaryCurrency.Text, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE)
                    Me.TxtExchangeRate.Text = Me.objCd.DigitGrouping(Me.TxtExchangeRate.Text, DeffPrecX)
                    If CboCustomerSuuplier.SelectedIndex = 0 Then
                        Me.TxtPrimaryArAccount.Text = Me.objSalesInv.GetPrimaryAkunAR(Me.TxtPrimaryReferedInvoice.Text)
                    ElseIf CboCustomerSuuplier.SelectedIndex = 1 Then
                        Me.TxtPrimaryArAccount.Text = Me.objAop.GetPrimaryAkunAP(Me.TxtPrimaryReferedInvoice.Text)
                    End If
                    ''- menentukan code account 
                    For intA As Integer = 0 To TblArAccount.Rows.Count - 1
                        RowArAccount = TblArAccount.Rows(intA)
                        If RowArAccount.Primary = CDbl(Me.TxtPrimaryArAccount.Text) Then
                            Me.CboArAccount.SelectedIndex = intA
                        End If
                    Next

                    'For Each RowArAccount In TblArAccount.Rows
                    '    If RowArAccount.Primary = Me.TxtPrimaryArAccount.Text Then
                    '        'Me.CboArAccount.SelectedIndex = Me.TblArAccount.Rows.IndexOf(RowArAccount)
                    '        Dim IntA As Integer = 0
                    '         Me.CboArAccount.SelectedIndex = 
                    '    End If
                    'Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
    Private Sub cekTextInvoiceNumber()
        Try
            If Me.TxtPrimaryReferedInvoice.Text = "-1" Then
                Me.cboCurrency.Enabled = True
                Me.CboArAccount.Enabled = True
                Me.TxtExchangeRate.Enabled = True
                If TxtPrimaryCurrency.Text <> "-1" Then
                    Me.TxtExchangeRate.Text = Me.objTransCN.Get_DefaultExchangerateByPrimary(Me.TxtPrimaryCurrency.Text, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE)
                End If
            Else
                Me.TxtExchangeRate.ReadOnly = True
                If TxtPrimaryCurrency.Text <> "-1" Then
                    Me.TxtExchangeRate.Text = Me.objTransCN.Get_DefaultExchangerateByPrimary(Me.TxtPrimaryCurrency.Text, ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE)
                End If
                Me.cboCurrency.Enabled = False
                Me.CboArAccount.Enabled = False
            End If
            If Me.cboCurrency.Text = "IDR" Then
                Me.TxtExchangeRate.Text = "1"
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
            Else
                Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                Me.TxtExchangeRate.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekCustomerSupplier()
        Try
            Dim hasil1 As MsgBoxResult
            If Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                Dim P1 As String = CStr(Me.txtPrimaryCustSupp.Text)
                Dim P2 As String = Me.objTransCN.GetPrimaryCustInvoice(Me.TxtPrimaryReferedInvoice.Text)
                If P1 <> P2 Then
                    '- jika no invoice = customer code
                    hasil1 = MsgBox("The invoice is not for the specified customer, continue and change customer data?", MsgBoxStyle.YesNo, "Confirmation")
                    If hasil1 = MsgBoxResult.Yes Then
                        Me.txtPrimaryCustSupp.Text = Me.objTransCN.GetPrimaryCustInvoice(Me.TxtPrimaryReferedInvoice.Text)
                        Me.txtCode.Text = Me.objCust.GetCustomerCode(txtPrimaryCustSupp.Text)
                        Me.txtNama.Text = Me.objCust.GetCustName(txtPrimaryCustSupp.Text)
                        Me.TxtContactPerson.Text = objCust.GetCustContactPerson(Me.txtPrimaryCustSupp.Text)
                        Me.txtSupplierInvoiceNo.Text = ""
                        IsikanContactPerson_Address_Cust()
                    Else
                        '-jika <>
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.TxtReferedInvoice.Text = ""
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtBalance.Text = 0
                        Me.TxtReferedInvoice.Focus()
                    End If
                End If
            ElseIf Me.CboCustomerSuuplier.SelectedIndex = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                Dim P1 As String = CStr(Me.txtPrimaryCustSupp.Text)
                Dim P2 As String = Me.objTransCN.GetPrimarySupplierAOP(Me.TxtPrimaryReferedInvoice.Text)
                If P1 <> P2 Then
                    hasil1 = MsgBox("The AOP is not from the specified supplier, continue and change supplier data?", MsgBoxStyle.YesNo, "Information")
                    If hasil1 = MsgBoxResult.Yes Then
                        Me.txtPrimaryCustSupp.Text = Me.objTransCN.GetPrimarySupplierAOP(Me.TxtPrimaryReferedInvoice.Text)
                        Me.txtCode.Text = Me.objSupp.GetSupplierCode(Me.txtPrimaryCustSupp.Text)
                        Me.txtNama.Text = Me.objSupp.GetSupplierName(Me.txtPrimaryCustSupp.Text)
                        Me.TxtContactPerson.Text = objSupp.Get_ContactPerson(Me.txtPrimaryCustSupp.Text)
                        If Me.TxtPrimaryReferedInvoice.Text <> -1 Then
                            Me.txtSupplierInvoiceNo.Text = Me.objAop.GetSuppInvoiceNo(Me.TxtPrimaryReferedInvoice.Text)
                        Else
                            Me.txtSupplierInvoiceNo.Text = "-"
                        End If
                        IsikanContactPerson_Address_Supp()
                    Else
                        Me.TxtPrimaryReferedInvoice.Text = "-1"
                        Me.txtSupplierInvoiceNo.Text = ""
                        Me.TxtReferedInvoice.Clear()
                        Me.TxtReferedInvoice.Focus()
                        Me.TxtBalance.Text = 0
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub TxtExpenseAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExpenseAcc.Leave
        Dim hj As New FrmMasterAkun_Cari
        hj.BackColorAll = Me.BackColorAll
        'Variable 'ProsesJalan' dan 'SedangLeave' terpaksa diadakan karena suka keluar 2x form carinya
        Dim ProsesJalan As Boolean = False

        Try
            If Me.SedangLeave = 0 Then
                Me.SedangLeave = 1
                ProsesJalan = True

                If Me.TxtExpenseAcc.Text <> "" Then
                    If objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtExpenseAcc.Text) = True Then
                        Dim level As Integer = 0
                        Me.TxtPrimaryExpenseAcc.Text = objAkun.GetPrimary(Me.TxtExpenseAcc.Text)
                        level = objAkun.GetHeader(Me.TxtPrimaryExpenseAcc.Text)
                        If level = 0 Then
                            Me.TxtExpenseAcc.Text = objAkun.GetAccountCode(Me.TxtPrimaryExpenseAcc.Text)
                            Me.TxtExpenseAccDesc.Text = objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryExpenseAcc.Text)
                        Else
                            MsgBox("You cannot entry header account for refered account.", MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryExpenseAcc.Text = -1
                            Me.TxtExpenseAcc.Text = ""
                            Me.TxtExpenseAccDesc.Text = ""
                            Exit Sub
                        End If
                    Else
                        Me.FindExpenseAccount.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryExpenseAcc.Text = "-1"
                    Me.TxtExpenseAcc.Text = ""
                    Me.TxtExpenseAccDesc.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
            If ProsesJalan = True Then
                Me.SedangLeave = 0
            End If
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Try
            objCd.Number_KeyDown(TxtExchangeRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Try
            objCd.Number_KeyUp(TxtExchangeRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        ' Dim vDef_Prec As String = objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Try
            If Me.TxtExchangeRate.Text = "0" Then Me.TxtExchangeRate.Text = "1"
            Me.objCd.Number_Leave(TxtExchangeRate, sender, e, DeffPrecX, objCd)
            Me.UpdateBalance()
            Me.HitungUnbalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCreditNote_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.LinkAddress.TabStop = False
            If Me.CboCustomerSuuplier.SelectedIndex = 0 Then
                Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
                Me.txtSupplierInvoiceNo.Visible = False
                Me.lblSuppInvoiceNo.Visible = False
            Else
                Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
                Me.txtSupplierInvoiceNo.Visible = True
                Me.lblSuppInvoiceNo.Visible = True
            End If
            Me.TipeKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
            If TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                Me.Text = "Credit Note"
            Else
                Me.Text = "Debit Note"
            End If
            Me.baru = True
            Me.KosongkanField()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.NonAktifkanDetail()
            ''
            'Me.ColPrimaryCurrency.Visible = True
            'Me.ColDefPrec.Visible = True
            'Me.ColExchangeRate.Visible = True
            'Me.ColDebet_OriginalCurr.Visible = True
            'Me.ColCredit_OriginalCurrency.Visible = True
            'Me.ColOriginalCurrency.Visible = True
            ''
            Me.btnFind.Select()

            Me.DefaultPrecision = Me.objCurr.GetDecimal_DefaultCurrency
            'Me.gboDetail.Width = Me.Width - 10
            'Me.gboDetail.Left = 5
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtBalance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBalance.KeyDown
        Try
            If IsNumeric(Me.TxtBalance.Text) Then
                objCd.Number_KeyDown(TxtBalance, sender, e)
            Else
                TxtBalance.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtBalance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBalance.KeyUp
        Try
            If IsNumeric(Me.TxtBalance.Text) Then
                objCd.Number_KeyUp(TxtBalance, sender, e, False)
            Else
                TxtBalance.Text = 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.Click
        Try
            If IsNumeric(Me.TxtBalance.Text) Then
                objCd.Number_Click(TxtBalance)
            Else
                TxtBalance.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.Enter
        Try
            If IsNumeric(Me.TxtBalance.Text) Then
                objCd.Number_Enter(TxtBalance, sender, e)
            Else
                TxtBalance.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryArAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryArAccount.TextChanged
        Dim PrimaryARAkun As Double = -1


        Try
            If IsNumeric(Me.TxtPrimaryArAccount.Text) Then
                PrimaryARAkun = CDbl(Me.TxtPrimaryArAccount.Text)
            End If

            If Me.TblArAccount.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.TblArAccount.Rows.Count - 1
                    Me.RowArAccount = Me.TblArAccount.Rows(IntA)
                    With Me.RowArAccount
                        If .Primary = PrimaryARAkun Then
                            Me.cboPrimaryReffAcc.SelectedIndex = IntA
                            Me.CboArAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

            If IsNumeric(Me.TxtPrimaryExpenseAcc.Text) AndAlso CDbl(Me.TxtPrimaryExpenseAcc.Text) Then
                Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                If Not String.IsNullOrEmpty(Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)) Then
                    CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                Else
                    CurrCodeReffAcc = Me.cboCurrency.Text
                End If

                Me.IsiCBOCurrencyReff()
                If Me.cboCurrency.Text = "IDR" Then
                    Me.TxtExchangeRate.Text = "1"
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                    Me.TxtExchangeRate.ReadOnly = False
                End If

                Dim deffPrecReff As Integer = 0
                deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

                If Me.cboCurrencyReff.Text = "IDR" Then
                    If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    End If
                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor
                Else
                    If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                        If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                        Else
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        End If
                        Me.cboCurrencyReff.Enabled = False
                        Me.txtBalanceReff.Enabled = False
                        Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        Me.cboCurrencyReff.Enabled = True
                        Me.txtBalanceReff.Enabled = True
                        Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor
                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryExpenseAcc_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryExpenseAcc.TextChanged

        Try
            If IsNumeric(Me.TxtPrimaryExpenseAcc.Text) AndAlso CDbl(Me.TxtPrimaryExpenseAcc.Text) Then
                Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                If Not String.IsNullOrEmpty(Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)) Then
                    CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                Else
                    CurrCodeReffAcc = Me.cboCurrency.Text
                End If


                Me.IsiCBOCurrencyReff()
                If Me.cboCurrency.Text = "IDR" Then
                    Me.TxtExchangeRate.Text = "1"
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.ReadOnly = False
                    Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                End If

                Dim deffPrecReff As Integer = 0
                deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

                If Me.cboCurrencyReff.Text = "IDR" Then
                    If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    End If
                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                        If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                        Else
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        End If

                        Me.cboCurrencyReff.Enabled = False
                        Me.txtBalanceReff.Enabled = False
                        Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        Me.cboCurrencyReff.Enabled = True
                        Me.txtBalanceReff.Enabled = True
                        Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor

                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryReferedInvoice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryReferedInvoice.TextChanged

        Try

            If IsNumeric(Me.TxtPrimaryExpenseAcc.Text) AndAlso CDbl(Me.TxtPrimaryExpenseAcc.Text) Then
                Me.PrimaryCurrReffAcc = Me.objAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAcc.Text)
                If Not String.IsNullOrEmpty(Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)) Then
                    CurrCodeReffAcc = Me.objCurr.GetCurrencyCode(Me.PrimaryCurrReffAcc)
                Else
                    CurrCodeReffAcc = Me.cboCurrency.Text
                End If

                Me.IsiCBOCurrencyReff()

                If Me.cboCurrency.Text = "IDR" Then
                    Me.TxtExchangeRate.Text = "1"
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.ReadOnly = False
                    Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                End If

                Dim deffPrecReff As Integer = 0
                deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

                If Me.cboCurrencyReff.Text = "IDR" Then
                    If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    End If
                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                        If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                        Else
                            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        End If

                        Me.cboCurrencyReff.Enabled = False
                        Me.txtBalanceReff.Enabled = False
                        Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                        Me.cboCurrencyReff.Enabled = True
                        Me.txtBalanceReff.Enabled = True
                        Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor

                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.TextChanged

        Try
            Me.IsiCBOCurrencyReff()
            If Me.cboCurrency.Text = "IDR" Then
                Me.TxtExchangeRate.Text = "1"
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.TxtExchangeRate.ReadOnly = True
            Else
                Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                Me.TxtExchangeRate.ReadOnly = False
            End If

            Dim deffPrecReff As Integer = 0
            deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

            If Me.cboCurrencyReff.Text = "IDR" Then
                If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                Else
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                End If
                Me.cboCurrencyReff.Enabled = False
                Me.txtBalanceReff.Enabled = False
                Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

            Else
                If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                    If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    End If

                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    Me.cboCurrencyReff.Enabled = True
                    Me.txtBalanceReff.Enabled = True
                    Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.TextChanged

        Try
            Me.IsiCBOCurrencyReff()
            If Me.cboCurrency.Text = "IDR" Then
                Me.TxtExchangeRate.Text = "1"
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.TxtExchangeRate.ReadOnly = True
            Else
                Me.TxtExchangeRate.BackColor = Me.txtRemarks.BackColor
                Me.TxtExchangeRate.ReadOnly = False
            End If

            Dim deffPrecReff As Integer = 0
            deffPrecReff = Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)

            If Me.cboCurrencyReff.Text = "IDR" Then
                If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text) * CDec(Me.TxtExchangeRate.Text), deffPrecReff)
                Else
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                End If
                Me.cboCurrencyReff.Enabled = False
                Me.txtBalanceReff.Enabled = False
                Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

            Else
                If Me.cboCurrencyReff.Text = Me.cboCurrency.Text Then
                    If IsNumeric(Me.TxtBalance.Text) AndAlso IsNumeric(Me.TxtExchangeRate.Text) Then
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(CDec(Me.TxtBalance.Text), deffPrecReff)
                    Else
                        Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    End If

                    Me.cboCurrencyReff.Enabled = False
                    Me.txtBalanceReff.Enabled = False
                    Me.txtBalanceReff.BackColor = Me.txtReadOnly.BackColor

                Else
                    Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(0, deffPrecReff)
                    Me.cboCurrencyReff.Enabled = True
                    Me.txtBalanceReff.Enabled = True
                    Me.txtBalanceReff.BackColor = Me.txtRemarks.BackColor

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub txtBalanceReff_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBalanceReff.Enter
        Try
            If IsNumeric(Me.txtBalanceReff.Text) Then
                objCd.Number_Enter(txtBalanceReff, sender, e)
            Else
                txtBalanceReff.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBalanceReff_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBalanceReff.KeyDown
        Try
            If IsNumeric(Me.txtBalanceReff.Text) Then
                objCd.Number_KeyDown(txtBalanceReff, sender, e)
            Else
                txtBalanceReff.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBalanceReff_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBalanceReff.KeyUp
        Try
            If IsNumeric(Me.txtBalanceReff.Text) Then
                objCd.Number_KeyUp(txtBalanceReff, sender, e, False)
            Else
                txtBalanceReff.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBalanceReff_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBalanceReff.Leave
        Dim vDef_Prec As String = objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)
        Try
            If IsNumeric(Me.txtBalanceReff.Text) Then
                objCd.Number_Leave(txtBalanceReff, sender, e, vDef_Prec, objCd)
            Else
                Me.txtBalanceReff.Text = 0
            End If

            'If baru = True Then
            '    '- menentukan batas max balance
            '    If Me.TxtPrimaryReferedInvoice.Text <> "-1" Then
            '        If IsNumeric(Me.txtBalanceReff.Text) = True Then
            '            If CDbl(Me.txtBalanceReff.Text) > Me.BatasBalanceMaksimum Then
            '                MsgBox("Maximum amount is : " & Me.objCd.DigitGrouping(Me.BatasBalanceMaksimum, Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)), MsgBoxStyle.Information, "Information")
            '                Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(Me.BatasBalanceMaksimum, Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc))
            '            End If
            '        End If
            '    End If
            'Else
            '    If IsNumeric(Me.txtBalanceReff.Text) = True Then
            '        If CDbl(Me.txtBalanceReff.Text) > Me.batasBalanceMaksimum2 Then
            '            MsgBox("Maximum amount is : " & Me.objCd.DigitGrouping(Me.batasBalanceMaksimum2, Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc)), MsgBoxStyle.Information, "Information")
            '            Me.txtBalanceReff.Text = Me.objCd.DigitGrouping(Me.batasBalanceMaksimum2, Me.objCurr.GetDefaultPrecission(Me.PrimaryCurrReffAcc))
            '        End If
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim I As New FrmTransaksiCreditNote_Cetak
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.txtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPrimary.Text <> -1 Then
                If Me.txtPrimaryJournal.Text = -1 Then
                    MsgBox("This data doesn't have journal transaction, you only can print if data already exist in journal transaction." + vbCrLf + _
                    "Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    btnEdit.Focus()
                    Exit Sub
                End If
                I.TipeKartu = Me.TipeKartu
                I.TipeTransaksi = Me.TipeTransaksi
                I.PrimaryJournal = Me.txtPrimaryJournal.Text
                I.PrimaryCustSupp = Me.txtPrimaryCustSupp.Text
                I.PrimaryCurrAkunAR = Me.TxtPrimaryCurrency.Text
                I.PrimaryCurrAkunReff = Me.PrimaryCurrReffAcc
                I.ShowDialog()
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_CreditDebitNote_Create(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryInv_New As Double)
        Try
            Me.baru = False

            Me.objTransCN.CariDetail(vPrimaryTransaksi, Me)
            Me.UserCode = Me.objTransCN.CreatedBy
            Me.objTransCN.PrimaryInvoice = vPrimaryInv_New
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_CreditDebitNote_Save() As Double
        Try
            Me.baru = True
            Me.objTransCN.SimpanBaru1(Me.UserCode, Me, False, True)
            Return Me.objTransCN.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note

        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
                Else
                    TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
                End If
                If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                    If Me.baru = False Then
                        If Me.objTransCN.Tgl.Month = Me.DtpDate.Value.Month And _
                        Me.objTransCN.Tgl.Year = Me.DtpDate.Value.Year Then
                            Me.txtNoTransaksi.Text = Me.objTransCN.NoTrans
                        Else
                            If Me.ObjNumbering.GetResetType(TipeNumbering) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Yearly Then
                                'klo yearly dan tahunnya berubah, maka ambil nomor baru
                                If Me.objTransCN.Tgl.Year <> Me.DtpDate.Value.Year Then
                                    Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, TipeTransaksi)
                                End If
                            Else
                                'klo monthly, bulan atau tahun berubah, ambil nomor baru
                                If Me.objTransCN.Tgl.Month <> Me.DtpDate.Value.Month Or _
                                Me.objTransCN.Tgl.Year <> Me.DtpDate.Value.Year Then
                                    If Me.objTransCN.Tgl.Year <> Me.DtpDate.Value.Year Then
                                        Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, TipeTransaksi)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Me.txtNoTransaksi.Text = Me.objTransCN.GetNoBaru(Me.DtpDate.Value, TipeTransaksi)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit


        Try
            If e.RowIndex = 0 Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            Else
                If e.ColumnIndex <> Me.ColAccount.Index Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColAccount.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                'If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Then
                '    e.Cancel = True
                '    Exit Sub
                'End If
            End If

            'If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
            '    Exit Sub
            'End If

            If e.ColumnIndex <> Me.ColAccount.Index Then
                If e.ColumnIndex <> Me.ColRemarks.Index And e.ColumnIndex <> Me.ColOriginalCurrency.Index And e.ColumnIndex <> Me.ColDescAccount.Index Then
                    If e.ColumnIndex = Me.ColDebet.Index Or e.ColumnIndex = Me.ColKredit.Index Or _
                    e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Or _
                    e.ColumnIndex = Me.ColExchangeRate.Index Or e.ColumnIndex = Me.ColJob.Index Or _
                    e.ColumnIndex = Me.ColDept.Index Then
                        'jika belum diisikan account codenya, maka tidak bisa edit dept dan job
                        If e.ColumnIndex = Me.ColDept.Index Or e.ColumnIndex = Me.ColJob.Index Then
                            If Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = -1 Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        'untuk kolom exchange rate, debet original, creadit original tidak bisa diedit untuk akun dengan mata uang idr
                        If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurr Then
                            If e.ColumnIndex = Me.ColExchangeRate.Index Or _
                            e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Or _
                            e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If IsNothing(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value) = True Then
                            'bukan baris teakhir
                            e.Cancel = True
                            Exit Sub
                        ElseIf Me.objAkun.ApakahSudahAda_Aktif_NonHeader2(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value) = False Then
                            'ada akunnya, aktif, dan bukan header
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            'buang nol dibelakang decimal
                            If e.ColumnIndex <> Me.ColDept.Index Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                                            Me.objCd.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString)
                            End If

                        End If
                    End If
                Else
                    If Me.dgvItem.Rows.Count - 1 = e.RowIndex Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If e.ColumnIndex = Me.ColRemarks.Index Then
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Else
                'col account
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Rounding As Integer = 0
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim ExRate As Decimal = 1
        Dim DefPrecOriginalCurr As Integer = 0
        Try
            Rounding = Me.DefaultPrecision
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'Remarks
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColAccount.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.EraseRowIdx = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColJob.Index Then
                    'hapuskan kode job
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = -1
                ElseIf e.ColumnIndex = Me.ColDept.Index Then
                    'hapuskan kode dept
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = -1
                Else
                    'untuk kolom debet kredit, langsung di balikin ke nilai semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            Else
                ''cek debet
                If e.ColumnIndex = Me.ColDebet.Index Then

                    If Me.objCd.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = objCd.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec_BaseCurr)
                    ''Cek kredit
                    If Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value = objCd.DigitGrouping(0, DefPrec_BaseCurr)
                    End If
                    Me.BalanceDebetKredit(e.RowIndex, True, Me.DefPrec_BaseCurr)
                    Me.HitungOriginalValues(e.RowIndex)

                    If e.RowIndex = 0 Then
                        If Me.TxtPrimaryCurrency.Text = Me.PrimaryBaseCurr Then
                            Me.TxtBalance.Text = Me.objCd.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value) / CDec(CDec(Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value)), _
                            Me.DefPrec_BaseCurr, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        Else
                            Me.TxtBalance.Text = Me.objCd.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value) / CDec(CDec(Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value)), _
                                                  Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        End If
                    End If
                End If

                ''Cek kredit 
                If e.ColumnIndex = Me.ColKredit.Index Then
                    If Me.objCd.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                            e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = objCd.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrec_BaseCurr)
                    ''Cek debet
                    If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value <> 0 Then
                        Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value = objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                    End If
                    Me.BalanceDebetKredit(e.RowIndex, False, Me.DefPrec_BaseCurr)
                    Me.HitungOriginalValues(e.RowIndex)
                End If

                ''data account number
                If e.ColumnIndex = Me.ColAccount.Index Then
                    If Me.Buffer = Me.dgvItem.Item(Me.ColAccount.Index, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    If e.ColumnIndex <> Me.ColDebet.Index Then
                        If e.ColumnIndex <> Me.ColKredit.Index Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            'kemungkinan
                            If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                                Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.objAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.objAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.objAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColOriginalCurrency.Index, e.RowIndex).Value = Me.objCurr.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value = Me.objCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                                Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                                Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColCredit_OriginalCurrency.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value)

                                ExRate = Me.ObjTaxRate.GetTaxRate(Me.DtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                If ExRate = 0 Then
                                    ExRate = 1
                                End If
                                If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurr Then
                                    DefPrecEx = 0
                                Else
                                    DefPrecEx = Core_Module.DefPrecExchangeRate
                                End If
                                Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = objCd.DigitGrouping(ExRate, DefPrecEx)
                                'DefPrecOriginalCurr = Me.ObjCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                Me.HitungOriginalValues(e.RowIndex)
                            Else
                                '2. kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New FrmMasterAkun_Cari
                                ij.BackColorAll = Me.BackColorAll
                                ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.ShowDialog()
                                If ij.mPrimaryAkun <> -1 Then
                                    '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.objAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.objAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.objAkun.GetPrimaryCurrency(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColOriginalCurrency.Index, e.RowIndex).Value = Me.objCurr.GetCurrencyCode(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value = Me.objCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                                    Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.DefPrec_BaseCurr)
                                    Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColCredit_OriginalCurrency.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(0, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value)
                                    ExRate = Me.ObjTaxRate.GetTaxRate(Me.DtpDate.Value, Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    If ExRate = 0 Then
                                        ExRate = 1
                                    End If

                                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurr Then
                                        DefPrecEx = 0
                                    Else
                                        DefPrecEx = Core_Module.DefPrecExchangeRate
                                    End If
                                    Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value = objCd.DigitGrouping(ExRate, DefPrecEx)
                                    'DefPrecOriginalCurr = Me.ObjCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                    Me.HitungOriginalValues(e.RowIndex)
                                ElseIf ij.mPrimaryAkun = -1 Then
                                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        'Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                        Me.EraseRowIdx = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                    End If
                    Me.SetMulticurrency()
                End If

                '3. exchange rate, hanya jika non base currency
                If e.ColumnIndex = Me.ColExchangeRate.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurr Then
                        If Me.objCd.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value = Me.PrimaryBaseCurr Then
                            DefPrecEx = 0
                        Else
                            DefPrecEx = Core_Module.DefPrecExchangeRate
                        End If
                        'digit groping
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        objCd.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DefPrecEx)
                        'hitung original value
                        Me.HitungOriginalValues(e.RowIndex)
                        Me.HitungBaseValues(e.RowIndex)
                        If e.RowIndex = 0 Then
                            Me.TxtExchangeRate.Text = Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value
                            Me.UpdateBalance()
                        Else
                            Me.dgvItem.Item(Me.ColCredit_OriginalCurrency.Index, e.RowIndex).Value = Me.objCd.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColKredit.Index, e.RowIndex).Value) / _
                            CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        End If
                    End If
                End If

                '3. debet original currency
                If e.ColumnIndex = Me.ColDebet_OriginalCurr.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurr Then
                        If Me.objCd.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.objCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        objCd.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvItem.Item(Me.ColCredit_OriginalCurrency.Index, e.RowIndex).Value = objCd.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, True)
                        End If

                        If e.RowIndex = 0 Then
                            If Me.TxtPrimaryCurrency.Text = Me.PrimaryBaseCurr Then
                                Me.TxtBalance.Text = Me.objCd.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value) / CDec(CDec(Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value)), _
                                Me.PrimaryBaseCurr, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                Me.TxtBalance.Text = Me.objCd.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value) / CDec(CDec(Me.dgvItem.Item(Me.ColExchangeRate.Index, e.RowIndex).Value)), _
                                                            Me.DefaultPrecision, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    End If
                End If

                '4. credit original currency
                If e.ColumnIndex = Me.ColCredit_OriginalCurrency.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value <> Me.PrimaryBaseCurr Then
                        If Me.objCd.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, _
                                                   e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'digit groping
                        DefPrecOriginalCurr = Me.objCurr.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        objCd.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                        DefPrecOriginalCurr)
                        'kosongkan credit
                        Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, e.RowIndex).Value = objCd.DigitGrouping(0, DefPrecOriginalCurr)

                        Me.HitungBaseValues(e.RowIndex)

                        If Me.dgvItem.Item(Me.ColDebet.Index, e.RowIndex).Value > 0 Then
                            Me.BalanceDebetKredit(e.RowIndex, True, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, True)
                        Else
                            Me.BalanceDebetKredit(e.RowIndex, False, Me.dgvItem.Item(Me.ColDefPrec.Index, e.RowIndex).Value, True)
                        End If
                    End If
                End If

                '5. Dept
                If e.ColumnIndex = Me.ColDept.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjDept.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = Me.ObjDept.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterDepartementCari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryDepartement <> -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                            Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement
                        ElseIf ij.PrimaryDepartement = -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvItem.Item(Me.ColPrimaryDept.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If

                '6. Job
                If e.ColumnIndex = Me.ColJob.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.ObjJob.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = Me.ObjJob.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        Dim ij As New Core_Module.FrmMasterJobCari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryJob <> -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjJob.GetJobCode(ij.PrimaryJob)
                            Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = ij.PrimaryJob
                        ElseIf ij.PrimaryJob = -1 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvItem.Item(Me.ColPrimaryJob.Index, e.RowIndex).Value = -1
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If

                If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value = Convert.ToString(Me.dgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value).Trim
                End If

                If e.ColumnIndex = Me.ColRemarks.Index Then
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("%", "").ToString.Replace("'", "").ToString.Replace("*", "").ToString.Replace("^", "").ToString.Replace("#", "").ToString.Replace("&", "").ToString
                        End If
                    End If
                End If
                '3 cek apakah ada kode akun yang sama pada referensi kode yang dibuat 
                'If Me.dgvItem.Rows.Count > 2 Then
                'Me.dgvItem.AllowUserToAddRows = False
                'Dim item As String = ""
                'For i As Integer = 0 To dgvItem.Rows.Count - 1
                'If e.RowIndex <> i Then
                'item = CStr(Me.dgvItem.Rows(i).Cells(Me.ColAccount.Index).Value)
                'If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = item Then
                'MsgBox("Duplication account number at data item, please change your number", _
                'MsgBoxStyle.Information, "Information")
                'Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                'Me.dgvItem.AllowUserToAddRows = True
                'Exit Sub
                'End If
                'End If
                'Next
                'Me.dgvItem.AllowUserToAddRows = True
                'End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ''hitung debet dan kredit
            Me.HitungUnbalance()
            ObjCurr = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.EraseRowIdx > 0 Then
                Me.dgvItem.Rows.RemoveAt(Me.EraseRowIdx)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.dgvItem.AllowUserToDeleteRows = True
            ''hitung debet dan kredit
            Dim row As DataGridViewRow
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                TotalDebet += CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value)
            Next
            Me.TxtUnbalance.Text = CStr(objCd.DigitGrouping(Me.HitungUnbalance(TotalDebet, TotalKredit), Me.DefaultPrecision))
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            Me.dgvItem.TabStop = True
            If Me.txtRemarks.Enabled = True And Me.dgvItem.AllowUserToAddRows = True Then
                'set nilai default baris baru
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.PrimaryBaseCurr
                DefPrecBase = Me.objCurr.GetDefaultPrecission(Me.PrimaryBaseCurr)
                rowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = Me.objCd.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = Me.objCd.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(0, DefPrecBase)
                rowDGV.Cells(Me.ColExchangeRate.Index).Value = 1
                rowDGV.Cells(Me.ColOriginalCurrency.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.ColAccount.Index).Value = ""
                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                rowDGV.Cells(Me.ColDept.Index).Value = ""
                rowDGV.Cells(Me.ColJob.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaryDept.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryJob.Index).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BalanceDebetKredit(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, _
 Optional ByVal vHitungOriginalCurrency As Boolean = True)

        Try
            'Dim rowAcuan As DataGridViewRow = Me.dgvItem.Rows(vIndexAcuan)
            'Dim rowTarget As DataGridViewRow

            'If Me.txtRemarks.Enabled = True And Me.dgvItem.Rows.Count = 3 Then
            '    'dibalance jika hanya ada 2 jurnal, maka debet = kredit
            '    If rowAcuan.Index = 0 Then
            '        rowTarget = Me.dgvItem.Rows(1)
            '    Else
            '        rowTarget = Me.dgvItem.Rows(0)
            '    End If

            '    If vAcuanDebet = True Then
            '        rowTarget.Cells(Me.ColKredit.Index).Value = rowAcuan.Cells(Me.ColDebet.Index).Value
            '        rowTarget.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(0, 0)

            '        rowTarget.Cells(Me.ColCreditNR.Index).Value = rowAcuan.Cells(Me.ColDebetNR.Index).Value
            '    Else
            '        rowTarget.Cells(Me.ColDebet.Index).Value = rowAcuan.Cells(Me.ColKredit.Index).Value
            '        rowTarget.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(0, 0)

            '        rowTarget.Cells(Me.ColDebetNR.Index).Value = rowAcuan.Cells(Me.ColCreditNR.Index).Value
            '    End If

            '    If vHitungOriginalCurrency = True Then
            '        Me.HitungOriginalValues(rowTarget.Index)
            '    End If
            'ElseIf Me.txtRemarks.Enabled = True And Me.dgvItem.Rows.Count > 3 Then
            '    If rowAcuan.Index = 0 Then
            '        rowTarget = Me.dgvItem.Rows(1)
            '    Else
            '        rowTarget = Me.dgvItem.Rows(0)
            '    End If

            '    'dibalance jika ada 3 jurnal, jika 2 dikredit atau 2 didebet
            '    If Me.GetJmlNol = 0 Then
            '        If vAcuanDebet = True Then
            '            If Me.GetJmlKredit = 1 Then
            '                rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlKredit)
            '                rowTarget.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(Me.GetTotalDebet, 0)
            '                rowTarget.Cells(Me.ColCreditNR.Index).Value = Me.GetTotalDebetNR
            '                rowTarget.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(0, 0)
            '            End If
            '        Else
            '            If Me.GetJmlDebet = 1 Then
            '                rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlDebet)
            '                rowTarget.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(Me.GetTotalKredit, 0)
            '                rowTarget.Cells(Me.ColDebetNR.Index).Value = Me.GetTotalKreditNR
            '                rowTarget.Cells(Me.ColKredit.Index).Value = Me.objCd.DigitGrouping(0, 0)
            '            End If
            '        End If

            '        If vHitungOriginalCurrency = True Then
            '            If IsNothing(rowTarget) = False Then
            '                Me.HitungOriginalValues(rowTarget.Index)
            '            End If
            '        End If
            '    End If
            'End If

            'If vIndexAcuan <> 0 Then

            '    Me.TxtBalance.Text = Me.dgvItem.Item(Me.ColDebet_OriginalCurr.Index, 0).Value
            '    'Me.UpdateBalance()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungBaseValues(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim DefPrecOriginal As Integer = 0

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            DefPrecBase = Me.objCurr.GetDefaultPrecission(Me.PrimaryBaseCurr)
            DefPrecOriginal = Me.objCurr.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColKredit.Index).Value = objCd.DigitGrouping(0, DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = Me.objCd.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColKredit.Index).Value = objCd.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = objCd.DigitGrouping(0, DefPrecBase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexJmlKredit() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count > 1 Then

                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2

                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        Return row.Index
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJmlDebet() As Integer
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        Return row.Index
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetTotalKreditNR() As Decimal
    '    Dim row As DataGridViewRow
    '    Dim retval As Decimal = 0

    '    Try
    '        If Me.dgvItem.Rows.Count > 1 Then
    '            For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
    '                row = Me.dgvItem.Rows(i)
    '                If row.Cells(Me.ColCreditNR.Index).Value <> 0 Then
    '                    retval += CDec(row.Cells(Me.ColCreditNR.Index).Value)
    '                End If
    '            Next
    '        End If

    '        Return retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function GetJmlNol() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value = 0 And row.Cells(Me.ColKredit.Index).Value = 0 Then
                        retval += 1
                    End If
                Next
            End If
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalKredit() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColKredit.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTotalDebet() As Decimal
        Dim row As DataGridViewRow
        Dim retval As Decimal = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        retval += CDec(row.Cells(Me.ColDebet.Index).Value)
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlKredit() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColKredit.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetTotalDebetNR() As Decimal
    '    Dim row As DataGridViewRow
    '    Dim retval As Decimal = 0

    '    Try
    '        If Me.dgvItem.Rows.Count > 1 Then
    '            For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
    '                row = Me.dgvItem.Rows(i)
    '                If row.Cells(Me.ColDebetNR.Index).Value <> 0 Then
    '                    retval += CDec(row.Cells(Me.ColDebetNR.Index).Value)
    '                End If
    '            Next
    '        End If

    '        Return retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function GetJmlDebet() As Integer
        Dim row As DataGridViewRow
        Dim retval As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    row = Me.dgvItem.Rows(i)
                    If row.Cells(Me.ColDebet.Index).Value <> 0 Then
                        retval += 1
                    End If
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    If Me.dgvItem.SelectedCells(0).RowIndex = 0 Then
                        Me.dgvItem.AllowUserToDeleteRows = False
                    Else
                        Me.dgvItem.AllowUserToDeleteRows = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBalance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBalance.Leave
        DefaultPrecision = objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Try
            If IsNumeric(Me.TxtBalance.Text) Then
                objCd.Number_Leave(TxtBalance, sender, e, Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), objCd)
                Me.UpdateBalance()
            Else
                TxtBalance.Text = Me.objCd.DigitGrouping(0, Me.objCurr.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                Me.UpdateBalance()
            End If
            Me.HitungUnbalance()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class