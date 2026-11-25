Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiJournal_CashBank
    Public UserCOde As String = ""
    Public BackcolorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
    Public TipeJurnall2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank

    Dim Buffer As String = ""
    Dim DefaultPrecision As Integer = 0
    Dim Baru As Boolean
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjCD As New ClsCekData
    Dim ObjQuickJurnal As New ClsTransaksiJournalQuickAdd
    Dim ObjLicense As New ClsLicense
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim PrimaryBaseCurrency As Double = 1
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDefaultValue As New Core_Module.ClsTransaksiSetting
    Dim ObjDept As New ClsMasterDepartement
    Dim ObjJob As New Core_Module.ClsMasterJob
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim objcek As New ClsTransaksiJournal_Item_Cheque
    Dim FrmCek As New FrmTransaksiJournal_CashBank_Cheque
    Dim SetDefaultValue As Boolean = True
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Dim TBLAkun As New DseMasterAkun.tblmaster_akun4DataTable
    Dim IndexRemove As Integer = -1

    '===========================================================================================================================================================
    'Erna's code mulai
    '===========================================================================================================================================================
    Private Sub dgvPayment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPayment.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvPayment.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryAcc As Double

        Try
            With row
                Dim DefPrec As Integer = Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value))
                Dim DefPrecIDR As Decimal = Me.ObjCur.GetDecimal_DefaultCurrency
                Me.ColItemPay_Amount.Visible = True

                If IsNothing(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan
                    If Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    '==AKUN
                    If e.ColumnIndex = Me.ColAccount.Index Then
                        If Me.cboAccount.Text = "" Then
                            Exit Sub
                        End If

                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryAcc = Me.ObjAkun.GetPrimary(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = -1 'Me.GetIndexItemInput(PrimaryAcc)
                            If index < 0 Then
                                Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvPayment.Item(Me.COlItemPay_Account.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(PrimaryAcc)
                                Me.dgvPayment.Item(Me.COlItemPay_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAcc)
                                Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.ObjAkun.GetPrimaryCurrency(PrimaryAcc)), Core_Module.DefPrecExchangeRate)
                                Me.dgvPayment.Item(Me.ColItemPay_Curr.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(PrimaryAcc))
                                Me.dgvPayment.Item(Me.COlItemPay_AmountOriginal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, _
                                 Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(PrimaryAcc)))
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in account code, please change your account code", MsgBoxStyle.Information, "Information")
                                    Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.IndexRemove = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New FrmMasterAkun_Cari
                            ij.txtCari.Text = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexItemInput(ij.mPrimaryAkun)
                            If ij.mPrimaryAkun = -1 Then
                                ''lewat
                                If Me.Buffer = "" Then
                                    Me.IndexRemove = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Exit Sub
                                Else
                                    Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            Else
                                If index <= 0 Then
                                    Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun.ToString
                                    Me.dgvPayment.Item(Me.COlItemPay_Account.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.dgvPayment.Item(Me.COlItemPay_Desc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                    Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.ObjAkun.GetPrimaryCurrency(ij.mPrimaryAkun)), Core_Module.DefPrecExchangeRate)
                                    Me.dgvPayment.Item(Me.ColItemPay_Curr.Index, e.RowIndex).Value = Me.ObjCur.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(ij.mPrimaryAkun))
                                    Me.dgvPayment.Item(Me.COlItemPay_AmountOriginal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, _
                                     Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(ij.mPrimaryAkun)))
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in account code, please change your account code", MsgBoxStyle.Information, "Information")
                                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        If Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.IndexRemove = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                    End If
                    ''==AMOUNT Ori
                    If e.ColumnIndex = Me.COlItemPay_AmountOriginal.Index Then
                        If Me.ObjCD.CekNumerik(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Please enter a number for amount.", MsgBoxStyle.Information, "Information")
                            Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        Else
                            Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                            Me.dgvPayment.Item(Me.ColItemPay_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping _
                            (Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value * Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value, _
                            DefPrecIDR)
                        End If
                    End If

                    ''==Exchange Rate
                    If e.ColumnIndex = Me.ColItemPay_ExRate.Index Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Please enter a positif number for exchange rate.", MsgBoxStyle.Information, "Information")
                            Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.DefPrecExchangeRate)
                        Me.dgvPayment.Item(Me.ColItemPay_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping _
                            (Me.dgvPayment.Item(Me.COlItemPay_AmountOriginal.Index, e.RowIndex).Value * Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value, _
                            DefPrecIDR)
                    End If
                    ''==Departemen
                    If e.ColumnIndex = Me.COlItemPay_DeptCode.Index Then
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjDept.ApakahSudahAda(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvPayment.Item(Me.COlItemPay_PrimaryDept.Index, e.RowIndex).Value = Me.ObjDept.GetPrimary(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            Dim ij As New FrmMasterDepartementCari
                            ij.txtCari.Text = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.PrimaryDepartement <> -1 Then
                                Me.dgvPayment.Item(Me.COlItemPay_PrimaryDept.Index, e.RowIndex).Value = ij.PrimaryDepartement.ToString
                                Me.dgvPayment.Item(Me.COlItemPay_DeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                            Else
                                Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvPayment.Item(Me.COlItemPay_PrimaryDept.Index, e.RowIndex).Value = -1
                            End If
                        End If
                    End If

                    ''==Job
                    If e.ColumnIndex = Me.COlItemPay_JobCode.Index Then
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'buat sementara tidak ada batasan
                        If Me.ObjJob.ApakahSudahAda(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvPayment.Item(Me.COlItemPay_PrimaryJob.Index, e.RowIndex).Value = Me.ObjJob.GetPrimary(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            Dim ij As New FrmMasterJobCari
                            ij.txtCari.Text = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.PrimaryJob <> -1 Then
                                Me.dgvPayment.Item(Me.COlItemPay_PrimaryJob.Index, e.RowIndex).Value = ij.PrimaryJob.ToString
                                Me.dgvPayment.Item(Me.COlItemPay_JobCode.Index, e.RowIndex).Value = Me.ObjJob.GetJobCode(ij.PrimaryJob)
                            Else
                                Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.dgvPayment.Item(Me.COlItemPay_PrimaryJob.Index, e.RowIndex).Value = -1
                            End If
                        End If
                    End If
                Else
                    'klo nothing
                    If e.ColumnIndex = Me.ColAccount.Index Then
                        If Me.cboAccount.Text = "" Then
                            Exit Sub
                        End If
                        If Me.Buffer = "" Then
                            'remove row
                            Me.IndexRemove = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                            Exit Sub
                        Else
                            'balikin ke buffer
                            Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.COlItemPay_AmountOriginal.Index Then
                        ''==AMOUNT Ori, jika dimasukkan nothing --> set nilainya = 0
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(0, DefPrec)
                        Me.dgvPayment.Item(Me.ColItemPay_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping _
                        (Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value * Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value, _
                        DefPrecIDR)
                    ElseIf e.ColumnIndex = Me.ColItemPay_ExRate.Index Then
                        ''==Exchange Rate, jika dimasukkan nothing --> set nilai = 0
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.ObjAkun.GetPrimaryCurrency(Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value)), DefPrec)
                        Me.dgvPayment.Item(Me.ColItemPay_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping _
                        (Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value * Me.dgvPayment.Item(Me.ColItemPay_ExRate.Index, e.RowIndex).Value, _
                        DefPrecIDR)
                    ElseIf e.ColumnIndex = Me.COlItemPay_DeptCode.Index Then
                        ''==Departemen, jika dimasukkan nothing, maka primary dept=-1, dept code =""
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.dgvPayment.Item(Me.COlItemPay_PrimaryDept.Index, e.RowIndex).Value = -1
                    ElseIf e.ColumnIndex = Me.COlItemPay_JobCode.Index Then
                        ''==Job, jika dimasukkan nothing, maka primary job =-1, job code =""
                        Me.dgvPayment.Item(Me.COlItemPay_PrimaryJob.Index, e.RowIndex).Value = -1
                        Me.dgvPayment.Item(Me.COlItemPay_JobCode.Index, e.RowIndex).Value = ""
                    End If
                End If
            End With

            Me.TransferKeAccountDetail()
            Me.SetTampilanDGVPayment()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub TransferKeAccountDetail()
        Dim i As Integer = 1
        Dim RowPayment As DataGridViewRow
        Dim DefPrecIDR As Decimal = 0
        Try
            'bersihkan dgv account
            Me.dgvItem.Rows.Clear()
            If Me.dgvPayment.AllowUserToAddRows = True Then
                i = 1
            Else
                i = 0
            End If
            DefPrecIDR = Me.ObjCur.GetDecimal_DefaultCurrency
            For Each RowPayment In Me.dgvPayment.Rows
                If i <> Me.dgvPayment.Rows.Count Then
                    Dim _Debit, _Kredit, _DebitOri, _KreditOri As Decimal
                    Dim DefPrecHeadrer As Integer = Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text))
                    Dim Debit, Kredit, DebitOri, KreditOri As Decimal
                    Dim DefPrecPayment As Integer = Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(RowPayment.Cells(Me.COlItemPay_PrimaryAkun.Index).Value))
                    'cek Receive to atau pay from
                    If Me.cboType.SelectedIndex = 0 Then
                        'receive to
                        If RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value > 0 Then
                            'payment
                            DebitOri = 0
                            KreditOri = RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value
                            Debit = 0
                            Kredit = KreditOri * RowPayment.Cells(Me.ColItemPay_ExRate.Index).Value
                            'header
                            _Debit = Kredit
                            _Kredit = 0
                            _DebitOri = _Debit / Me.TxtExchangeRate.Text
                            _KreditOri = 0
                        Else
                            'jika hasilnya minus , balikan nilai amountnya
                            'payment
                            DebitOri = RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value * -1
                            KreditOri = 0
                            Debit = DebitOri * RowPayment.Cells(Me.ColItemPay_ExRate.Index).Value
                            Kredit = 0
                            'header
                            _Debit = 0
                            _Kredit = Debit
                            _DebitOri = 0
                            _KreditOri = _Kredit / Me.TxtExchangeRate.Text
                        End If
                    Else
                        'pay from
                        If RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value > 0 Then
                            'payment
                            DebitOri = RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value
                            KreditOri = 0
                            Debit = DebitOri * RowPayment.Cells(Me.ColItemPay_ExRate.Index).Value
                            Kredit = 0
                            'header
                            _Debit = 0
                            _Kredit = Debit
                            _DebitOri = 0
                            _KreditOri = _Kredit / Me.TxtExchangeRate.Text
                        Else
                            'jika hasilnya minus , balikan nilai amountnya
                            'payment
                            DebitOri = 0
                            KreditOri = RowPayment.Cells(Me.COlItemPay_AmountOriginal.Index).Value * -1
                            Debit = 0
                            Kredit = KreditOri * RowPayment.Cells(Me.ColItemPay_ExRate.Index).Value
                            'header
                            _Debit = Kredit
                            _Kredit = 0
                            _DebitOri = _Debit / Me.TxtExchangeRate.Text
                            _KreditOri = 0
                        End If
                    End If

                    'header
                    Me.dgvItem.Rows.Add(Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAccount.Text), Me.cboAccount.Text, Me.ObjCD.DigitGrouping(_Debit, DefPrecIDR), _
                    Me.ObjCD.DigitGrouping(_Kredit, DefPrecIDR), Me.TxtExchangeRate.Text, Me.LblCurrency.Text, Me.ObjCD.DigitGrouping(_DebitOri, DefPrecHeadrer), _
                    Me.ObjCD.DigitGrouping(_KreditOri, DefPrecHeadrer), RowPayment.Cells(Me.COlItemPay_Remarks.Index).Value, "", "", -1, -1, Me.TxtPrimaryAccount.Text, Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text))
                    'payment
                    Me.ColJob.Visible = True
                    Me.dgvItem.Rows.Add(RowPayment.Cells(Me.COlItemPay_Account.Index).Value, RowPayment.Cells(Me.COlItemPay_Desc.Index).Value, Me.ObjCD.DigitGrouping(Debit, DefPrecIDR), _
                    Me.ObjCD.DigitGrouping(Kredit, DefPrecIDR), RowPayment.Cells(Me.ColItemPay_ExRate.Index).Value, Me.ObjCur.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(RowPayment.Cells(Me.COlItemPay_PrimaryAkun.Index).Value)), _
                    Me.ObjCD.DigitGrouping(DebitOri, DefPrecPayment), _
                    Me.ObjCD.DigitGrouping(KreditOri, DefPrecPayment), RowPayment.Cells(Me.COlItemPay_Remarks.Index).Value, RowPayment.Cells(Me.COlItemPay_DeptCode.Index).Value, _
                    RowPayment.Cells(Me.COlItemPay_JobCode.Index).Value, RowPayment.Cells(Me.COlItemPay_PrimaryDept.Index).Value, RowPayment.Cells(Me.COlItemPay_PrimaryJob.Index).Value, _
                    RowPayment.Cells(Me.COlItemPay_PrimaryAkun.Index).Value, Me.ObjAkun.GetPrimaryCurrency(RowPayment.Cells(Me.COlItemPay_PrimaryAkun.Index).Value))

                    i = i + 1
                End If
            Next
            Me.formatDgvAccount()
            Me.HitungAmount()
            Me.SetMulticurrency()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '===========================================================================================================================================================
    'Erna's Code selesai()
    '===========================================================================================================================================================

    Public Sub ShowJurnal(ByVal vPrimaryJurnal As Double)
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemarks.Enabled = False Then
                If Me.ObjJurnal.CariDetail(vPrimaryJurnal, Me.dgvItem, Me, Me.TipeJournal, Me.dgvPayment) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''Me.cboJournalTipe.SelectedIndex = Me.TipeJournal
        Dim NullValue As String = ""

        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.ObjJurnal = New ClsTransaksiJournal
            Me.Baru = True
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvPayment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll
            Me.dgvItem.Rows.Clear()
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal Then
                Me.Text = "Adjustment Journal"
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                Me.Text = "Closing Journal"
                'persiapkan untuk closing 
                Me.dtpDate.CustomFormat = "yyyy"
                'jika sudah ada jurnal closing sebelumnya, maka tanggal closing harus setahun setelahnya
                If Me.ObjJurnal.GetJml(Me.TipeJournal) > 0 Then
                    Me.dtpDate.Enabled = False
                End If
            ElseIf Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal Then
                Me.Text = "General Journal"
            End If
            Me.Text = "Cash And Bank"

            ' Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackcolorAll

            'set null value untuk dgv
            Me.DefaultPrecision = Me.ObjCur.GetDecimal_DefaultCurrency
            NullValue = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
            Me.ColDebet.DefaultCellStyle.NullValue = NullValue
            Me.ColKredit.DefaultCellStyle.NullValue = NullValue

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            Me.ColDept.Visible = False
            'End If

            'jika tidak ada hak akses atas job, maka kolom job ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            Me.ColJob.Visible = False
            'End If

            Me.cboType.SelectedIndex = 0
            Me.FrmCek.dgvCheque.Rows.Clear()
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnViewJurnal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewJurnal.Click
        Dim FrmParent As Windows.Forms.Form = Me.MdiParent

        Try
            Dim ij As FrmTransaksiJournal_View
            Dim j As Windows.Forms.Form

            For Each j In FrmParent.MdiChildren
                ij = TryCast(j, FrmTransaksiJournal_View)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmTransaksiJournal_View
            i.UserCode = Me.UserCOde
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = FrmParent
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Pengaturan DGV Item
    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvItemProcess_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            ''hitung debet dan kredit
            Dim row As DataGridViewRow
            Dim TotalDebet As Decimal = 0
            Dim TotalKredit As Decimal = 0
            For Each row In dgvItem.Rows
                TotalDebet += CDec(Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value)
                TotalKredit += CDec(Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value)
            Next
            Me.SetMulticurrency()
            'If Me.TxtRemarks.Enabled = True Then
            'If Me.dgvItem.RowCount = 0 Then
            ''
            'End If
            'End If
            ''Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvPayment.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvPayment.FirstDisplayedScrollingRowIndex = Me.dgvPayment.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvPayment.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvPayment.Rows(row.Index).Displayed = False
                    Me.dgvPayment.FirstDisplayedScrollingRowIndex = Me.dgvPayment.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvPayment.DisplayedRowCount(True) <> Me.dgvPayment.DisplayedRowCount(False) Then
                    Me.dgvPayment.FirstDisplayedScrollingRowIndex = Me.dgvPayment.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.dgvPayment.AllowUserToAddRows = False
            Me.SetDefaultValue = False

            If Me.dgvPayment.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvPayment.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvPayment.Rows.Remove(row)
                    Me.dgvPayment.Rows.InsertRange(indexbaru, row)
                    Me.dgvPayment.Focus()
                    Me.dgvPayment.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvPayment)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvPayment.AllowUserToAddRows = True
            Me.SetDefaultValue = True
            Me.TransferKeAccountDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvPayment.AllowUserToAddRows = False
            Me.SetDefaultValue = False

            If Me.dgvPayment.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvPayment.SelectedRows(0)
                If row.Index <> Me.dgvPayment.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvPayment.Rows.Remove(row)
                    Me.dgvPayment.Rows.InsertRange(indexbaru, row)
                    Me.dgvPayment.Focus()
                    Me.dgvPayment.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvPayment)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvPayment.AllowUserToAddRows = True
            Me.SetDefaultValue = True
            Me.TransferKeAccountDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPayment_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvPayment.RowsRemoved
        Try
            If Me.dgvPayment.Rows.Count = 0 Then
                Me.dgvPayment.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProcess_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPayment.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvPayment.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmTransaksiJournal_Cetak
        ij.TipeJournal = Me.TipeJournal
        ij.TipeJurnal2 = Me.TipeJurnall2

        Try
            ij.BackColorAll = Me.BackcolorAll
            ij.ShowDialog()
            Me.dgvItem.AllowUserToAddRows = False
            If Not ij.PrimaryPP = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive) = "1" Then
                    If Me.ObjJurnal.ApakahMemilikiLedgerAccess(Me.UserCOde, Me.ObjJurnal.GetCreatedBy(ij.PrimaryPP)) = False Then
                        MsgBox("You are not authorized to view this data.", MsgBoxStyle.Information, "Information")
                        Me.KosongkanField()
                        Exit Sub
                    End If
                End If

                If Me.ObjJurnal.CariDetail(ij.PrimaryPP, Me.dgvItem, Me, ij.TipeJournal, Me.dgvPayment) = True Then
                    Me.SalinkeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    ''fungsi standar 
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryJournal.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryJournal.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvPayment.AllowUserToDeleteRows = True
            Me.dgvPayment.AllowUserToAddRows = True
            Me.dgvPayment.ReadOnly = False
            Me.dgvItem.ReadOnly = False
            Me.btnUP.Enabled = False
            Me.btnDown.Enabled = False

            Me.SetTxtExchangeRateHeader()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, True, False, True, False)
            Me.CancelButton = Me.btnExit

            'Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True

            Me.dgvPayment.AllowUserToDeleteRows = False
            Me.dgvPayment.AllowUserToAddRows = False
            Me.dgvPayment.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, True)

            Me.SetTampilanDGVPayment()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryJournal.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "


            Me.TxtAmount.Text = 0
            Me.TxtAmountOriginal.Text = 0
            Me.TxtPrimaryAccount.Text = -1
            Me.cboAccount.DataSource = Nothing
            Me.dgvItem.Rows.Clear()
            Me.dgvPayment.Rows.Clear()
            ''Me.ObjCD.SetNumber(Me.dgvItem)
            Me.FrmCek.dgvCheque.Rows.Clear()

            Me.SetTampilanDGVPayment()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeObjek()
        Try
            With ObjJurnal
                .RefNumber = Me.TxtRefNo.Text
                .Description = Me.TxtRemarks.Text
                .Tanggal = Me.dtpDate.Value
                .PrimaryAkunPayTo = Me.TxtPrimaryAccount.Text
                .ExchangeRate = Me.TxtExchangeRate.Text
                .PrimaryCurrency = Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text)
                If Me.cboType.SelectedItem = "Receive To" Then
                    .TipeInOut = ClsTransaksiJournal.EnumTipeInOut._In
                ElseIf Me.cboType.SelectedItem = "Pay From" Then
                    .TipeInOut = ClsTransaksiJournal.EnumTipeInOut.out
                End If
                .PrimaryCurrency = Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text)
                .ReferenceParty = Me.TxtReferenceParty.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinkeField()
        Dim ObjCheque As New ClsTransaksiJournal_Item_Cheque
        Try
            Me.ObjAkun.IsiCboCashAndBank(Me.cboAccount, Me.TBLAkun)
            With ObjJurnal
                Me.TxtPrimaryJournal.Text = .Primary
                Me.TxtRefNo.Text = .RefNumber
                Me.TxtRemarks.Text = .Description
                Me.dtpDate.Value = .Tanggal
                Me.TxtPrimaryAccount.Text = .PrimaryAkunPayTo
                Me.cboAccount.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccount.Text).ToString
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.ExchangeRate, Core_Module.DefPrecExchangeRate)
                '========================================================================================================
                'tambahan erna
                Me.LblCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text))
                '========================================================================================================
                If .TipeInOut = ClsTransaksiJournal.EnumTipeInOut._In Then
                    Me.cboType.SelectedIndex = 0
                Else
                    Me.cboType.SelectedIndex = 1
                End If
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TxtReferenceParty.Text = .ReferenceParty
                Me.HitungAmount()

                Me.SetMulticurrency()
                Me.SetTampilanDGVPayment()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCheque = Nothing
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim TBL As New DseTransaksiJournal.tbltrans_jurnal_chequeDataTable
        Dim rowcek As DseTransaksiJournal.tbltrans_jurnal_chequeRow

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Reference number")
            KoleksiTextBox.Add(Me.TxtRefNo)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtRemarks)
            If Me.cboType.SelectedIndex = ClsTransaksiJournal.EnumTipeInOut._In Then
                KoleksiNamaData.Add("Received From")
            Else
                KoleksiNamaData.Add("Payment To")
            End If
            KoleksiTextBox.Add(Me.TxtReferenceParty)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'Cek DGVItem
            If Me.dgvItem.Rows.Count < 2 Then
                MsgBox("Minimum item = 2, please input item again", MsgBoxStyle.Information, "Information")
                Me.dgvItem.Focus()
                Return False
            End If

            'tidak boleh bikin transaksi di belakang convertion date
            If Me.ObjClosing.GetConvertionDate.Date > Me.dtpDate.Value.Date Then
                MsgBox("Minimum transaction date is " + Format(Me.ObjClosing.GetConvertionDate.Date, "dd MMM yyyy") + ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah ada account number yang debet dan kreditnya bernilai 0
            'For Each row In Me.dgvItem.Rows
            '    If Me.dgvItem.Item(Me.ColDebet.Index, row.Index).Value = 0 And _
            '    Me.dgvItem.Item(Me.ColKredit.Index, row.Index).Value = 0 Then
            '        MsgBox("Please add positive number in account no. " + Me.dgvItem.Item(Me.ColAccount.Index, row.Index).Value, _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next

            'cek perubahan nomor transaksi
            'If Me.Baru = False Then
            '    If Me.ObjPP.TanggalTransaksi.Month = Me.dtpTanggalTransaksi.Value.Month And _
            '    Me.ObjPP.TanggalTransaksi.Year = Me.dtpTanggalTransaksi.Value.Year Then
            '        Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            '    Else
            '        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
            '        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            '            Me.dtpTanggalTransaksi.Value = Me.ObjPP.TanggalTransaksi
            '            Me.txtNoTransaksi.Text = Me.ObjPP.NoTransaksi
            '            Return False
            '        End If
            '    End If
            'End If


            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                    MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtRefNo.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.ObjJurnal.GetRefCode(Me.TxtPrimaryJournal.Text)
                If TipeCodeLama <> Me.TxtRefNo.Text Then
                    If Me.ObjJurnal.ApakahSudahAda(Me.TxtRefNo.Text) = True Then
                        MsgBox("Duplication in no journal =" + Me.TxtRefNo.Text + ", please change your number", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtRefNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.ObjJurnal.GetTimestamp(Me.TxtPrimaryJournal.Text)
                If Me.ObjJurnal.Timestamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.NonAktifkanDetail()
                        Me.ObjJurnal.CariDetail(Me.ObjJurnal.Primary, Me.dgvItem, Me, Me.TipeJournal, Me.dgvPayment)
                        Me.SalinkeField()
                        Me.AktifkanDetail()
                    End If
                    Return False
                End If
            End If

            'pastikan semua kode akun ada di dalam master akun, di hide dulu
            For Each row In Me.dgvItem.Rows
                If IsNothing(row.Cells(Me.ColPrimaryAkun.Index).Value) = True Then
                    MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Return False
                Else

                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader2(row.Cells(Me.ColPrimaryAkun.Index).Value) = False Then
                        MsgBox("Invalid account in row " + row.Index.ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        row.Selected = True
                        Return False
                    End If
                End If
            Next

            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpDate.Value), _
            Me.ObjCD.GetLastDateOfTheYear(Me.dtpDate.Value.Date)) > 0 Then
                MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dim i As Integer = 1
            For Each row In Me.dgvPayment.Rows
                If i <> Me.dgvPayment.Rows.Count Then
                    If row.Cells(Me.ColItemPay_Amount.Index).Value = "" Or _
                    CDec(row.Cells(Me.COlItemPay_AmountOriginal.Index).Value) = 0 Then
                        MsgBox("Please enter a number for amount.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvPayment.ClearSelection()
                        Me.dgvPayment.Rows(row.Index).Selected = True
                        Return False

                        Exit Try
                    End If
                    If row.Cells(Me.COlItemPay_Account.Index).Value = Me.cboAccount.Text Then
                        MsgBox("Account should not be the same with account header", MsgBoxStyle.Information, "Information")
                        Me.dgvPayment.ClearSelection()
                        Me.dgvPayment.Rows(row.Index).Selected = True

                        Return False
                        Exit Try
                    End If
                    i = i + 1
                End If
            Next

            DTA.FillBy_PrimaryJurnal(TBL, Me.TxtPrimaryJournal.Text)
            For Each rowcek In TBL.Rows
                If rowcek.Cheque_date.Date < Me.dtpDate.Value.Date Then
                    MsgBox("There should not be cheque date smaller than jurnal date", MsgBoxStyle.Information, "Information")
                    Me.LinkLabel1.Focus()
                    Return False
                End If
            Next

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub txtRemarks_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Enter
        Try
            Me.Buffer = Me.TxtRemarks.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''back color dan tampilan 
    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.BackcolorAll
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Leave
        Try
            Me.TransferKeAccountDetail()
            Me.HitungAmount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = False Then
                    Me.TxtRemarks.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtRemarks.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.GotFocus
        Try
            If Me.TxtRefNo.Enabled = True Then
                If Me.TxtRefNo.ReadOnly = False Then
                    Me.TxtRefNo.BackColor = Me.BackcolorAll
                Else
                    TxtRefNo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScheduleNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRefNo.LostFocus
        Try
            If Me.TxtRefNo.Enabled = True Then
                If Me.TxtRefNo.ReadOnly = False Then
                    Me.TxtRefNo.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtRefNo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim PrimaryDummy As Double = -1
        Dim PrimaryAkunRE As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                'jika akun retained earning tidak ada maka tidak bisa lanjut
                PrimaryAkunRE = Me.ObjDefaultValue.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning) = ""
                If PrimaryAkunRE = "" Or PrimaryAkunRE = "-1" Then
                    MsgBox("Please specify retained earning account in global setting data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            'Me.dgvItem.AllowUserToAddRows = True
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtRefNo.Focus()
            Me.ObjJurnal.ResetObject(Me.dgvItem)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCOde
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCOde
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.ObjAkun.IsiCboCashAndBank(Me.cboAccount, Me.TBLAkun)
            If Me.cboAccount.Items.Count > 0 Then
                Me.cboAccount.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            'Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehEdit(Me.ObjJurnal.Primary, UserCOde, True, True) = True Then
                    Me.Baru = False
                    'Me.ObjJurnal.CariDetail( Me.ObjJurnal.Primary, M e.dgvItem, Me, TipeJurnall2, Me.dgvPayment)
                    'Me.SalinkeField()
                    Me.AktifkanDetail()
                    'Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvPayment.AllowUserToAddRows = True
                    Me.TxtRefNo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                If Me.ObjJurnal.ApakahBolehHapus(Me.TxtPrimaryJournal.Text, True, UserCOde, False, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjJurnal.Hapus(Me.TxtPrimaryJournal.Text, True, UserCOde, False) = True Then
                            Me.objcek.HapusItem(Me.TxtPrimaryJournal.Text)
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.ObjJurnal.TipeInOut = Me.cboType.SelectedIndex
            Me.ObjJurnal.TipeJournal = Me.TipeJournal
            Me.ObjJurnal.TipeJurnal2 = Me.TipeJurnall2
            ' Me.dgvItem.AllowUserToAddRows = False
            Me.SalinkeObjek()
            If Baru = True Then
                If ObjJurnal.ApakahDataBenar1(Baru, Me.dgvItem, Me.ColPrimaryCurrency.Index, Me.dtpDate.Value) = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If Me.ApakahDataBenar() = True Then
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru(Me.dgvItem, Me.dgvPayment, Me.UserCOde, Me, Me.FrmCek.dgvCheque, Me.FrmCek) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnFind.Focus()
                            End If
                            ' Me.dgvItem.AllowUserToAddRows = False
                            'Me.dgvPayment.AllowUserToAddRows = True
                        ElseIf hasil = MsgBoxResult.No Then
                            'Me.dgvItem.AllowUserToAddRows = True
                            Me.dgvPayment.AllowUserToAddRows = True
                        End If
                    Else
                        'Me.dgvItem.AllowUserToAddRows = True
                        Me.dgvPayment.AllowUserToAddRows = True
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvPayment.AllowUserToAddRows = True
                End If
            Else
                'konfirmasi disimpan atau tidak
                If Me.ApakahDataBenar = True Then
                    If ObjJurnal.ApakahBolehEdit(Me.TxtPrimaryJournal.Text, UserCOde, True, True) = True Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan(Me.dgvItem, Me.dgvPayment, Me.UserCOde, Me, Me.FrmCek, Me.FrmCek.dgvCheque) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinkeField()
                                Me.btnEdit.Focus()
                            End If
                            'Me.dgvItem.AllowUserToAddRows = False
                            'Me.dgvPayment.AllowUserToAddRows = True
                        ElseIf hasil = MsgBoxResult.No Then
                            'Me.dgvItem.AllowUserToAddRows = True
                            Me.dgvPayment.AllowUserToAddRows = True
                        End If
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvPayment.AllowUserToAddRows = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            'Me.dgvItem.AllowUserToAddRows = False
            Me.ObjJurnal.TipeInOut = Me.cboType.SelectedIndex
            Me.ObjJurnal.TipeJournal = Me.TipeJournal
            Me.ObjJurnal.TipeJurnal2 = Me.TipeJurnall2
            If Me.TxtRemarks.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtRefNo.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.NonAktifkanDetail()
                    ''Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.objcek.HapusItem(Me.TxtPrimaryJournal.Text)
                    Else
                        Me.ObjJurnal.CariDetail(Me.TxtPrimaryJournal.Text, Me.dgvItem, Me, TipeJournal, Me.dgvPayment)
                        Me.SalinkeField()
                    End If
                Else
                    'Me.dgvItem.AllowUserToAddRows = True
                    Me.dgvPayment.AllowUserToAddRows = True
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.FrmCek.dgvCheque.Rows.Clear()
        End Try
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            'Me.dgvItem.AllowUserToAddRows = False
            Me.ObjJurnal.TipeInOut = Me.cboType.SelectedIndex
            Me.ObjJurnal.TipeJournal = Me.TipeJournal
            Me.ObjJurnal.TipeJurnal2 = Me.TipeJurnall2
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinkeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanBaru(Me.dgvItem, Me.dgvPayment, Me.UserCOde, Me, Me.FrmCek.dgvCheque, Me.FrmCek) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjJurnal.SimpanPerubahan(Me.dgvItem, Me.dgvPayment, Me.UserCOde, Me, Me.FrmCek, Me.FrmCek.dgvCheque) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinkeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
            'Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Fungsi DGV
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
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
                End If
            Else
                'remarks, max 250 karakter
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
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

    Private Sub BalanceDebetKredit(ByVal vIndexAcuan As Integer, ByVal vAcuanDebet As Boolean, ByVal rounding As Integer, _
    Optional ByVal vHitungOriginalCurrency As Boolean = True)

        Try
            Dim rowAcuan As DataGridViewRow = Me.dgvItem.Rows(vIndexAcuan)
            Dim rowTarget As DataGridViewRow

            If Me.TxtRemarks.Enabled = True And Me.dgvItem.Rows.Count = 3 Then
                'dibalance jika hanya ada 2 jurnal, maka debet = kredit
                If rowAcuan.Index = 0 Then
                    rowTarget = Me.dgvItem.Rows(1)
                Else
                    rowTarget = Me.dgvItem.Rows(0)
                End If

                If vAcuanDebet = True Then
                    rowTarget.Cells(Me.ColKredit.Index).Value = rowAcuan.Cells(Me.ColDebet.Index).Value
                    rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                Else
                    rowTarget.Cells(Me.ColDebet.Index).Value = rowAcuan.Cells(Me.ColKredit.Index).Value
                    rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                End If

                If vHitungOriginalCurrency = True Then
                    Me.HitungOriginalValues(rowTarget.Index)
                End If
            ElseIf Me.TxtRemarks.Enabled = True And Me.dgvItem.Rows.Count > 3 Then
                'dibalance jika ada 3 jurnal, jika 2 dikredit atau 2 didebet
                If Me.GetJmlNol = 0 Then
                    If vAcuanDebet = True Then
                        If Me.GetJmlKredit = 1 Then
                            rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlKredit)
                            rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalDebet, rounding)
                            rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    Else
                        If Me.GetJmlDebet = 1 Then
                            rowTarget = Me.dgvItem.Rows(Me.GetIndexJmlDebet)
                            rowTarget.Cells(Me.ColDebet.Index).Value = Me.ObjCD.DigitGrouping(Me.GetTotalKredit, rounding)
                            rowTarget.Cells(Me.ColKredit.Index).Value = Me.ObjCD.DigitGrouping(0, rounding)
                        End If
                    End If

                    If vHitungOriginalCurrency = True Then
                        If IsNothing(rowTarget) = False Then
                            Me.HitungOriginalValues(rowTarget.Index)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    ''' <summary>
    ''' fungsi ini mengembalikan jumlah baris yang nilai debet dan kreditnya =0
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
                        If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> Me.PrimaryBaseCurrency Then
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

            'cek kolom job dan dept
            '2. tutup kolom departement dan kolom job, jika tidak ada di dalam masternya
            'If Me.ObjJob.GetJml = 0 Then
            Me.ColJob.Visible = False
            'Else
            'Me.ColJob.Visible = True
            'End If
            'If Me.ObjDept.GetJml = 0 Then
            Me.ColDept.Visible = False
            ' Else
            'Me.ColDept.Visible = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtRefNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtRefNo.Text <> "" Then
                Dim i As New FrmTransaksiJournal_Cetak1
                i.PrimaryJurnal = Me.TxtPrimaryJournal.Text
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungExchangeRate(ByVal vRowIndex As Integer)
        Dim vRowDGV As DataGridViewRow

        Try
            vRowDGV = Me.dgvItem.Rows(vRowIndex)
            vRowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(Math.Round(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
            CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value), DefPrecExchangeRate, MidpointRounding.AwayFromZero))
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
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
            Else
                vRowDGV.Cells(Me.ColKredit.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value) * _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecBase)
                vRowDGV.Cells(Me.ColDebet.Index).Value = ObjCD.DigitGrouping(0, DefPrecBase)
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
            DefPrecBase = Me.ObjCur.GetDefaultPrecission(Me.PrimaryBaseCurrency)
            DefPrecOriginal = Me.ObjCur.GetDefaultPrecission(vRowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)

            If CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) > 0 Then
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColDebet.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
            Else
                vRowDGV.Cells(Me.ColCredit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(CDec(vRowDGV.Cells(Me.ColKredit.Index).Value) / _
                CDec(vRowDGV.Cells(Me.ColExchangeRate.Index).Value), DefPrecOriginal)
                vRowDGV.Cells(Me.ColDebet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(0, DefPrecOriginal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            If Me.TxtRefNo.Enabled = True And Me.cboAccount.Text <> "" Then
                If Me.cboType.SelectedIndex = 0 Then
                    Me.FrmCek.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut._In
                Else
                    Me.FrmCek.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut.out
                End If
                Me.FrmCek.mPrimaryJurnal = Me.TxtPrimaryJournal.Text
                Me.FrmCek.mPrimaryCurr = Me.ObjCur.GetPrimary(Me.LblCurrency.Text)
                Me.FrmCek.mBaru = Me.Baru
                Me.FrmCek.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtExchangeRateHeader()
        Try
            'jika currency = Primary base currency, maka:
            If Core_Module.PrimaryBaseCurrency = Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text) Then
                '1. txtexchange rate = read only, tab stop = false
                Me.TxtExchangeRate.ReadOnly = True
                Me.TxtExchangeRate.TabStop = False
                Me.TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
            Else
                'jika currency <> primary base currency:
                '2. txtexchange rate = read only = false, tab stop = true
                Me.TxtExchangeRate.ReadOnly = False
                Me.TxtExchangeRate.TabStop = True
                Me.TxtExchangeRate.BackColor = Me.TxtPrimaryAccount.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Dim row As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboAccount.SelectedIndex <> -1 Then
                    Me.TxtPrimaryAccount.Text = Me.ObjAkun.GetPrimaryByDesc(Me.cboAccount.Text)
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCur.GetDefaultExhangeRate(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text)), _
                    Core_Module.DefPrecExchangeRate)

                    Me.LblCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text))
                    Me.FrmCek.mPrimaryCurr = Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text)

                    If Me.TxtRemarks.Enabled = True Then
                        Me.SetTxtExchangeRateHeader()
                        Me.TransferKeAccountDetail()
                        Me.HitungAmount()
                    End If
                Else
                    Me.TxtPrimaryAccount.Text = -1
                End If

                For Each row In Me.FrmCek.dgvCheque.Rows
                    row.Cells(Me.FrmCek.ColPrimaryCurrency.Index).Value = Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryAccount.Text)
                Next

                Me.SetTampilanDGVPayment()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPayment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPayment.CellBeginEdit
        Try
            Dim DefPrec As Integer = Me.ObjCur.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value))

            If String.IsNullOrEmpty(Me.dgvPayment(Me.COlItemPay_Account.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColItemPay_Amount.Index Or e.ColumnIndex = Me.COlItemPay_AmountOriginal.Index Or _
                e.ColumnIndex = Me.COlItemPay_DeptCode.Index Or e.ColumnIndex = Me.COlItemPay_JobCode.Index Or _
                e.ColumnIndex = Me.COlItemPay_Account.Index Or e.ColumnIndex = Me.COlItemPay_Remarks.Index Or _
                e.ColumnIndex = Me.ColItemPay_ExRate.Index Then
                    If e.ColumnIndex = Me.ColItemPay_ExRate.Index Then
                        If Me.ObjAkun.GetPrimaryCurrency(Me.dgvPayment.Item(Me.COlItemPay_PrimaryAkun.Index, e.RowIndex).Value) = _
                        Core_Module.PrimaryBaseCurrency Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.COlItemPay_AmountOriginal.Index Or _
                    e.ColumnIndex = Me.ColItemPay_ExRate.Index Then
                        Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                    Me.Buffer = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.COlItemPay_Account.Index Then
                    Me.Buffer = Me.dgvPayment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function GetIndexItemInput(ByVal vPrimaryAkun As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvPayment.Rows
                If rowDGV.Cells(Me.COlItemPay_PrimaryAkun.Index).Value = vPrimaryAkun Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub formatDgvAccount()
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColExchangeRate.Index).Value = 1 Then
                    Me.ColDebet_OriginalCurr.Visible = False
                    Me.ColCredit_OriginalCurrency.Visible = False
                    Me.ColExchangeRate.Visible = False
                Else
                    Me.ColDebet_OriginalCurr.Visible = True
                    Me.ColCredit_OriginalCurrency.Visible = True
                    Me.ColExchangeRate.Visible = True
                    Exit For
                End If
            Next
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColOriginalCurrency.Index).Value = "IDR" Then
                    Me.ColOriginalCurrency.Visible = False
                Else
                    Me.ColOriginalCurrency.Visible = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungAmount()
        Dim row As DataGridViewRow
        Dim amount As Decimal
        Dim amount_Ori As Decimal
        Try
            Me.TxtAmount.Text = 0
            Me.TxtAmountOriginal.Text = 0
            For Each row In Me.dgvItem.Rows
                If (row.Index + 1) Mod 2 <> 0 Then
                    If Me.cboType.SelectedIndex = 0 Then
                        'receive to

                        amount_Ori = CDec(row.Cells(Me.ColDebet_OriginalCurr.Index).Value - row.Cells(Me.ColCredit_OriginalCurrency.Index).Value)
                        Me.TxtAmountOriginal.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtAmountOriginal.Text) + amount_Ori, Me.ObjCur.GetDefaultPrecission _
                        (row.Cells(Me.ColPrimaryCurrency.Index).Value))

                        amount = CDec(row.Cells(Me.ColDebet.Index).Value - row.Cells(Me.ColKredit.Index).Value)
                        Me.TxtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtAmount.Text) + amount, Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency))
                    Else
                        amount_Ori = CDec(row.Cells(Me.ColCredit_OriginalCurrency.Index).Value - row.Cells(Me.ColDebet_OriginalCurr.Index).Value)
                        Me.TxtAmountOriginal.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtAmountOriginal.Text) + amount_Ori, Me.ObjCur.GetDefaultPrecission _
                        (row.Cells(Me.ColPrimaryCurrency.Index).Value))

                        amount = CDec(row.Cells(Me.ColKredit.Index).Value - row.Cells(Me.ColDebet.Index).Value)
                        Me.TxtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtAmount.Text) + amount, Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.GotFocus
        Try
            Me.Buffer = Me.TxtExchangeRate.Text
            If Me.TxtExchangeRate.Enabled = True Then
                If Me.TxtExchangeRate.ReadOnly = False Then
                    Me.TxtExchangeRate.BackColor = Me.BackcolorAll
                Else
                    TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtExchangeRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        Try
            If Me.ObjCD.CekBilanganPositif(Me.TxtExchangeRate.Text) = False Then
                MsgBox("Please enter positive number")
                Me.TxtExchangeRate.Text = Buffer
            End If
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.TxtExchangeRate.Text, Core_Module.DefPrecExchangeRate)

            Me.TransferKeAccountDetail()
            Me.HitungAmount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.LostFocus
        If Me.TxtExchangeRate.Enabled = True Then
            If Me.TxtExchangeRate.ReadOnly = False Then
                Me.TxtExchangeRate.BackColor = Me.TxtPrimaryJournal.BackColor
            Else
                TxtExchangeRate.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub cboType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboType.SelectedIndexChanged
        Dim row As DataGridViewRow
        Try
            If Me.cboType.SelectedIndex = 0 Then
                Me.FrmCek.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut._In
                Me.LblReferenceParty.Text = "Received From"
            Else
                Me.FrmCek.mTypeInOut = ClsTransaksiJournal.EnumTipeInOut.out
                Me.LblReferenceParty.Text = "Payment To"
            End If
            For Each row In Me.FrmCek.dgvCheque.Rows
                row.Cells(Me.FrmCek.ColTipeInOut.Index).Value = Me.FrmCek.mTypeInOut
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboType.SelectedValueChanged
        Try
            Me.TransferKeAccountDetail()
            Me.HitungAmount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPayment_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvPayment.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True And Me.SetDefaultValue = True Then
                rowDGV = Me.dgvPayment.Rows(Me.dgvPayment.Rows.Count - 1)
                rowDGV.Cells(Me.COlItemPay_Remarks.Index).Value = ""
                rowDGV.Cells(Me.COlItemPay_PrimaryJob.Index).Value = -1
                rowDGV.Cells(Me.COlItemPay_PrimaryDept.Index).Value = -1
                rowDGV.Cells(Me.COlItemPay_PrimaryAkun.Index).Value = -1
                rowDGV.Cells(Me.COlItemPay_JobCode.Index).Value = ""
                rowDGV.Cells(Me.ColItemPay_ExRate.Index).Value = "1.0000"
                rowDGV.Cells(Me.COlItemPay_Desc.Index).Value = ""
                rowDGV.Cells(Me.COlItemPay_DeptCode.Index).Value = ""
                rowDGV.Cells(Me.ColItemPay_Curr.Index).Value = ""
                rowDGV.Cells(Me.COlItemPay_AmountOriginal.Index).Value = "0.00"
                rowDGV.Cells(Me.ColItemPay_Amount.Index).Value = "0.00"
                rowDGV.Cells(Me.COlItemPay_Account.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPayment_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvPayment.UserDeletedRow
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.TransferKeAccountDetail()
                Me.HitungAmount()
                Me.SetTampilanDGVPayment()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTampilanDGVPayment()
        Dim SemuaIDR As Boolean = True
        Dim rowDGV As DataGridViewRow

        Try
            '1. tutup exchange rate bila semua item payment, memiliki currency idr
            For Each rowDGV In Me.dgvPayment.Rows
                If rowDGV.Cells(Me.COlItemPay_PrimaryAkun.Index).Value <> -1 Then
                    If Me.ObjAkun.GetPrimaryCurrency(rowDGV.Cells(Me.COlItemPay_PrimaryAkun.Index).Value) <> Core_Module.PrimaryBaseCurrency Then
                        SemuaIDR = False
                        Exit For
                    End If
                End If
            Next
            If SemuaIDR = True Then
                Me.ColItemPay_ExRate.Visible = False
            Else
                Me.ColItemPay_ExRate.Visible = True
            End If

            '2. tutup kolom departement dan kolom job, jika tidak ada di dalam masternya
            'If Me.ObjJob.GetJml = 0 Then
            Me.COlItemPay_JobCode.Visible = False
            'End If
            'If Me.ObjDept.GetJml = 0 Then
            Me.COlItemPay_DeptCode.Visible = False
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvPayment.Rows.RemoveAt(Me.IndexRemove)
            Me.IndexRemove = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReferenceParty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReferenceParty.GotFocus
        Try
            If Me.TxtReferenceParty.Enabled = True Then
                If Me.TxtReferenceParty.ReadOnly = False Then
                    Me.TxtReferenceParty.BackColor = Me.BackcolorAll
                Else
                    TxtReferenceParty.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReferenceParty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReferenceParty.LostFocus
        Try
            If Me.TxtReferenceParty.Enabled = True Then
                If Me.TxtReferenceParty.ReadOnly = False Then
                    Me.TxtReferenceParty.BackColor = Me.TxtPrimaryJournal.BackColor
                Else
                    TxtReferenceParty.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class