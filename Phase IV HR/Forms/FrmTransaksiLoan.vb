Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms

Public Class FrmTransaksiLoan
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public EnumDipanggilDari As ClsTransaksiLoan.EnumFrmTransLoanDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan
    Dim baru As Boolean
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjLoan As New ClsTransaksiLoan
    Dim ObjAcc As New Core_Module.ClsMasterAkun1
    Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
    Dim ObjTransJurnal As New AccModule.ClsTransaksiJournal
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim defprec As Integer = ObjCurr.GetDecimal_DefaultCurrency()
    Dim BufferDate As String = ""
    Dim BufferLoan As String = ""
    Dim BufferType As Integer = -1
    Dim Buffer As String = ""
    Dim ObjClosing As New AccModule.ClsTransaksiClosing
    'buat hitungan
    Dim Outstanding As Decimal = 0
    Dim pokok As Decimal = 0
    Dim bunga As Decimal = 0
    Dim Angsuran As Decimal = 0
    Dim objJournal As New AccModule.ClsTransaksiJournal

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.ObjLoan.IsiCboCashAcc(Me.CboCashAccount)
    End Sub

    Private Sub FrmTransaksiLoan_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtEmployeeCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjLoan.SimpanBaru(Me.UserCode, Me.dgvItem, Me) = True Then
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
                            If Me.ObjLoan.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Text = "Loan Application"
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnInformation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInformation.Click
        Dim ij As New FrmTransaksiLoan_Information
        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.SetBackColor(Me.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtPrimaryLoan.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtPrimaryLoan.BackColor, False)

            Me.CboStatus.Enabled = False
            If Me.baru = False Then
                If Me.ObjLoan.GetJmlLoanPayment(Me.TxtPrimaryLoan.Text) > 0 Then
                    Me.NonAktifkanDetail()
                    Me.dtpDate.Enabled = True
                    Me.btnOK.Enabled = True
                    Me.btnCancel.Enabled = True
                End If
            End If
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.dgvItem.ReadOnly = False
            Me.CancelButton = Me.btnCancel
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.SetBackColor(Me.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)

            Me.dgvItem.ReadOnly = True
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.TxtInstallmentTo.Text = Me.ObjCD.DigitGrouping(0, Me.defprec)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)

            Me.TxtPrimaryReceivableAcc.Text = -1
            Me.TxtPrimaryLoan.Text = -1
            Me.TxtPrimaryJurnal.Text = -1
            Me.TxtPrimaryEmployee.Text = -1
            Me.TxtPrimaryCashAcc.Text = -1
            Me.TxtTotalAmount.Text = Me.ObjCD.DigitGrouping(0, Me.defprec)
            Me.TxtRatePersen.Text = Me.ObjCD.DigitGrouping(0, Me.defprec)
            Me.dgvItem.Rows.Clear()
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ObjJurnal As New ClsTransaksiJournal1
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Dim RowDGV As DataGridViewRow
        Dim PayrollConventionDate As Date = Me.ObjClosing.GetPayrollConvertionDate()
        Dim DocNoBuffer As String = ""

        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Employee code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode)
            KoleksiNamaData.Add("Receivable account")
            KoleksiTextBox.Add(Me.TxtReceivableAccount)
            KoleksiNamaData.Add("Loan amount")
            KoleksiTextBox.Add(Me.TxtTotalAmount)
            KoleksiNamaData.Add("Rate yearly")
            KoleksiTextBox.Add(Me.TxtRatePersen)
            KoleksiNamaData.Add("Installment to")
            KoleksiTextBox.Add(Me.TxtInstallmentTo)
            KoleksiNamaData.Add("Journal No")
            KoleksiTextBox.Add(Me.txtJournalNo)
            KoleksiNamaData.Add("Journal Description")
            KoleksiTextBox.Add(Me.txtJournalDesc)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'If Me.CboCashAccount.SelectedText = "" Then
            '    MsgBox("Please select Cash out account.", MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'Harus input Convention date HR
            If PayrollConventionDate = #1/1/1950# Then
                MsgBox("Please input payroll convention date in global setting data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh double
            If Me.baru = True Then
                If Me.ObjLoan.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                    DocNoBuffer = Me.TxtDocumentNo.Text
                    Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
                    'MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                    'Me.TxtDocumentNo.Focus()
                    'Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjLoan.GetDocNoByPrimary(Me.TxtPrimaryLoan.Text)
                If CodeLama <> Me.TxtDocumentNo.Text Then
                    If Me.ObjLoan.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                        DocNoBuffer = Me.TxtDocumentNo.Text
                        Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
                        'MsgBox("Duplication in Document no, please change your code", MsgBoxStyle.Information, "Information")
                        'Me.TxtDocumentNo.Focus()
                        'Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjLoan.GetTimeStampByPrimary(Me.TxtPrimaryLoan.Text)
                If Me.ObjLoan.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjLoan.CariDetail(Me.TxtPrimaryJurnal.Text, Me, Me.dgvItem, Me.defprec)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'harus ada minimal 1 item 
            If Me.dgvItem.RowCount < 1 Then
                MsgBox("Please input item", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.dtpDate.Value > Me.DtpFirst.Value Then
                MsgBox("First payment date can't be smaller than transaction date", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In Me.dgvItem.Rows
                    If CDec(RowDGV.Cells(Me.ColInstallmentAmount.Index).Value) = 0 Then
                        MsgBox("Installment Amount can't input zero value.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
            End If

            If Me.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan Then
                'harus dibuat setelah konvention date HR
                If Me.dtpDate.Value < PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction loan, loan's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Me.dtpDate.Focus()
                    Return False
                End If
            ElseIf Me.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                'harus dibuat setelah konvention date HR
                If Me.dtpDate.Value >= PayrollConventionDate Then
                    MsgBox("Please make sure date in transaction beginning balance loan, loan's date must be bigger than payroll convention date.", MsgBoxStyle.Information, "Information")
                    Me.dtpDate.Focus()
                    Return False
                End If
            End If

            If DocNoBuffer <> "" Then
                If Me.txtJournalNo.Text = DocNoBuffer Then
                    Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
                End If
                If Me.txtJournalDesc.Text.Contains(DocNoBuffer) = True Then
                    Me.txtJournalDesc.Text = Me.txtJournalDesc.Text.Replace(DocNoBuffer, Me.TxtDocumentNo.Text)
                End If
            Else
                If Me.baru And Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                    MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                    Me.txtJournalNo.Focus()
                    Return False
                ElseIf Me.baru = False And Me.txtJournalNo.Text <> Me.objJournal.GetRefCode(Me.ObjLoan.GetPrimaryJournal(Me.TxtPrimaryLoan.Text)) And _
                        Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                    MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                    Me.txtJournalNo.Focus()
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            With Me.ObjLoan
                Me.TxtPrimaryLoan.Text = .Primary
                Me.TxtDocumentNo.Text = .DocNo
                Me.dtpDate.Value = .Date_
                Me.DtpFirst.Value = .FirstPaymentDate
                Me.TxtPrimaryEmployee.Text = .PrimaryKaryawan
                Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(.PrimaryKaryawan)
                Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(.PrimaryKaryawan)
                Me.TxtTotalAmount.Text = Me.ObjCD.DigitGrouping(.LoanAmount, Me.defprec)
                Me.CboLoanType.SelectedIndex = .LoanType
                Me.TxtRatePersen.Text = Me.ObjCD.DigitGrouping(.PersenRate, 2)
                Me.TxtInstallmentTo.Text = Me.ObjCD.DigitGrouping(.InstallmentTo, 0)
                Me.TxtPrimaryCashAcc.Text = .PrimaryAccCashOut
                Me.CboCashAccount.Text = Me.ObjAcc.GetAccountCode(.PrimaryAccCashOut)
                Me.TxtDescCash.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(.PrimaryAccCashOut)
                Me.TxtPrimaryReceivableAcc.Text = .PrimaryAccRev
                Me.TxtDescRec.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(.PrimaryAccRev)
                Me.TxtReceivableAccount.Text = Me.ObjAcc.GetAccountCode(.PrimaryAccRev)
                Me.TxtRemark.Text = .Remarks
                Me.CboStatus.SelectedIndex = .Status
                Me.TxtPrimaryJurnal.Text = .PrimaryJurnal
                Me.txtJournalNo.Text = .JournalNo  'Me.objJournal.GetRefCode(Me.TxtPrimaryJurnal.Text)
                Me.txtJournalDesc.Text = .JournalDesc 'Me.objJournal.GetDescription(Me.TxtPrimaryJurnal.Text)
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjLoan
                .DocNo = Me.TxtDocumentNo.Text
                .Urutan = Me.TxtDocumentNo.Text.Substring(0, 6)
                .Date_ = Me.dtpDate.Value
                .FirstPaymentDate = Me.DtpFirst.Value
                .PrimaryKaryawan = Me.TxtPrimaryEmployee.Text
                .LoanAmount = Me.TxtTotalAmount.Text
                .LoanType = Me.CboLoanType.SelectedIndex
                .PersenRate = Me.TxtRatePersen.Text
                .InstallmentTo = Me.TxtInstallmentTo.Text
                .PrimaryAccCashOut = Me.TxtPrimaryCashAcc.Text
                .PrimaryAccRev = Me.TxtPrimaryReceivableAcc.Text
                .Remarks = Me.TxtRemark.Text
                .Status = Me.CboStatus.SelectedIndex
                .PrimaryJurnal = Me.TxtPrimaryJurnal.Text
                .JournalNo = Me.txtJournalNo.Text
                .JournalDesc = Me.txtJournalDesc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungNoRate(ByVal vPlafon As Decimal, ByVal vTenor As Double, ByVal vBaris As Integer)
        Try
            Me.bunga = 0
            Me.pokok = 0
            Me.pokok = Math.Round(vPlafon / vTenor, Me.defprec, MidpointRounding.AwayFromZero)
            Me.Angsuran = Me.pokok
            If vBaris = 1 Then
                'kalo baris pertama
                Me.Outstanding = vPlafon - Me.Angsuran
            Else
                If vBaris <> vTenor Then
                    'kalo bukan baris terakhir
                    Me.Outstanding = Me.Outstanding - Me.Angsuran
                Else
                    'kalo baris terakhir
                    Me.pokok = Me.Outstanding
                    Me.Angsuran = Me.Outstanding
                    Me.Outstanding = Me.pokok - Me.Outstanding
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fixed installment dengan bunga
    Private Sub HitungAnuitas(ByVal vPlafon As Decimal, ByVal vTenor As Decimal, _
    ByVal vRate As Decimal, ByVal vBaris As Integer)
        Try
            Me.pokok = 0
            Me.bunga = 0

            'angsuran
            If vRate > 0 Then
                Me.Angsuran = Math.Round(vPlafon * ((vRate / 100) / 12) * (1 / (1 - (1 / ((1 + (vRate / 100) / 12) ^ vTenor)))), Me.defprec, MidpointRounding.AwayFromZero)
            Else
                Me.Angsuran = Math.Round(vPlafon / vTenor, Me.defprec, MidpointRounding.AwayFromZero)
            End If

            'bunga
            If vBaris = 1 Then
                'jika baris pertama
                Me.bunga = Math.Round(vPlafon * ((vRate / 100) / 12), Me.defprec, MidpointRounding.AwayFromZero)
            Else
                Me.bunga = Math.Round(Me.Outstanding * ((vRate / 100) / 12), Me.defprec, MidpointRounding.AwayFromZero)
            End If

            'pokok
            Me.pokok = Me.Angsuran - Me.bunga

            'out standing
            If vBaris = 1 Then
                Me.Outstanding = vPlafon - Me.pokok
            Else
                Me.Outstanding = Me.Outstanding - Me.pokok
                If vTenor = vBaris Then
                    If Outstanding > 0 Then
                        Me.pokok -= Outstanding
                        Me.Angsuran += Outstanding
                    ElseIf Outstanding < 0 Then
                        Me.pokok += Outstanding
                        Me.Angsuran -= Outstanding
                    End If
                    Me.Outstanding = 0
                    Me.Angsuran = Me.pokok + Me.bunga
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungFlat(ByVal vPlafon As Decimal, ByVal vTenor As Decimal, _
    ByVal vRate As Decimal, ByVal vBaris As Integer)
        Try
            Me.pokok = 0
            Me.bunga = 0
            'bunga
            Me.bunga = Math.Round((vPlafon * ((vRate / 100) * (vTenor / 12))) / vTenor, Me.defprec, MidpointRounding.AwayFromZero)
            'pokok
            Me.pokok = Math.Round(vPlafon / vTenor, Me.defprec, MidpointRounding.AwayFromZero)
            'angsuran
            Me.Angsuran = Me.bunga + Me.pokok
            'out standing
            If vBaris = 1 Then
                Me.Outstanding = vPlafon - Me.pokok
            Else
                Me.Outstanding = Me.Outstanding - Me.pokok
                If vTenor = vBaris Then
                    If Outstanding > 0 Then
                        Me.pokok -= Outstanding
                        Me.Angsuran += Outstanding
                    ElseIf Outstanding < 0 Then
                        Me.pokok += Outstanding
                        Me.Angsuran -= Outstanding
                    End If
                    Me.Outstanding = 0
                    Me.Angsuran = Me.pokok + Me.bunga
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungEfektif(ByVal vPlafon As Decimal, ByVal vTenor As Decimal, _
    ByVal vRate As Decimal, ByVal vBaris As Integer)
        Try
            Me.pokok = 0
            Me.bunga = 0
            
            'pokok
            Me.pokok = Math.Round(vPlafon / vTenor, Me.defprec, MidpointRounding.AwayFromZero)
            'bunga
            Me.bunga = Math.Round((vPlafon - ((vBaris - 1) * Me.pokok)) * ((vRate / 100) / 12), Me.defprec, MidpointRounding.AwayFromZero)
            'angsuran
            Me.Angsuran = Me.bunga + Me.pokok
            'out standing
            If vBaris = 1 Then
                Me.Outstanding = vPlafon - Me.pokok
            Else
                Me.Outstanding = Me.Outstanding - Me.pokok
                If vTenor = vBaris Then
                    If Outstanding > 0 Then
                        Me.pokok -= Outstanding
                        Me.Angsuran += Outstanding
                    ElseIf Outstanding < 0 Then
                        Me.pokok += Outstanding
                        Me.Angsuran -= Outstanding
                    End If
                    Me.Outstanding = 0
                    Me.Angsuran = Me.pokok + Me.bunga
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
            Me.ObjLoan.IsiCboCashAcc(Me.CboCashAccount)
            Me.CboCashAccount.Refresh()
            Me.TxtRatePersen.Enabled = False
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.dtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim IsBeginingBalance As Boolean = False

        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                    IsBeginingBalance = True
                Else
                    IsBeginingBalance = False
                End If

                If Me.ObjLoan.ApakahBolehedit(Me.UserCode, Me.TxtPrimaryLoan.Text, IsBeginingBalance) = True Then
                    Me.baru = False
                    Me.ObjLoan.CariDetail(Me.TxtPrimaryLoan.Text, Me, Me.dgvItem, Me.defprec)
                    Me.AktifkanDetail()
                    Me.TxtEmployeeCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjLoan.SimpanBaru(Me.UserCode, Me.dgvItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.SalinKeObject()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjLoan.SimpanPerubahan(Me.UserCode, Me.dgvItem, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.SalinKeObject()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.ObjLoan.CariDetail(Me.TxtPrimaryLoan.Text, Me, Me.dgvItem, Me.defprec)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Event Textbox"

    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.GotFocus
        Me.Buffer = Me.TxtPrimaryEmployee.Text
        Me.TxtEmployeeCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.LostFocus
        Me.TxtEmployeeCode.BackColor = Color.White
    End Sub

    Private Sub TxtTotalAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalAmount.GotFocus
        Try
            Me.ObjCD.Number_Enter(Me.TxtTotalAmount, sender, e)
            Me.TxtTotalAmount.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalAmount.LostFocus
        Try
            Me.TxtTotalAmount.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtTotalAmount, sender, e, 999999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtTotalAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalAmount.Leave
        Try
            If Me.ObjCD.CekNonNegatif(Me.TxtTotalAmount.Text) = False Then
                MsgBox("Please input positif number.", MsgBoxStyle.Information, "Information")
                Me.TxtTotalAmount.Focus()
            Else
                Me.TxtTotalAmount.Text = Me.ObjCD.DigitGrouping(Me.TxtTotalAmount.Text, Me.defprec)
                If Me.TxtTotalAmount.Text <> Me.BufferLoan Then
                    Me.IsiDgvLoan(Me.TxtInstallmentTo.Text, Me.DtpFirst.Value, Me.dgvItem)
                    Me.BufferLoan = Me.TxtTotalAmount.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInstallmentTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInstallmentTo.GotFocus
        Try
            Me.ObjCD.Number_Enter(Me.TxtInstallmentTo, sender, e)
            Me.TxtInstallmentTo.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInstallmentTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInstallmentTo.LostFocus
        Try
            Me.TxtInstallmentTo.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInstallmentTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtInstallmentTo.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtInstallmentTo, sender, e, 999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInstallmentTo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtInstallmentTo.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtInstallmentTo, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInstallmentTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInstallmentTo.Leave
        Try
            If Me.ObjCD.CekNonNegatif(Me.TxtInstallmentTo.Text) = False Then
                MsgBox("Please input positif number.", MsgBoxStyle.Information, "Information")
                Me.TxtInstallmentTo.Focus()
            Else
                Me.IsiDgvLoan(Me.TxtInstallmentTo.Text, Me.DtpFirst.Value, Me.dgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatePersen_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatePersen.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtRatePersen, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatePersen_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatePersen.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtRatePersen, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatePersen_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatePersen.GotFocus
        Try
            Me.ObjCD.Number_Enter(Me.TxtRatePersen, sender, e)
            Me.TxtRatePersen.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatePersen_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatePersen.LostFocus
        Try
            Me.TxtRatePersen.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatePersen_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatePersen.Leave
        Try
            If Me.ObjCD.CekNonNegatif(Me.TxtRatePersen.Text) = False Then
                MsgBox("Please input positif number.", MsgBoxStyle.Information, "Information")
                Me.TxtRatePersen.Focus()
            Else
                Me.TxtRatePersen.Text = Me.ObjCD.DigitGrouping(Me.TxtRatePersen.Text, 2)
            End If
            If Me.TxtRatePersen.Enabled = True Then
                Me.IsiDgvLoan(Me.TxtInstallmentTo.Text, Me.DtpFirst.Value, Me.dgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccount.GotFocus
        Me.TxtReceivableAccount.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtReceivableAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccount.LostFocus
        Me.TxtReceivableAccount.BackColor = Color.White
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Me.TxtRemark.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Me.TxtRemark.BackColor = Color.White
    End Sub
#End Region

#Region "Cari Karyawan"

    Private Sub BtnFindEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindEmployee.Click
        Dim ij As New Core_Module.FrmMasterKaryawanCari
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.txtCari.Text = Me.TxtEmployeeCode.Text
            ij.ShowDialog()

            Me.Buffer = Me.TxtPrimaryEmployee.Text
            If CDbl(ij.PrimaryEmploye) > 0 Then
                Me.TxtPrimaryEmployee.Text = ij.PrimaryEmploye
                Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(ij.PrimaryEmploye)
                Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(ij.PrimaryEmploye)
                Me.TxtTotalAmount.Focus()
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
            End If
            Me.SetTxtJournal(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.ObjKaryawan.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtTotalAmount.Focus()
                Else
                    Me.BtnFindEmployee.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.Leave
        Dim ij As New Core_Module.FrmMasterKaryawanCari
        Try
            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.ObjKaryawan.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.ObjKaryawan.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.ObjKaryawan.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                Else
                    Me.BtnFindEmployee.PerformClick()
                End If
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
            End If
            Me.SetTxtJournal(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Cari Receivable Acc"

    Private Sub BtnFindReceivable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindReceivable.Click
        Dim ij As New Core_Module.FrmMasterAkun_Cari1
        Try
            'ij.AkunAsset = ClsMasterAsset.EnumAkunAsset.Asset
            ij.BackColorAll = Me.BackColorAll
            ij.txtCari.Text = Me.TxtReceivableAccount.Text
            ij.ShowDialog()

            If Me.ObjAcc.GetTipeAkun(ij.mPrimaryAkun) <> Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset Then
                MsgBox("Receivable account must asset account.", MsgBoxStyle.Information, "Information")
                Me.TxtPrimaryReceivableAcc.Text = -1
                Me.TxtReceivableAccount.Text = ""
                Me.TxtDescRec.Text = ""
                Me.TxtReceivableAccount.Focus()
                Exit Sub
            End If
            If ij.mPrimaryAkun = Me.TxtPrimaryCashAcc.Text And Me.TxtPrimaryCashAcc.Text <> -1 Then
                MsgBox("Receivable account cannot be same with cash out account.", MsgBoxStyle.Information, "Information")
                Me.TxtPrimaryReceivableAcc.Text = -1
                Me.TxtReceivableAccount.Text = ""
                Me.TxtDescRec.Text = ""
                Me.TxtReceivableAccount.Focus()
            Else
                If CDbl(ij.mPrimaryAkun) > 0 Then
                    Me.TxtPrimaryReceivableAcc.Text = ij.mPrimaryAkun
                    Me.TxtReceivableAccount.Text = Me.ObjAcc.GetAccountCode(ij.mPrimaryAkun)
                    Me.TxtDescRec.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                    Me.TxtRemark.Focus()
                Else
                    Me.TxtPrimaryReceivableAcc.Text = -1
                    Me.TxtReceivableAccount.Text = ""
                    Me.TxtDescRec.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtReceivableAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtReceivableAccount.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtReceivableAccount.Text = Me.CboCashAccount.Text And Me.CboCashAccount.Text <> "" Then
                    MsgBox("Receivable account cannot be same with cash out account.", MsgBoxStyle.Information, "Information")
                    Me.TxtPrimaryReceivableAcc.Text = -1
                    Me.TxtReceivableAccount.Text = ""
                    Me.TxtDescRec.Text = ""
                    Me.TxtReceivableAccount.Focus()
                Else
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.TxtReceivableAccount.Text) = True Then
                        If Me.ObjAcc.GetTipeAkun(Me.ObjAcc.GetPrimary(Me.TxtReceivableAccount.Text)) <> Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset Then
                            MsgBox("Receivable account must asset account.", MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryReceivableAcc.Text = -1
                            Me.TxtReceivableAccount.Text = ""
                            Me.TxtDescRec.Text = ""
                            Me.TxtReceivableAccount.Focus()
                            Exit Sub
                        End If
                        Me.TxtPrimaryReceivableAcc.Text = Me.ObjAcc.GetPrimary(Me.TxtReceivableAccount.Text)
                        Me.TxtReceivableAccount.Text = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtDescRec.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtRemark.Focus()
                    Else
                        Me.BtnFindReceivable.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReceivableAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReceivableAccount.Leave
        Dim ij As New Core_Module.FrmMasterAkun_Cari3
        Try
            If Me.TxtReceivableAccount.Text <> "" Then
                If Me.TxtReceivableAccount.Text = Me.CboCashAccount.Text And Me.CboCashAccount.Text <> "" Then
                    MsgBox("Receivable account cannot be same with cash out account.", MsgBoxStyle.Information, "Information")
                    Me.TxtPrimaryReceivableAcc.Text = -1
                    Me.TxtReceivableAccount.Text = ""
                    Me.TxtDescRec.Text = ""
                    Me.TxtReceivableAccount.Focus()
                Else
                    If Me.ObjAcc.ApakahSudahAda_Aktif_NonHeader(Me.TxtReceivableAccount.Text) = True Then
                        If Me.ObjAcc.GetTipeAkun(Me.ObjAcc.GetPrimary(Me.TxtReceivableAccount.Text)) <> Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset Then
                            MsgBox("Receivable account must asset account.", MsgBoxStyle.Information, "Information")
                            Me.TxtPrimaryReceivableAcc.Text = -1
                            Me.TxtReceivableAccount.Text = ""
                            Me.TxtDescRec.Text = ""
                            Me.TxtReceivableAccount.Focus()
                            Exit Sub
                        End If
                        Me.TxtPrimaryReceivableAcc.Text = Me.ObjAcc.GetPrimary(Me.TxtReceivableAccount.Text)
                        Me.TxtReceivableAccount.Text = Me.ObjAcc.GetAccountCode(Me.TxtPrimaryReceivableAcc.Text)
                        Me.TxtDescRec.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryReceivableAcc.Text)
                    Else
                        Me.BtnFindReceivable.PerformClick()
                    End If
                End If
            Else
                Me.TxtPrimaryReceivableAcc.Text = -1
                Me.TxtReceivableAccount.Text = ""
                Me.TxtDescRec.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

    Private Sub CboLoanType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboLoanType.SelectedIndexChanged
        Try
            If Me.CboLoanType.SelectedIndex = 0 Then
                Me.TxtRatePersen.Text = Me.ObjCD.DigitGrouping(0, 2)
                Me.TxtRatePersen.Enabled = False
            Else
                Me.TxtRatePersen.Enabled = True
            End If
            If Me.CboLoanType.Enabled = True Then
                If Me.CboLoanType.SelectedIndex <> Me.BufferType Then
                    Me.IsiDgvLoan(Me.TxtInstallmentTo.Text, Me.DtpFirst.Value, Me.dgvItem)
                    Me.BufferType = Me.CboLoanType.SelectedIndex
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        End Try
    End Sub

    Private Sub CboCashAccount_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCashAccount.SelectedValueChanged
        Try
            Me.TxtPrimaryCashAcc.Text = Me.ObjAcc.GetPrimary(Me.CboCashAccount.Text)
            Me.TxtDescCash.Text = Me.ObjAcc.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCashAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.dtpDate.Enabled = True Then
                Me.Buffer = Me.TxtDocumentNo.Text
                'Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgvLoan(ByVal vTotalIstalment As Integer, ByVal vTanggalAwal As Date, ByVal vDgv As DataGridView)
        Dim Index As Integer = 0
        Dim RowDgv As DataGridViewRow
        Dim ObjItem As New ClsTransaksiLoan_Item
        Try
            Me.Cursor = Cursors.WaitCursor
            vDgv.Rows.Clear()
            While (Index < vTotalIstalment)
                vDgv.Rows.Add()
                RowDgv = vDgv.Rows(vDgv.Rows.Count - 1)
                RowDgv.Cells.Item(Me.ColIntallmentPeriod.Index).Value = Index + 1
                'date schedule
                If vDgv.Rows.Count <= 1 Then
                    RowDgv.Cells.Item(Me.ColDateSchedule.Index).Value = vTanggalAwal
                Else
                    RowDgv.Cells.Item(Me.ColDateSchedule.Index).Value = _
                    Me.ObjCD.GetNextMonthDate(vDgv.Item(Me.ColDateSchedule.Index, Index - 1).Value, vTanggalAwal)
                End If
                'cari perhitungan type loan nya
                If Me.CboLoanType.SelectedIndex = ClsTransaksiLoan.EnumLoanType.NoRate Then
                    Me.HitungNoRate(Me.TxtTotalAmount.Text, Me.TxtInstallmentTo.Text, RowDgv.Index + 1)
                ElseIf Me.CboLoanType.SelectedIndex = ClsTransaksiLoan.EnumLoanType.Annuity Then
                    Me.HitungAnuitas(Me.TxtTotalAmount.Text, Me.TxtInstallmentTo.Text, Me.TxtRatePersen.Text, RowDgv.Index + 1)
                ElseIf Me.CboLoanType.SelectedIndex = ClsTransaksiLoan.EnumLoanType.Flat Then
                    Me.HitungFlat(Me.TxtTotalAmount.Text, Me.TxtInstallmentTo.Text, Me.TxtRatePersen.Text, RowDgv.Index + 1)
                ElseIf Me.CboLoanType.SelectedIndex = ClsTransaksiLoan.EnumLoanType.Effective Then
                    Me.HitungEfektif(Me.TxtTotalAmount.Text, Me.TxtInstallmentTo.Text, Me.TxtRatePersen.Text, RowDgv.Index + 1)
                End If
                RowDgv.Cells.Item(Me.ColInterest.Index).Value = Me.ObjCD.DigitGrouping(Me.bunga, Me.defprec)
                RowDgv.Cells.Item(Me.ColBase.Index).Value = Me.ObjCD.DigitGrouping(Me.pokok, Me.defprec)
                RowDgv.Cells.Item(Me.ColInstallmentAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.Angsuran, Me.defprec)
                RowDgv.Cells.Item(Me.ColBalance.Index).Value = Me.ObjCD.DigitGrouping(Me.Outstanding, Me.defprec)
                If RowDgv.Cells(Me.ColIdentifier.Index).Value = "" Then
                    Dim Identifier As String = ""
                    Identifier = ObjItem.GetIdentifierUrutan _
                    (Me.dgvItem, Me.ColIdentifier.Index, Me.baru, Me.TxtPrimaryJurnal.Text, RowDgv.Index)
                    RowDgv.Cells(Me.ColIdentifier.Index).Value = Identifier
                End If
                Index = Index + 1
            End While
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFirst_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpFirst.ValueChanged
        Try
            If Me.DtpFirst.Enabled = True Then
                If Me.DtpFirst.Value.ToString <> Me.BufferDate Then
                    Me.IsiDgvLoan(Me.TxtInstallmentTo.Text, Me.DtpFirst.Value, Me.dgvItem)
                    Me.BufferDate = Me.DtpFirst.Value.ToString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim IsBeginingBalance As Boolean = False

        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                    IsBeginingBalance = True
                Else
                    IsBeginingBalance = False
                End If

                If Me.ObjLoan.ApakahBolehHapus(Me.UserCode, Me.TxtPrimaryLoan.Text, IsBeginingBalance) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjLoan.Hapus(Me.TxtPrimaryLoan.Text, Me.UserCode, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiLoan_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.FormTransaksiLoan = Me
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjLoan.CariDetail(hj.PrimaryTransaksi, Me, Me.dgvItem, Me.defprec) = True Then
                    Me.baru = False
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim Frm As New FrmPilihCetakLoan
        Dim Frmdetail As New FrmTransaksiLoan_CetakDetail
        Dim FrmJournal As New FrmTransaksiLoan_CetakJournal
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.EnumDipanggilDari = ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                    Frm.vFromBeginingBalance = True
                Else
                    If Frm.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        Exit Sub
                    End If
                End If
                If Frm.RbtJournal.Checked = True Then
                    FrmJournal.PrimaryJournal = Me.TxtPrimaryJurnal.Text
                    FrmJournal.ObjLoan = Me.ObjLoan
                    FrmJournal.cetak = FrmTransaksiLoan_CetakJournal.EnumCetak.Loan
                    FrmJournal.ShowDialog()
                Else
                    Frmdetail.PrimaryLoan = Me.TxtPrimaryLoan.Text
                    Frmdetail.Type = Me.CboLoanType.Text
                    Frmdetail.ObjLoan = Me.ObjLoan
                    Frmdetail.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
            Frmdetail.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvItem(Me.ColIntallmentPeriod.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtJournalNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalNo.GotFocus
        Me.txtJournalNo.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtJournalNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalNo.LostFocus
        Me.txtJournalNo.BackColor = Color.White
    End Sub

    Private Sub txtJournalDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalDesc.GotFocus
        Me.txtJournalDesc.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtJournalDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJournalDesc.LostFocus
        Me.txtJournalDesc.BackColor = Color.White
    End Sub

    Private Sub TxtDocumentNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocumentNo.TextChanged
        Try
            Me.SetTxtJournal(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtJournal(ByVal vChangeDoc As Boolean)
        Dim JournalDesc As String = ""

        Try
            'set journal no
            If Me.txtJournalNo.Text = Me.Buffer Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            ElseIf Me.txtJournalNo.Text = "" Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            End If

            'set journal description
            If Me.txtJournalDesc.Text = "" Then
                Me.txtJournalDesc.Text = Me.TxtDocumentNo.Text
            End If
            
            If vChangeDoc = True Then
                If Me.Buffer <> "" Then
                    If Me.txtJournalDesc.Text.Contains(Me.Buffer) = True Then
                        Me.txtJournalDesc.Text = Me.txtJournalDesc.Text.Replace(Me.Buffer, Me.TxtDocumentNo.Text)
                    End If
                End If
            Else
                If Me.Buffer <> "-1" Then
                    JournalDesc = Me.TxtDocumentNo.Text & "-" & Me.ObjKaryawan.GetCodeByPrimary(Me.Buffer) + "-" & Me.ObjKaryawan.GetFullNameByPrimary(Me.Buffer)
                Else
                    JournalDesc = Me.TxtDocumentNo.Text
                End If

                If Me.txtJournalDesc.Text = JournalDesc Then
                    If Me.TxtPrimaryEmployee.Text <> "-1" Then
                        Me.txtJournalDesc.Text = Me.TxtDocumentNo.Text & "-" & Me.TxtEmployeeCode.Text + "-" & Me.TxtEmployeeName.Text
                    Else
                        Me.txtJournalDesc.Text = Me.TxtDocumentNo.Text
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = True Then
                    Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
                Else
                    If Me.dtpDate.Value.Month <> Me.ObjLoan.Date_.Month Or _
                    Me.dtpDate.Value.Year <> Me.ObjLoan.Date_.Year Then
                        Me.TxtDocumentNo.Text = Me.ObjLoan.GetNoBaru(Me.dtpDate.Value)
                    Else
                        Me.TxtDocumentNo.Text = Me.ObjLoan.GetDocNoByPrimary(Me.TxtPrimaryLoan.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class