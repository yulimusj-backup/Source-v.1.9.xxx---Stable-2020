Option Explicit On
Option Strict Off
Imports System.Windows.Forms, System.Drawing
Public Class FrmTransaksiPayment
    Dim mFrmCetak As New FrmTransaksiPayment_Cetak
    Dim mFrmCetakOption As New FrmTransaksiPayment_Cetak_Option
    Dim ObjPosition As New Core_Module.ClsMasterJobPosition
    Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim Baru As Boolean
    Dim Buffer As String = ""
    Dim BufferTransPaymentDate As Date = Now
    Dim EraseRowIndex_PayAccount As Integer = 0
    Dim EraseRowIndex_Item As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Obj As New ClsTrans_Payment
    Dim ObjItem As New ClsTrans_Payment_Item
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjJurnal As New AccModule.ClsTransaksiJournal
    Dim ObjSalary As New ClsTransaksiSalary
    Dim TBLItemTax As New DseTransaksiSalary.tbltrans_salary_itemtaxDataTable

    Dim mBufferIndexSelectionDGVSalary As String = "-1"
    Dim mIndexSelectionDGVSalary As Double = 0
    Public mCalledCellEndEdit As Boolean = False
    'GetStatusPayment 
    Dim TBLPaid As New DseTransaksiSalary.tbltrans_salary2DataTable

    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public mTBLItem_Employee_d As New DseTransPayment.tbltrans_payment_item_employee_dDataTable
    Public Def_Prec As Double = ObjCurr.GetDecimal_DefaultCurrency()
    Public PrimaryCurrency As Double = ObjCurr.GetPrimary("IDR")
    Public mDipanggilDariCalulateSalary As Boolean = False
    Public mPrimarySalary As Double = -1
    Public mTanggalSalary As Date = #1/1/1950#
    Public mPrimaryPaymentDariSalary As Double = -1

    'Data table untuk DGV
    Public TBLSalary As New DseTransPayment.tbltrans_salaryDataTable
    Dim RowSalary As DseTransPayment.tbltrans_salaryRow
    Dim TBLItemSalary_Income As New DseTransPayment.tbltrans_salary_itemDataTable
    Dim TBLItemSalary_Deduction As New DseTransPayment.tbltrans_salary_itemDataTable
    Dim TBLItemSalary_Other As New DseTransPayment.tbltrans_salary_item_otherdeductionDataTable
    Dim TBLLoanPayment As New DseTransPayment.TBLTransLoanPaymentDataTable
    Dim TBLLoanPaymentItem As New DseTransPayment.TBLTransLoanPaymentItemDataTable
    Dim TBLHasil As New DseTransPayment.DataTable1DataTable

    Dim objJournal As New AccModule.ClsTransaksiJournal
    Dim ObjLoan As New ClsTransaksiLoan

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Try
            ''rosadin test
            'Dim rowDgv As New Windows.Forms.DataGridViewRow
            'Me.DGVJournalItem.Rows.Clear()
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Salary Expense"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "8,500,000"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "-"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "OT Expense"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "750,000"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "-"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Loss Acc"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "10,000"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "-"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "BCA"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "8,235,000"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "KK"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "500,000"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Pajak Acc"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "75,000"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Kop Acc"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "25,000"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Gain Acc"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "400,000"
            'Me.DGVJournalItem.Rows.Add()
            'rowDgv = Me.DGVJournalItem.Rows(Me.DGVJournalItem.Rows.Count - 1)
            'rowDgv.Cells(Me.ColJurnal_Account.Index).Value = "Receivable Acc"
            'rowDgv.Cells(Me.ColJurnal_Debet.Index).Value = "-"
            'rowDgv.Cells(Me.ColJurnal_Kredit.Index).Value = "25,000"

            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.DGVJournalItem.Rows.Count > 0 Then
                    Me.mFrmCetakOption.mPrimaryJournal = Me.TxtPrimaryJournal.Text
                    Me.GetNilaiPrimarySalary(PrimaryMin, PrimaryMax)
                    Me.mFrmCetakOption.mPrimarySalaryMin = PrimaryMin
                    Me.mFrmCetakOption.mPrimarySalaryMax = PrimaryMax
                    Me.mFrmCetakOption.mNoSalary = Me.TxtDocumentNo.Text
                    Me.mFrmCetakOption.mTanggalTransaksi = Me.DtpDate.Value
                    Me.mFrmCetakOption.mPeriodFrom = Me.dtpPeriod.Value
                    Me.mFrmCetakOption.mPeriodTo = Me.DtpTo.Value
                    Me.mFrmCetakOption.mDefPrec = Me.Def_Prec
                    Me.mFrmCetakOption.mFrmTransPayment = Me
                    Me.mFrmCetakOption.ShowDialog()
                End If
            End If

            If Me.TxtRemarks.Enabled = False Then
                If Me.DGVPayment.AllowUserToAddRows = True Then
                    Me.DGVPayment.AllowUserToAddRows = False
                End If
                If Me.DGVJournalItem.AllowUserToAddRows = True Then
                    Me.DGVJournalItem.AllowUserToAddRows = False
                End If
                If Me.dgvItem.AllowUserToAddRows = True Then
                    Me.dgvItem.AllowUserToAddRows = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Obj = New ClsTrans_Payment
            Me.Baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvItem.AllowUserToAddRows = False
            Me.DGVPayment.AllowUserToAddRows = False
            Me.DGVJournalItem.AllowUserToAddRows = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVPayment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVJournalItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.LblAttendance.Text = "Hide Detail"
            Me.SplitContainer4.Panel2Collapsed = False
            If Me.mDipanggilDariCalulateSalary = True Then
                'Cek Status Attendance Paid atau Unpaid
                If Me.ObjSalary.GetPaymentStatus(Me.mPrimarySalary) = ClsTransaksiSalary.EnumPaymentStatus.paid Then
                    Me.btnFind.PerformClick()
                    Me.btnEdit.PerformClick()

                Else
                    Me.btnNew.PerformClick()
                    Me.IsiDGVItemByTransactionCalculateSalary(Me.dgvItem, Me.Baru)
                End If
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.GroupBox2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToAddRows = True
            Me.DGVJournalItem.ReadOnly = True
            Me.DGVJournalItem.AllowUserToDeleteRows = False
            Me.DGVJournalItem.AllowUserToAddRows = False
            Me.DGVPayment.ReadOnly = False
            Me.DGVPayment.AllowUserToDeleteRows = True
            Me.DGVPayment.AllowUserToAddRows = True

            'Label
            Me.Lbl_Deduction.Enabled = True
            Me.Lbl_Income.Enabled = True
            Me.Lbl_Loan_Installment.Enabled = True
            Me.Lbl_Other_Deduction.Enabled = True
            Me.LblIncomeTax.Enabled = True

            Me.TabControl1.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToAddRows = False
            Me.DGVPayment.ReadOnly = True
            Me.DGVPayment.AllowUserToDeleteRows = False
            Me.DGVPayment.AllowUserToAddRows = False
            Me.DGVJournalItem.ReadOnly = True
            Me.DGVJournalItem.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToAddRows = False

            'Label
            If Me.dgvItem.Rows.Count > 0 Then
                Me.Lbl_Deduction.Enabled = True
                Me.Lbl_Income.Enabled = True
                Me.Lbl_Loan_Installment.Enabled = True
                Me.Lbl_Other_Deduction.Enabled = True
                Me.LblIncomeTax.Enabled = True
            Else
                Me.Lbl_Deduction.Enabled = False
                Me.Lbl_Income.Enabled = False
                Me.Lbl_Loan_Installment.Enabled = False
                Me.Lbl_Other_Deduction.Enabled = False
                Me.LblIncomeTax.Enabled = False
            End If

            'Me.KosongkanFieldItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.GroupBox2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.GroupBox1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.GroupBox2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryJournal.Text = "-1"
            Me.TxtIncome.Text = "0"
            Me.TxtDeduction.Text = "0"
            Me.TxtGrossSalary.Text = "0"
            Me.TxtIncomeTax.Text = "0"
            Me.TxtLoanInstallment.Text = "0"
            Me.TxtOtherDedution.Text = "0"
            Me.TxtSalaryReceived.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            'Me.KosongkanFieldItem()
            Me.dgvItem.Rows.Clear()
            Me.DGVPayment.Rows.Clear()
            Me.DGVJournalItem.Rows.Clear()
            Me.mTBLItem_Employee_d.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim RowDGV As DataGridViewRow
        Dim Debet, Credit As Decimal
        Dim DocNoBuffer As String = ""

        Try
            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.TxtDocumentNo)
            KoleksiNamaData.Add("Journal No")
            KoleksiTextBox.Add(Me.txtJournalNo)
            KoleksiNamaData.Add("Journal Description")
            KoleksiTextBox.Add(Me.TxtJournalDesc)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. harus ada data item
            If Me.dgvItem.Rows.Count <= 1 Then
                MsgBox("Please specify at least 1 employee.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek perubahan nomor transaksi
            If Me.Baru = False Then
                If Me.Obj.TanggalTransaksi.Month = Me.DtpDate.Value.Month And _
                Me.Obj.TanggalTransaksi.Year = Me.DtpDate.Value.Year Then
                    Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                Else
                    Dim NoTransaksiLama As String = Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    If NoTransaksiLama <> Me.TxtDocumentNo.Text Then
                        If Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = False Then
                            If MsgBox("You just changed transaction date, this will change your transaction " _
                                      + "number and there's a possibility your transaction no cannot be change back." _
                                      + "Continue and accept this change?", _
                                      MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.Obj.TanggalTransaksi
                                Me.TxtDocumentNo.Text = Me.Obj.DocumentNo
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                    DocNoBuffer = Me.TxtDocumentNo.Text
                    Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtDocumentNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtDocumentNo.Text) = True Then
                        DocNoBuffer = Me.TxtDocumentNo.Text
                        Me.TxtDocumentNo.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, _
                        Me.Def_Prec, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. Remarks
            If Me.ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtRemarks.Focus()
                Return False
            End If

            '4. Batasan jika jurnal telah diclosing 
            'jika sudah ada closing jurnal untuk tahun ybs maka tidak bisa input data baru
            If Me.ObjJurnal.GetJml(AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, ObjCD.GetFisrtDateOfTheYear(Me.DtpDate.Value), _
            ObjCD.GetLastDateOfTheYear(Me.DtpDate.Value)) > 0 Then
                MsgBox("There are closing jurnal for this year you cannot save this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. Batasan Take home pay harus sama dengan jumlah bayar atau payment
            If Me.Obj.ApakahDataPembayaranAmountNyaSesuaiDenganTakeHomePay(Me.ObjCD, Me.dgvItem, Me.mTBLItem_Employee_d, Me) = False Then
                Return False
            End If

            '6. Tanggal transaksi Payment, harus lebih besar sama dengan period to
            'If Me.DtpDate.Value.Date < Me.DtpTo.Value.Date Then
            '    MsgBox("You cannot input transaction date less than period to.", MsgBoxStyle.Information, "Information")
            '    Me.DtpTo.Focus()
            '    Return False
            'End If
            If Me.dtpPeriod.Value > Me.DtpTo.Value Then
                MsgBox("You cannot input period from more than period to.", MsgBoxStyle.Information, "Information")
                Me.dtpPeriod.Focus()
                Return False
            End If

            '7. cek batasan tanggal ke dalam DGVItem 
            Me.dgvItem.AllowUserToAddRows = False
            Dim DateMax As Date = #1/1/1950#
            If Me.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In Me.dgvItem.Rows
                    If CDate(RowDGV.Cells(Me.ColEmp_SalaryDate.Index).Value) > DateMax Then
                        DateMax = CDate(RowDGV.Cells(Me.ColEmp_SalaryDate.Index).Value)
                    End If
                Next
            End If
            If Me.DtpTo.Value < DateMax Then
                MsgBox("You cannot input period to less than maximum salary date.", MsgBoxStyle.Information, "Information")
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If
            Me.dgvItem.AllowUserToAddRows = True

            '8. unbalance journal
            Debet = 0
            Credit = 0
            For Each RowDGV In Me.DGVJournalItem.Rows
                Debet += CDec(RowDGV.Cells(Me.ColJurnal_Debet.Index).Value)
                Credit += CDec(RowDGV.Cells(Me.ColJurnal_Kredit.Index).Value)
            Next
            If Debet <> Credit Then
                MsgBox("Logic error: unbalance journal, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.Baru And Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                Me.txtJournalNo.Focus()
                Return False
            ElseIf Me.Baru = False And Me.txtJournalNo.Text <> Me.objJournal.GetRefCode(Me.Obj.GetPrimaryJurnalRef(Me.TxtPrimary.Text)) And _
                    Me.ObjLoan.GetJmlByJournalNo(Me.txtJournalNo.Text) > 0 Then
                MsgBox("Journal No: " & Me.txtJournalNo.Text & " is already exist.", MsgBoxStyle.Information, "Information")
                Me.txtJournalNo.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Public Sub SalinKeObjek()
        Try
            With Me.Obj
                .DocumentNo = Me.TxtDocumentNo.Text
                .TanggalTransaksi = Me.DtpDate.Value
                .StartDate = Me.dtpPeriod.Value
                .EndDate = Me.DtpTo.Value
                .PrimaryJurnalRef = Me.TxtPrimaryJournal.Text
                .Remarks = Me.TxtRemarks.Text
                .JournalNo = Me.txtJournalNo.Text
                .JournalDesc = Me.TxtJournalDesc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocumentNo.Text = .DocumentNo
                Me.TxtPrimaryJournal.Text = .PrimaryJurnalRef
                Me.txtJournalNo.Text = .JournalNo  'Me.objJournal.GetRefCode(Me.TxtPrimaryJurnal.Text)
                Me.TxtJournalDesc.Text = .JournalDesc 'Me.objJournal.GetDescription(Me.TxtPrimaryJurnal.Text)
                Me.DtpDate.Value = .TanggalTransaksi
                Me.dtpPeriod.Value = .StartDate
                Me.DtpTo.Value = .EndDate
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Dim Debet, Credit As Decimal
        Dim RowDGV As DataGridViewRow

        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.GetNilaiPrimarySalary(PrimaryMin, PrimaryMax)
                        Me.Obj.BuatJournalViewPayment(PrimaryMin, PrimaryMax, Me.mTBLItem_Employee_d, Me.TBLHasil, Me.Def_Prec, Me)
                        Debet = 0
                        Credit = 0
                        For Each RowDGV In Me.DGVJournalItem.Rows
                            Debet += CDec(RowDGV.Cells(Me.ColJurnal_Debet.Index).Value)
                            Credit += CDec(RowDGV.Cells(Me.ColJurnal_Kredit.Index).Value)
                        Next
                        If Debet <> Credit Then
                            MsgBox("Logic error: unbalance journal, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                        Me.DGVPayment, Me.DGVJournalItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.ObjItem.IsiDGVPayment(Me.TxtPrimary.Text, Me.DGVPayment, Me.dgvItem, Me.Def_Prec, Me.dgvItem.SelectedCells(0).RowIndex, Me)
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            If Me.mDipanggilDariCalulateSalary = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.GetNilaiPrimarySalary(PrimaryMin, PrimaryMax)
                        Me.Obj.BuatJournalViewPayment(PrimaryMin, PrimaryMax, Me.mTBLItem_Employee_d, Me.TBLHasil, Me.Def_Prec, Me)
                        Debet = 0
                        Credit = 0
                        For Each RowDGV In Me.DGVJournalItem.Rows
                            Debet += CDec(RowDGV.Cells(Me.ColJurnal_Debet.Index).Value)
                            Credit += CDec(RowDGV.Cells(Me.ColJurnal_Kredit.Index).Value)
                        Next
                        If Debet <> Credit Then
                            MsgBox("Logic error: unbalance journal, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        'Simpan Perubahan
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                        Me.DGVPayment, Me.DGVJournalItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.ObjItem.IsiDGVPayment(Me.TxtPrimary.Text, Me.DGVPayment, Me.dgvItem, Me.Def_Prec, Me.dgvItem.SelectedCells(0).RowIndex, Me)
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            If Me.mDipanggilDariCalulateSalary = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtDocumentNo.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.mDipanggilDariCalulateSalary = True Then
                        Me.Close()
                    Else
                        If Baru = True Then
                            Me.KosongkanField()
                            Me.TxtDocumentNo.Clear()
                        Else
                            Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, _
                            Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me.Def_Prec, Me)
                            Me.SalinKeField()
                        End If
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim Primary_Akun As Double
        Dim PrimaryAkunStr As String = ""

        Try
            'pastikan tax payable account dan tax expense account <> -1
            PrimaryAkunStr = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxPayableAccount1)
            If PrimaryAkunStr = "" Then
                MsgBox("Please specify tax payable account.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                Primary_Akun = PrimaryAkunStr
                If Primary_Akun = -1 Then
                    MsgBox("Please specify tax payable account.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            PrimaryAkunStr = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryTaxExpenseAccount)
            If PrimaryAkunStr = "" Then
                MsgBox("Please specify tax expense account.", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                Primary_Akun = PrimaryAkunStr
                If Primary_Akun = -1 Then
                    MsgBox("Please specify tax expense account.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.Baru = True
            Obj.ResetObject(Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem)
            Me.AktifkanDetail()
            Me.KosongkanField()
            If Me.mDipanggilDariCalulateSalary = True Then
                Me.DtpDate.Value = Me.mTanggalSalary
                Me.dtpPeriod.Value = Me.mTanggalSalary
                Me.DtpTo.Value = Me.mTanggalSalary
            End If
            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.mTBLItem_Employee_d.Rows.Clear()
            Me.TabControl1.SelectedIndex = 0
            Me.DtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmTransaksiPayment_Cari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            If Me.mDipanggilDariCalulateSalary = False Then
                ijk.ShowDialog()
                If Not ijk.mPrimary = -1 Then
                    If Me.Obj.ApakahBolehFindDataPayment(ijk.mPrimary, Me.UserCode) = True Then
                        'Bukan dari transaksi Salary
                        If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, _
                        Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me.Def_Prec, Me) = True Then
                            Me.SalinKeField()
                            If Me.TxtRemarks.Enabled = False Then
                                If Me.DGVPayment.AllowUserToAddRows = True Then
                                    Me.DGVPayment.AllowUserToAddRows = False
                                End If
                            End If
                            Me.btnEdit.Focus()
                        Else
                            Me.btnFind.Focus()
                        End If
                        Me.btnFind.Focus()
                    Else
                        Exit Sub
                    End If
                End If
                'Label
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.Lbl_Deduction.Enabled = True
                    Me.Lbl_Income.Enabled = True
                    Me.Lbl_Loan_Installment.Enabled = True
                    Me.Lbl_Other_Deduction.Enabled = True
                    Me.LblIncomeTax.Enabled = True
                Else
                    Me.Lbl_Deduction.Enabled = False
                    Me.Lbl_Income.Enabled = False
                    Me.Lbl_Loan_Installment.Enabled = False
                    Me.Lbl_Other_Deduction.Enabled = False
                    Me.LblIncomeTax.Enabled = False
                End If
            Else
                If Not Me.mPrimaryPaymentDariSalary = -1 Then
                    If Me.Obj.ApakahBolehFindDataPayment(ijk.mPrimary, Me.UserCode) = True Then
                        'Dari Transaksi Salary
                        If Me.Obj.CariDetail(Me.mPrimaryPaymentDariSalary, Me.ObjCD, _
                        Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me.Def_Prec, Me) = True Then
                            Me.SalinKeField()
                            If Me.TxtRemarks.Enabled = False Then
                                If Me.DGVPayment.AllowUserToAddRows = True Then
                                    Me.DGVPayment.AllowUserToAddRows = False
                                End If
                            End If
                            Me.btnEdit.Focus()
                        Else
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True, Me.dgvItem, Me) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    Me.DtpDate.Focus()
                    Me.TabControl1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDocumentNo.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, Me.dgvItem, _
                Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, _
                        Me.DGVPayment, Me.DGVJournalItem, Me, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            If Me.DGVPayment.AllowUserToAddRows = True Then
                                Me.DGVPayment.AllowUserToAddRows = False
                            End If
                            If Me.dgvItem.AllowUserToAddRows = True Then
                                Me.dgvItem.AllowUserToAddRows = False
                            End If
                            If Me.DGVJournalItem.AllowUserToAddRows = True Then
                                Me.DGVJournalItem.AllowUserToAddRows = False
                            End If
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
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.mDipanggilDariCalulateSalary = False Then
                If Me.TxtRemarks.Enabled = True Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If Baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                                Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me) = True Then
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
                                If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                                Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me) = True Then
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Enter
        Try
            Me.BufferTransPaymentDate = Me.DtpDate.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.BufferTransPaymentDate <> Me.DtpDate.Value.Date Then
                    Me.Buffer = Me.TxtDocumentNo.Text
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            If Me.TxtRemarks.Enabled = True Then
                If Me.Baru = True Then
                    Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                Else
                    If Me.DtpDate.Value.Month = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                    Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                        Me.TxtDocumentNo.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    Else
                        If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                        Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                            Me.TxtDocumentNo.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD)
                        Else
                            Me.TxtDocumentNo.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim PrimaryAkun As String = ""
        Try
            If Me.DtpTo.Value.Date < Me.dtpPeriod.Value.Date Then
                MsgBox("Period cannot larger than period to.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            PrimaryAkun = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
            If String.IsNullOrEmpty(PrimaryAkun) = True Or PrimaryAkun = "-1" Then
                MsgBox("Please setting cash account salary in global setting.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnHitung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHitung.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TabControl1.SelectedIndex = 1 Then
                Me.TabControl1.SelectedIndex = 0
            End If
            hasil = MsgBox("Do you want to calculate transaction payment ?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                If Me.ApakahDataItemBenar() = True Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    'Kumpulkan semua data item
                    Me.Obj.FungsiBtnHitungPayment_KumpulkanData(Me.TBLSalary, Me.mTBLItem_Employee_d, Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me)
                    'Me.Obj.FungsiBtnHitungPayment_KumpulkanData(Me.TBLSalary, Me.TBLItemSalary_Income, Me.TBLItemSalary_Deduction, _
                    'Me.TBLItemSalary_Other, Me.TBLLoanPayment, Me.TBLLoanPaymentItem, Me.dgvItem, Me.DGVPayment, Me.DGVJournalItem, Me)
                    'Input DGV Salary
                    Me.Obj.FungsiBtnHitungPayment_IsiUlangDGVSalary(Me.TBLSalary, Me.dgvItem, Me.Def_Prec, Me.Baru, Me.TxtPrimary.Text, Me)
                    'Pengaturan selection
                    Me.dgvItem.AllowUserToAddRows = False
                    If Me.dgvItem.Rows.Count > 0 Then
                        If Me.dgvItem.SelectedCells.Count > 0 Then
                            Me.mBufferIndexSelectionDGVSalary = -1
                            Me.mIndexSelectionDGVSalary = Me.dgvItem.SelectedCells(0).RowIndex
                            'Masuk in value pada masing-masing textbox
                            Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, Me.mIndexSelectionDGVSalary, Me.dgvItem, Me.Def_Prec, Me)
                            'Masuk in payment item
                            Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountDefault(Me.ObjSetting, _
                            Me.ObjCD, Me.DGVPayment, Me.Def_Prec, Me.mIndexSelectionDGVSalary, Me)
                        Else
                            Me.TxtDeduction.Text = "0"
                            Me.TxtIncome.Text = "0"
                            Me.TxtGrossSalary.Text = "0"
                            Me.TxtLoanInstallment.Text = "0"
                            Me.TxtIncomeTax.Text = "0"
                            Me.TxtOtherDedution.Text = "0"
                            Me.TxtSalaryReceived.Text = "0"
                            Me.DGVPayment.Rows.Clear()
                            Me.DGVJournalItem.Rows.Clear()
                            'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                            Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.Rows.Count - 1
                        End If
                    End If
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEmp_SalaryDoc.Index Or e.ColumnIndex = Me.ColEmp_Remarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanDetailPayment()
        Try
            Me.TxtDeduction.Text = "0"
            Me.TxtIncome.Text = "0"
            Me.TxtGrossSalary.Text = "0"
            Me.TxtLoanInstallment.Text = "0"
            Me.TxtIncomeTax.Text = "0"
            Me.TxtOtherDedution.Text = "0"
            Me.TxtSalaryReceived.Text = "0"
            Me.DGVPayment.Rows.Clear()
            Me.DGVJournalItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailInfo(ByVal vRowIndex As Integer)
        Try
            If IsNothing(Me.dgvItem.Item(1, vRowIndex).Value) = False Then
                If Me.dgvItem.Rows.Count > 0 Then
                    If Me.dgvItem.SelectedCells.Count > 0 Then
                        Me.mIndexSelectionDGVSalary = Me.dgvItem.SelectedCells(0).RowIndex
                        If Me.mBufferIndexSelectionDGVSalary <> Me.mIndexSelectionDGVSalary Then
                            'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                            Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.SelectedCells(0).RowIndex.ToString
                            'Masuk in value pada masing-masing textbox
                            Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, Me.mIndexSelectionDGVSalary, Me.dgvItem, Me.Def_Prec, Me)
                            'Masuk in Payment Item Selection Changed
                            Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged(Me.DGVPayment, Me.mIndexSelectionDGVSalary, Me.ObjCD, Me.Def_Prec, Me)
                        End If
                    End If
                End If
            Else
                Me.KosongkanDetailPayment()
                'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.Rows.Count - 1
            End If
            If Me.TxtRemarks.Enabled = False Then
                Me.DGVPayment.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If
            Me.ShowDetailInfo(e.RowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimarySalary As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value = vPrimarySalary Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub IsiDGVItemByTransactionCalculateSalary(ByVal vDGVItem As DataGridView, ByVal vBaru As Boolean)
        Dim RowDGV As DataGridViewRow
        Dim IdentifierStr As String
        Dim CboDGV As DataGridViewComboBoxCell
        'Tabel Adapter untuk transaksi salary
        Dim DTASalary_DGV As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary_DGV As New DseTransPayment.tbltrans_salaryDataTable
        Dim RowSalary_DGV As DseTransPayment.tbltrans_salaryRow
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow
        Try
            If vBaru = True Then
                vDGVItem.Rows.Clear()
                vDGVItem.AllowUserToAddRows = False
                DTASalary_DGV.FillBy_PrimarySalary(TBLSalary_DGV, Me.mPrimarySalary)
                vDGVItem.Rows.Add()
                RowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                If TBLSalary_DGV.Rows.Count > 0 Then
                    RowSalary_DGV = TBLSalary_DGV.Rows(0)
                    Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, RowDGV.Index).Value = Me.mPrimarySalary
                    Me.dgvItem.Item(Me.ColEmp_SalaryDoc.Index, RowDGV.Index).Value = Me.ObjSalary.GetNoTransaksi(Me.mPrimarySalary)
                    Me.dgvItem.Item(Me.ColEmp_PrimaryEmployee.Index, RowDGV.Index).Value = RowSalary_DGV.Primary_Employee
                    IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.dgvItem, Me.ColIdentifier.Index, Me.Baru, Me.TxtPrimary.Text, RowDGV.Index)
                    Me.dgvItem.Item(Me.ColIdentifier.Index, RowDGV.Index).Value = IdentifierStr
                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColEmp_Remarks.Index, RowDGV.Index).Value) = True Then
                        Me.dgvItem.Item(Me.ColEmp_Remarks.Index, RowDGV.Index).Value = ""
                    Else
                        Me.dgvItem.Item(Me.ColEmp_Remarks.Index, RowDGV.Index).Value = RowSalary_DGV.Remarks
                    End If
                    Me.dgvItem.Item(Me.ColEmp_SalaryDate.Index, RowDGV.Index).Value = Format(RowSalary_DGV.Tanggal_Transaksi, "dd MMM yyyy")
                    Me.dgvItem.Item(Me.ColEmp_TakeHomePay.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Salary_Received, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_OtherDeduction.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Other_Deduction, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_LoanInstallment.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Loan_Installment, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_IncomeTax.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income_Tax, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_GrossSalary.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Gross_Salary, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_Deduction.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Deduction, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_Income.Index, RowDGV.Index).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColEmp_EmployeeCode.Index, RowDGV.Index).Value = RowSalary_DGV.Employee_Code
                    Me.dgvItem.Item(Me.ColEmp_Name.Index, RowDGV.Index).Value = RowSalary_DGV.Employee_Name
                    CboDGV = Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, RowDGV.Index)
                    Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, RowDGV.Index).Value = CboDGV.Items.Item(RowSalary_DGV.Tax)
                    Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome_Enum.Index, RowDGV.Index).Value = RowSalary_DGV.Tax
                    'Insert mTBLTrans_Payment_Employee_d
                    RowItem = Me.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                    RowItem.primary = -1
                    RowItem.primary_transaksi = CDbl(Me.TxtPrimary.Text)
                    RowItem.identifier = RowDGV.Cells(Me.ColIdentifier.Index).Value
                    RowItem.primary_akun_payment = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
                    RowItem.akun_payment_code = Me.ObjAkun.GetAccountCode(RowItem.primary_akun_payment)
                    RowItem.akun_payment_desc = Me.ObjAkun.GetDescription_ByPrimaryAcc(RowItem.primary_akun_payment)
                    RowItem.amount = Me.ObjCD.DigitGrouping(CDec(RowDGV.Cells(Me.ColEmp_TakeHomePay.Index).Value), Me.Def_Prec)
                    RowItem.remarks = ""
                    RowItem.Index_Row_Salary = RowDGV.Index
                    Me.mTBLItem_Employee_d.Rows.Add(RowItem)
                    'Insert Tbl Salary
                    If TBLSalary_DGV.Rows.Count > 0 Then
                        RowSalary_DGV = TBLSalary_DGV.Rows(0)
                        RowSalary = Me.TBLSalary.Newtbltrans_salaryRow
                        For ikl As Integer = 0 To Me.TBLSalary.Columns.Count - 1
                            RowSalary.Item(ikl) = RowSalary_DGV.Item(ikl)
                        Next
                        Me.TBLSalary.Rows.Add(RowSalary)
                    End If
                    'Event Selection Changed dipanggil ketika CellEndEdit
                    Me.mCalledCellEndEdit = True
                    If Me.mCalledCellEndEdit = True Then
                        'Display DGVPayment
                        Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged( _
                        Me.DGVPayment, RowDGV.Index, Me.ObjCD, Me.Def_Prec, Me)
                        'Display TextBox Hitungan
                        Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, _
                        RowDGV.Index, Me.dgvItem, Me.Def_Prec, Me)
                    End If
                End If
                Me.ObjCD.SetNumber(Me.dgvItem)
                vDGVItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary_DGV.Dispose()
            TBLSalary_DGV.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim Index As Integer
        Dim IdentifierStr As String
        Dim CboDGV As DataGridViewComboBoxCell
        'Tabel Adapter untuk transaksi salary
        Dim DTASalary_DGV As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary_DGV As New DseTransPayment.tbltrans_salaryDataTable
        Dim RowSalary_DGV As DseTransPayment.tbltrans_salaryRow
        'Row Untuk mTBLTrans_Payment_Employee_d
        Dim RowItem As DseTransPayment.tbltrans_payment_item_employee_dRow

        Try
            DTASalary_DGV.Connection.ConnectionString = Core_Module.GetConstring
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColEmp_SalaryDoc.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjSalary.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput(Me.ObjSalary.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value = Me.ObjSalary.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                DTASalary_DGV.FillBy_PrimarySalary(TBLSalary_DGV, CDec(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value))
                                If TBLSalary_DGV.Rows.Count > 0 Then
                                    'Batasan Payroll -- baru
                                    If Me.Obj.ApakahBolehInputSalary(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value, Me) = False Then
                                        Me.dgvItem.Item(Me.ColEmp_SalaryDoc.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Item = e.RowIndex
                                            Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                    RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                    Me.dgvItem.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = RowSalary_DGV.Primary_Employee
                                    IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.dgvItem, Me.ColIdentifier.Index, Me.Baru, Me.TxtPrimary.Text, Row.Index)
                                    Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value = IdentifierStr
                                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value) = True Then
                                        Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = ""
                                    Else
                                        Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = RowSalary_DGV.Remarks
                                    End If
                                    Me.dgvItem.Item(Me.ColEmp_SalaryDate.Index, e.RowIndex).Value = Format(RowSalary_DGV.Tanggal_Transaksi, "dd MMM yyyy")
                                    Me.dgvItem.Item(Me.ColEmp_TakeHomePay.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Salary_Received, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_OtherDeduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Other_Deduction, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_LoanInstallment.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Loan_Installment, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_IncomeTax.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income_Tax, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_GrossSalary.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Gross_Salary, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_Deduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Deduction, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_Income.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income, Me.Def_Prec)
                                    Me.dgvItem.Item(Me.ColEmp_EmployeeCode.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Code
                                    Me.dgvItem.Item(Me.ColEmp_Name.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Name
                                    CboDGV = Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex)
                                    Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex).Value = CboDGV.Items.Item(RowSalary_DGV.Tax)
                                    Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome_Enum.Index, e.RowIndex).Value = RowSalary_DGV.Tax
                                    'Insert mTBLTrans_Payment_Employee_d
                                    RowItem = Me.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                                    RowItem.primary = -1
                                    RowItem.primary_transaksi = CDbl(Me.TxtPrimary.Text)
                                    RowItem.identifier = Row.Cells(Me.ColIdentifier.Index).Value
                                    RowItem.primary_akun_payment = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
                                    RowItem.akun_payment_code = Me.ObjAkun.GetAccountCode(RowItem.primary_akun_payment)
                                    RowItem.akun_payment_desc = Me.ObjAkun.GetDescription_ByPrimaryAcc(RowItem.primary_akun_payment)
                                    RowItem.amount = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColEmp_TakeHomePay.Index).Value), Me.Def_Prec)
                                    RowItem.remarks = ""
                                    RowItem.Index_Row_Salary = Row.Index
                                    Me.mTBLItem_Employee_d.Rows.Add(RowItem)
                                    'Insert Tbl Salary
                                    If TBLSalary_DGV.Rows.Count > 0 Then
                                        RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                        RowSalary = Me.TBLSalary.Newtbltrans_salaryRow
                                        For ikl As Integer = 0 To Me.TBLSalary.Columns.Count - 1
                                            RowSalary.Item(ikl) = RowSalary_DGV.Item(ikl)
                                        Next
                                        Me.TBLSalary.Rows.Add(RowSalary)
                                    End If
                                    'Event Selection Changed dipanggil ketika CellEndEdit
                                    Me.mCalledCellEndEdit = True
                                    If Me.mCalledCellEndEdit = True Then
                                        'Display DGVPayment
                                        Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged( _
                                        Me.DGVPayment, Row.Index, Me.ObjCD, Me.Def_Prec, Me)
                                        'Display TextBox Hitungan
                                        Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, _
                                        Row.Index, Me.dgvItem, Me.Def_Prec, Me)
                                    End If
                                End If
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication salary number, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColEmp_SalaryDoc.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Item = e.RowIndex
                                        Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New HR.FrmTrans_Gaji_Cari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.mPrimary)
                                If Index < 0 Then
                                    Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value = ij.mPrimary
                                    DTASalary_DGV.FillBy_PrimarySalary(TBLSalary_DGV, CDec(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value))
                                    If TBLSalary_DGV.Rows.Count > 0 Then
                                        'Batasan Payroll -- baru
                                        If Me.Obj.ApakahBolehInputSalary(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value, Me) = False Then
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex_Item = e.RowIndex
                                                Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                            Else
                                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.ObjCD.SetNumber(Me.dgvItem)
                                            Exit Sub
                                        End If
                                        RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                        Me.dgvItem.Item(Me.ColEmp_SalaryDoc.Index, e.RowIndex).Value = RowSalary_DGV.No_Transaksi
                                        Me.dgvItem.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = RowSalary_DGV.Primary_Employee
                                        IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.dgvItem, Me.ColIdentifier.Index, Me.Baru, Me.TxtPrimary.Text, Row.Index)
                                        Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value = IdentifierStr
                                        If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value) = True Then
                                            Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = ""
                                        Else
                                            Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = RowSalary_DGV.Remarks
                                        End If
                                        Me.dgvItem.Item(Me.ColEmp_SalaryDate.Index, e.RowIndex).Value = Format(RowSalary_DGV.Tanggal_Transaksi, "dd MMM yyyy")
                                        Me.dgvItem.Item(Me.ColEmp_TakeHomePay.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Salary_Received, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_OtherDeduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Other_Deduction, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_LoanInstallment.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Loan_Installment, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_IncomeTax.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income_Tax, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_GrossSalary.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Gross_Salary, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_Deduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Deduction, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_Income.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income, Me.Def_Prec)
                                        Me.dgvItem.Item(Me.ColEmp_EmployeeCode.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Code
                                        Me.dgvItem.Item(Me.ColEmp_Name.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Name
                                        CboDGV = Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex)
                                        Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex).Value = CboDGV.Items.Item(RowSalary_DGV.Tax)
                                        Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome_Enum.Index, e.RowIndex).Value = RowSalary_DGV.Tax
                                        'Insert mTBLTrans_Payment_Employee_d
                                        RowItem = Me.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                                        RowItem.primary = -1
                                        RowItem.primary_transaksi = CDbl(Me.TxtPrimary.Text)
                                        RowItem.identifier = Row.Cells(Me.ColIdentifier.Index).Value
                                        RowItem.primary_akun_payment = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                        Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
                                        RowItem.akun_payment_code = Me.ObjAkun.GetAccountCode(RowItem.primary_akun_payment)
                                        RowItem.akun_payment_desc = Me.ObjAkun.GetDescription_ByPrimaryAcc(RowItem.primary_akun_payment)
                                        RowItem.amount = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColEmp_TakeHomePay.Index).Value), Me.Def_Prec)
                                        RowItem.remarks = ""
                                        RowItem.Index_Row_Salary = Row.Index
                                        Me.mTBLItem_Employee_d.Rows.Add(RowItem)
                                        'Insert Tbl Salary
                                        If TBLSalary_DGV.Rows.Count > 0 Then
                                            RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                            RowSalary = Me.TBLSalary.Newtbltrans_salaryRow
                                            For ikl As Integer = 0 To Me.TBLSalary.Columns.Count - 1
                                                RowSalary.Item(ikl) = RowSalary_DGV.Item(ikl)
                                            Next
                                            Me.TBLSalary.Rows.Add(RowSalary)
                                        End If
                                        'Event Selection Changed dipanggil ketika CellEndEdit
                                        Me.mCalledCellEndEdit = True
                                        If Me.mCalledCellEndEdit = True Then
                                            'Display DGVPayment
                                            Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged( _
                                            Me.DGVPayment, Row.Index, Me.ObjCD, Me.Def_Prec, Me)
                                            'Display TextBox Hitungan
                                            Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, _
                                            Row.Index, Me.dgvItem, Me.Def_Prec, Me)
                                        End If
                                    End If
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication salary number, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Item = e.RowIndex
                                            Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value = ij.mPrimary
                                        DTASalary_DGV.FillBy_PrimarySalary(TBLSalary_DGV, CDec(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value))
                                        If TBLSalary_DGV.Rows.Count > 0 Then
                                            'Batasan Payroll -- baru
                                            If Me.Obj.ApakahBolehInputSalary(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, e.RowIndex).Value, Me) = False Then
                                                If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                    Me.EraseRowIndex_Item = e.RowIndex
                                                    Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                                Else
                                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                End If
                                                Me.ObjCD.SetNumber(Me.dgvItem)
                                                Exit Sub
                                            End If
                                            RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                            Me.dgvItem.Item(Me.ColEmp_SalaryDoc.Index, e.RowIndex).Value = RowSalary_DGV.No_Transaksi
                                            Me.dgvItem.Item(Me.ColEmp_PrimaryEmployee.Index, e.RowIndex).Value = RowSalary_DGV.Primary_Employee
                                            IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.dgvItem, Me.ColIdentifier.Index, Me.Baru, Me.TxtPrimary.Text, Row.Index)
                                            Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value = IdentifierStr
                                            If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value) = True Then
                                                Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = ""
                                            Else
                                                Me.dgvItem.Item(Me.ColEmp_Remarks.Index, e.RowIndex).Value = RowSalary_DGV.Remarks
                                            End If
                                            Me.dgvItem.Item(Me.ColEmp_SalaryDate.Index, e.RowIndex).Value = Format(RowSalary_DGV.Tanggal_Transaksi, "dd MMM yyyy")
                                            Me.dgvItem.Item(Me.ColEmp_TakeHomePay.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Salary_Received, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_OtherDeduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Other_Deduction, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_LoanInstallment.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Loan_Installment, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_IncomeTax.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income_Tax, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_GrossSalary.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Gross_Salary, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_Deduction.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Deduction, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_Income.Index, e.RowIndex).Value = ObjCD.DigitGrouping(RowSalary_DGV.Income, Me.Def_Prec)
                                            Me.dgvItem.Item(Me.ColEmp_EmployeeCode.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Code
                                            Me.dgvItem.Item(Me.ColEmp_Name.Index, e.RowIndex).Value = RowSalary_DGV.Employee_Name
                                            CboDGV = Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex)
                                            Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome.Index, e.RowIndex).Value = CboDGV.Items.Item(RowSalary_DGV.Tax)
                                            Me.dgvItem.Item(Me.ColEmp_TipeTaxIncome_Enum.Index, e.RowIndex).Value = RowSalary_DGV.Tax
                                            'Insert mTBLTrans_Payment_Employee_d
                                            RowItem = Me.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                                            RowItem.primary = -1
                                            RowItem.primary_transaksi = CDbl(Me.TxtPrimary.Text)
                                            RowItem.identifier = Row.Cells(Me.ColIdentifier.Index).Value
                                            RowItem.primary_akun_payment = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                                            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
                                            RowItem.akun_payment_code = Me.ObjAkun.GetAccountCode(RowItem.primary_akun_payment)
                                            RowItem.akun_payment_desc = Me.ObjAkun.GetDescription_ByPrimaryAcc(RowItem.primary_akun_payment)
                                            RowItem.amount = Me.ObjCD.DigitGrouping(CDec(Row.Cells(Me.ColEmp_TakeHomePay.Index).Value), Me.Def_Prec)
                                            RowItem.remarks = ""
                                            RowItem.Index_Row_Salary = Row.Index
                                            Me.mTBLItem_Employee_d.Rows.Add(RowItem)
                                            'Insert Tbl Salary
                                            If TBLSalary_DGV.Rows.Count > 0 Then
                                                RowSalary_DGV = TBLSalary_DGV.Rows(0)
                                                RowSalary = Me.TBLSalary.Newtbltrans_salaryRow
                                                For ikl As Integer = 0 To Me.TBLSalary.Columns.Count - 1
                                                    RowSalary.Item(ikl) = RowSalary_DGV.Item(ikl)
                                                Next
                                                Me.TBLSalary.Rows.Add(RowSalary)
                                            End If
                                            'Event Selection Changed dipanggil ketika CellEndEdit
                                            Me.mCalledCellEndEdit = True
                                            If Me.mCalledCellEndEdit = True Then
                                                'Display DGVPayment
                                                Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged( _
                                                Me.DGVPayment, Row.Index, Me.ObjCD, Me.Def_Prec, Me)
                                                'Display TextBox Hitungan
                                                Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, _
                                                Row.Index, Me.dgvItem, Me.Def_Prec, Me)
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Item = e.RowIndex
                                    Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvItem)
                    End If
                    If e.ColumnIndex = Me.ColEmp_Remarks.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColEmp_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColEmp_Remarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASalary_DGV.Dispose()
            TBLSalary_DGV.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItem.RowHeaderMouseClick
        Try
            If IsNothing(Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value) = False Then
                If Me.dgvItem.Rows.Count > 0 Then
                    If Me.dgvItem.SelectedCells.Count > 0 Then
                        Me.mIndexSelectionDGVSalary = Me.dgvItem.SelectedCells(0).RowIndex
                        If Me.mBufferIndexSelectionDGVSalary <> Me.mIndexSelectionDGVSalary Then
                            'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                            Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.SelectedCells(0).RowIndex.ToString
                            'Masuk in value pada masing-masing textbox
                            Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, Me.mIndexSelectionDGVSalary, Me.dgvItem, Me.Def_Prec, Me)
                            'Masuk in Payment Item Selection Changed
                            Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged(Me.DGVPayment, Me.mIndexSelectionDGVSalary, Me.ObjCD, Me.Def_Prec, Me)
                        End If
                    End If
                End If
            Else
                Me.TxtDeduction.Text = "0"
                Me.TxtIncome.Text = "0"
                Me.TxtGrossSalary.Text = "0"
                Me.TxtLoanInstallment.Text = "0"
                Me.TxtIncomeTax.Text = "0"
                Me.TxtOtherDedution.Text = "0"
                Me.TxtSalaryReceived.Text = "0"
                Me.DGVPayment.Rows.Clear()
                Me.DGVJournalItem.Rows.Clear()
                'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.Rows.Count - 1
            End If
            If Me.TxtRemarks.Enabled = False Then
                Me.DGVPayment.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        'Dim ind As Integer = 0
        'Dim PrimaryAkun As Double = -1
        'Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Try
            'PrimaryAkun = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_Salary_PrimaryCashAccount)
            'If Me.dgvItem.SelectedCells.Count = 0 Then
            '    Me.TxtDeduction.Text = "0"
            'Else
            '    ind = Me.dgvItem.SelectedCells(0).RowIndex
            'End If
            If Me.mCalledCellEndEdit = True Then
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    If Me.dgvItem.SelectedCells(0).RowIndex = Me.dgvItem.Rows.Count - 1 Then
                        'bersihkan objek
                        Me.TxtDeduction.Text = "0"
                        Me.TxtIncome.Text = "0"
                        Me.TxtGrossSalary.Text = "0"
                        Me.TxtLoanInstallment.Text = "0"
                        Me.TxtIncomeTax.Text = "0"
                        Me.TxtOtherDedution.Text = "0"
                        Me.TxtSalaryReceived.Text = "0"
                        Me.DGVPayment.Rows.Clear()
                        Me.DGVJournalItem.Rows.Clear()
                        'Masukan Index ke buffer index, supaya ga load 2 kali jika index-nya sama
                        Me.mBufferIndexSelectionDGVSalary = Me.dgvItem.Rows.Count - 1
                    End If
                End If
            End If
            'Return ke nilai default
            Me.mCalledCellEndEdit = False

            If Me.dgvItem.SelectedCells.Count > 0 Then
                'If Me.dgvItem.SelectedCells(0).RowIndex <> Me.dgvItem.Rows.Count - 1 Then
                Me.ShowDetailInfo(Me.dgvItem.SelectedCells(0).RowIndex)
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.dgvItem.AllowUserToAddRows = False
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.dgvItem.Rows.Count > 0 Then
                If IsNothing(Me.dgvItem.SelectedCells(0).RowIndex) = False Then
                    'Load Ulang Data yang disorot terakhir
                    Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_NilaiAccountSelectionChanged(Me.DGVPayment, _
                    Me.dgvItem.SelectedCells(0).RowIndex, Me.ObjCD, Me.Def_Prec, Me)
                    'Nilai TextBox-nya, diperbaharui
                    Me.Obj.FungsiBtnHitungPayment_IsiTxtSalary(Me.TBLSalary, Me.dgvItem.SelectedCells(0).RowIndex, Me.dgvItem, Me.Def_Prec, Me)
                End If
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim RowDelete As DataGridViewRow = Me.dgvItem.Rows(e.Row.Index)
        Try
            'Hapus juga baris yang di mTBLTrans_Payment_Employee_d by Identifier, IndexRowSalary, dan Primary Trans Salary
            Me.Obj.FungsiBtnHitungPayment_IsiDGVItem_UserDeletingRow(Me.TBLSalary, Me.mTBLItem_Employee_d, e.Row, Me.dgvItem, Me.DGVPayment, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPayment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVPayment.CellBeginEdit
        Try
            If Me.TxtSalaryReceived.Text = "0" Then
                e.Cancel = True
            End If
            If e.ColumnIndex = Me.ColDetailPayment_AccountCode.Index _
            Or e.ColumnIndex = Me.ColDetailPayment_Amount.Index _
            Or e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                Me.Buffer = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput_PayAccount(ByVal vPrimaryAkun As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DGVPayment.Rows
                If rowDGV.Cells(Me.ColDetailPayment_PrimaryAkun.Index).Value = vPrimaryAkun Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DGVPayment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVPayment.CellEndEdit
        Dim Row As DataGridViewRow = Me.DGVPayment.Rows(e.RowIndex)
        Dim Index As Integer
        Dim BufferPrimaryAkunPay As Double = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColDetailPayment_AccountCode.Index Then
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput_PayAccount(Me.ObjAkun.GetPrimary(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                'Batasan Akun
                                If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = False Then
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_PayAccount = e.RowIndex
                                        Me.BackgroundWorker_Payment.RunWorkerAsync()
                                    Else
                                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                                BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                    Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                End If
                                Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColIdentifier.Index, CInt(Me.mIndexSelectionDGVSalary)).Value
                                Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Me.dgvItem.SelectedCells(0).RowIndex
                                'Koding Settingan Ke mTBLTransPayment_Item_Employee_d
                                Me.PengaturanIsi_mTBLItem_Employee_d(Row, BufferPrimaryAkunPay)
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_PayAccount = e.RowIndex
                                        Me.BackgroundWorker_Payment.RunWorkerAsync()
                                    Else
                                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DGVPayment)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New AccModule.FrmMasterAkun_Cari
                            ij.txtCari.Text = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput_PayAccount(ij.mPrimaryAkun)
                                If Index < 0 Then
                                    'Batasan Akun
                                    If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(ij.mPrimaryAkun) = False Then
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_PayAccount = e.RowIndex
                                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                                        Else
                                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                    If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                        Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                    End If
                                    Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColIdentifier.Index, CInt(Me.mIndexSelectionDGVSalary)).Value
                                    Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Me.dgvItem.SelectedCells(0).RowIndex
                                    BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                    Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                    'Koding Settingan Ke mTBLTransPayment_Item_Employee_d
                                    Me.PengaturanIsi_mTBLItem_Employee_d(Row, BufferPrimaryAkunPay)
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_PayAccount = e.RowIndex
                                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                                        Else
                                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DGVPayment)
                                        Exit Sub
                                    Else
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                        'Batasan Akun
                                        If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(ij.mPrimaryAkun) = False Then
                                            Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex_PayAccount = e.RowIndex
                                                Me.BackgroundWorker_Payment.RunWorkerAsync()
                                            Else
                                                Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.ObjCD.SetNumber(Me.dgvItem)
                                            Exit Sub
                                        End If
                                        If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                            Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                        End If
                                        Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColIdentifier.Index, CInt(Me.mIndexSelectionDGVSalary)).Value
                                        Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Me.dgvItem.SelectedCells(0).RowIndex
                                        BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                        Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                        'Koding Settingan Ke mTBLTransPayment_Item_Employee_d
                                        Me.PengaturanIsi_mTBLItem_Employee_d(Row, BufferPrimaryAkunPay)
                                    End If
                                End If
                            ElseIf ij.mPrimaryAkun = -1 Then
                                If Me.DGVPayment.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_PayAccount = e.RowIndex
                                    Me.BackgroundWorker_Payment.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DGVPayment)
                        '----------------------------------------------------
                    End If
                    If e.ColumnIndex = Me.ColDetailPayment_Amount.Index Then
                        If Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = "-1" Or _
                        String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.Def_Prec)
                            Exit Sub
                        End If
                        'Koding Settingan Ke mTBLTransPayment_Item_Employee_d
                        BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                        Me.PengaturanIsi_mTBLItem_Employee_d(Row, BufferPrimaryAkunPay)
                        'Digitgrouping
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value, Me.Def_Prec)
                    End If
                    If e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                        If Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = "-1" Or _
                        String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColDetailPayment_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        'Koding Settingan Ke mTBLTransPayment_Item_Employee_d
                        BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                        Me.PengaturanIsi_mTBLItem_Employee_d(Row, BufferPrimaryAkunPay)
                    End If
                Else
                    If e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PengaturanIsi_mTBLItem_Employee_d(ByVal RowPayment As DataGridViewRow, _
    ByVal vPrimaryAkunDummy As Double)
        Dim Filter As String = ""
        Dim Row As DseTransPayment.tbltrans_payment_item_employee_dRow
        Try
            Filter = "[identifier] = '" + RowPayment.Cells(Me.ColDetailPayment_Identifier.Index).Value.ToString + _
            "' AND [Index Row Salary] = " + RowPayment.Cells(Me.ColDetailPayment_IndexRowSalary.Index).Value.ToString + _
            " AND [primary akun payment] = " + vPrimaryAkunDummy.ToString + ""
            Me.mTBLItem_Employee_d.DefaultView.RowFilter = Filter
            If Me.mTBLItem_Employee_d.DefaultView.Count > 0 Then
                'diganti
                Row = Me.mTBLItem_Employee_d.DefaultView(0).Row
                Row.identifier = RowPayment.Cells(Me.ColDetailPayment_Identifier.Index).Value.ToString
                Row.amount = CDec(RowPayment.Cells(Me.ColDetailPayment_Amount.Index).Value)
                Row.primary_akun_payment = CDbl(RowPayment.Cells(Me.ColDetailPayment_PrimaryAkun.Index).Value)
                If IsNothing(RowPayment.Cells(Me.ColDetailPayment_Remarks.Index).Value) = True Then
                    Row.remarks = ""
                Else
                    Row.remarks = RowPayment.Cells(Me.ColDetailPayment_Remarks.Index).Value.ToString
                End If
                Row.akun_payment_code = RowPayment.Cells(Me.ColDetailPayment_AccountCode.Index).Value.ToString
                Row.akun_payment_desc = RowPayment.Cells(Me.ColDetailPayment_AccountDesc.Index).Value.ToString
                Row.Index_Row_Salary = CDbl(RowPayment.Cells(Me.ColDetailPayment_IndexRowSalary.Index).Value)
            Else
                'ditambahkan row baru
                Row = Me.mTBLItem_Employee_d.Newtbltrans_payment_item_employee_dRow
                Row.primary = -1
                Row.primary_transaksi = CDbl(Me.TxtPrimary.Text)
                Row.identifier = RowPayment.Cells(Me.ColDetailPayment_Identifier.Index).Value.ToString
                Row.amount = CDec(RowPayment.Cells(Me.ColDetailPayment_Amount.Index).Value)
                Row.primary_akun_payment = CDbl(RowPayment.Cells(Me.ColDetailPayment_PrimaryAkun.Index).Value)
                If IsNothing(RowPayment.Cells(Me.ColDetailPayment_Remarks.Index).Value) = True Then
                    Row.remarks = ""
                Else
                    Row.remarks = RowPayment.Cells(Me.ColDetailPayment_Remarks.Index).Value.ToString
                End If
                Row.akun_payment_code = RowPayment.Cells(Me.ColDetailPayment_AccountCode.Index).Value.ToString
                Row.akun_payment_desc = RowPayment.Cells(Me.ColDetailPayment_AccountDesc.Index).Value.ToString
                Row.Index_Row_Salary = CDbl(RowPayment.Cells(Me.ColDetailPayment_IndexRowSalary.Index).Value)
                Me.mTBLItem_Employee_d.Rows.Add(Row)
            End If
            Me.mTBLItem_Employee_d.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPayment_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVPayment.UserDeletingRow
        Dim RowDelete As DataGridViewRow = Me.DGVPayment.Rows(e.Row.Index)
        Try
            'Hapus juga baris yang di mTBLTrans_Payment_Employee_d by identifier, Primary Akun
            Me.Obj.FungsiBtnHitungPayment_IsiDGVPayment_UserDeletingRow(Me.mTBLItem_Employee_d, e.Row, Me.DGVPayment, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_Payment_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_Payment.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_Payment_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_Payment.RunWorkerCompleted
        Try
            'Hapus yang di datagridview Payment
            Me.DGVPayment.Rows.RemoveAt(Me.EraseRowIndex_PayAccount)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DetailEmployee_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_DetailEmployee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DetailEmployee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_DetailEmployee.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex_Item)
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahRowItemSalaryKosong() As Boolean
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count <= 0 Then
                Me.dgvItem.AllowUserToAddRows = True
                Return True
            Else
                Me.dgvItem.AllowUserToAddRows = True
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Lbl_Income_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Income.LinkClicked
        Dim Frm_a As New FrmTransaksiPayment_Income_Deduction
        Dim Row_a As DseTransPayment.tbltrans_salary_itemRow
        Dim RowDGV As DataGridViewRow
        Dim CboDGV As DataGridViewComboBoxCell
        Dim Filter As String = ""
        Dim RV As Data.DataRowView
        Try
            Frm_a.dgvItem.Rows.Clear()
            If Me.ApakahRowItemSalaryKosong() = True Then
                Exit Sub
            End If
            Me.Obj.FungsiBtnHitungPayment_LinkLabel(ClsTrans_Payment.EnumLinkLabel.Income, _
            Convert.ToDouble(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value), _
            Me.TBLSalary, Me.TBLItemSalary_Income, Me.TBLItemSalary_Deduction, Me.TBLItemSalary_Other, Me.TBLLoanPayment, _
            Me.TBLLoanPaymentItem)
            If Me.TBLItemSalary_Income.Rows.Count > 0 Then
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                'Frm_a.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Frm_a.BackColorAll
                If IsNothing(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value) = False Then
                    Filter = "[Primary Transaksi] = " + Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value.ToString + ""
                    Me.TBLItemSalary_Income.DefaultView.RowFilter = Filter
                    If Me.TBLItemSalary_Income.DefaultView.Count > 0 Then
                        For Each RV In Me.TBLItemSalary_Income.DefaultView
                            Row_a = RV.Row
                            Frm_a.dgvItem.Rows.Add()
                            RowDGV = Frm_a.dgvItem.Rows(Frm_a.dgvItem.Rows.Count - 1)
                            With RowDGV
                                .Cells(Frm_a.ColNo.Index).Value = ""
                                .Cells(Frm_a.ColSlipDesc.Index).Value = Row_a.SlipDescDescription
                                .Cells(Frm_a.ColComponentName.Index).Value = Row_a.Description
                                If Row_a.Type = ClsTransaksiSalary_Item.EnumType.Income Then
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Income"
                                Else
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Deduction"
                                End If
                                CboDGV = .Cells(Frm_a.ColTaxElement.Index)
                                .Cells(Frm_a.ColTaxElement.Index).Value = CboDGV.Items.Item(Row_a.Tax_Component)
                                .Cells(Frm_a.ColExpenseAccount.Index).Value = Row_a.Account_Code
                                .Cells(Frm_a.ColAmount.Index).Value = ObjCD.DigitGrouping(Row_a.Amount, Me.Def_Prec)
                                If Row_a.Override = 1 Then
                                    .Cells(Frm_a.ColOverride.Index).Value = True
                                Else
                                    .Cells(Frm_a.ColOverride.Index).Value = False
                                End If
                                .Cells(Frm_a.ColOverrideReason.Index).Value = Row_a.Override_Reason
                                .Cells(Frm_a.ColOriginalValue.Index).Value = ObjCD.DigitGrouping(Row_a.Original_Value, Me.Def_Prec)
                                .Cells(Frm_a.ColRemarks.Index).Value = Row_a.Remarks
                                .Cells(Frm_a.ColPrimarySlipDesc.Index).Value = Row_a.Primary_Slip_Desc
                                .Cells(Frm_a.ColIdentifierKomponen.Index).Value = Row_a.Identifier_Salary_Component
                                .Cells(Frm_a.ColPrimaryAkunExpense.Index).Value = Row_a.Primary_Expense_Account
                                .Cells(Frm_a.ColExplaination.Index).Value = Row_a.Explaination
                            End With
                        Next
                    End If
                End If
                Me.TBLItemSalary_Income.DefaultView.RowFilter = ""
                Me.ObjCD.SetNumber(Frm_a.dgvItem)
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Frm_a.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm_a.Dispose()
        End Try
    End Sub

    Private Sub Lbl_Deduction_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Deduction.LinkClicked
        Dim Frm_a As New FrmTransaksiPayment_Income_Deduction
        Dim Row_a As DseTransPayment.tbltrans_salary_itemRow
        Dim RowDGV As DataGridViewRow
        Dim CboDGV As DataGridViewComboBoxCell
        Dim Filter As String = ""
        Dim RV As Data.DataRowView
        Try
            Frm_a.dgvItem.Rows.Clear()
            If Me.ApakahRowItemSalaryKosong() = True Then
                Exit Sub
            End If
            Me.Obj.FungsiBtnHitungPayment_LinkLabel(ClsTrans_Payment.EnumLinkLabel.Deduction, _
            Convert.ToDouble(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value), _
            Me.TBLSalary, Me.TBLItemSalary_Income, Me.TBLItemSalary_Deduction, Me.TBLItemSalary_Other, Me.TBLLoanPayment, _
            Me.TBLLoanPaymentItem)
            If Me.TBLItemSalary_Deduction.Rows.Count > 0 Then
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Frm_a.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Frm_a.BackColorAll
                If IsNothing(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value) = False Then
                    Filter = "[Primary Transaksi] = " + Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value.ToString + ""
                    Me.TBLItemSalary_Deduction.DefaultView.RowFilter = Filter
                    If Me.TBLItemSalary_Deduction.DefaultView.Count > 0 Then
                        For Each RV In Me.TBLItemSalary_Deduction.DefaultView
                            Row_a = RV.Row
                            Frm_a.dgvItem.Rows.Add()
                            RowDGV = Frm_a.dgvItem.Rows(Frm_a.dgvItem.Rows.Count - 1)
                            With RowDGV
                                .Cells(Frm_a.ColNo.Index).Value = ""
                                .Cells(Frm_a.ColSlipDesc.Index).Value = Row_a.SlipDescDescription
                                .Cells(Frm_a.ColComponentName.Index).Value = Row_a.Description
                                If Row_a.Type = ClsTransaksiSalary_Item.EnumType.Income Then
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Income"
                                Else
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Deduction"
                                End If
                                CboDGV = .Cells(Frm_a.ColTaxElement.Index)
                                .Cells(Frm_a.ColTaxElement.Index).Value = CboDGV.Items.Item(Row_a.Tax_Component)
                                .Cells(Frm_a.ColExpenseAccount.Index).Value = Row_a.Account_Code
                                .Cells(Frm_a.ColAmount.Index).Value = ObjCD.DigitGrouping(Row_a.Amount, Me.Def_Prec)
                                If Row_a.Override = 1 Then
                                    .Cells(Frm_a.ColOverride.Index).Value = True
                                Else
                                    .Cells(Frm_a.ColOverride.Index).Value = False
                                End If
                                .Cells(Frm_a.ColOverrideReason.Index).Value = Row_a.Override_Reason
                                .Cells(Frm_a.ColOriginalValue.Index).Value = ObjCD.DigitGrouping(Row_a.Original_Value, Me.Def_Prec)
                                .Cells(Frm_a.ColRemarks.Index).Value = Row_a.Remarks
                                .Cells(Frm_a.ColPrimarySlipDesc.Index).Value = Row_a.Primary_Slip_Desc
                                .Cells(Frm_a.ColIdentifierKomponen.Index).Value = Row_a.Identifier_Salary_Component
                                .Cells(Frm_a.ColPrimaryAkunExpense.Index).Value = Row_a.Primary_Expense_Account
                                .Cells(Frm_a.ColExplaination.Index).Value = Row_a.Explaination
                            End With
                        Next
                    End If
                End If
                Me.TBLItemSalary_Deduction.DefaultView.RowFilter = ""
                Me.ObjCD.SetNumber(Frm_a.dgvItem)
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Frm_a.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm_a.Dispose()
        End Try
    End Sub

    Private Sub Lbl_Other_Deduction_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Other_Deduction.LinkClicked
        Dim Frm_a As New FrmTransaksiPayment_Income_Deduction
        Dim Row_a As DseTransPayment.tbltrans_salary_item_otherdeductionRow
        Dim RowDGV As DataGridViewRow
        Dim CboDGV As DataGridViewComboBoxCell
        Dim Filter As String = ""
        Dim RV As Data.DataRowView
        Try
            Frm_a.dgvItem.Rows.Clear()
            If Me.ApakahRowItemSalaryKosong() = True Then
                Exit Sub
            End If
            Me.Obj.FungsiBtnHitungPayment_LinkLabel(ClsTrans_Payment.EnumLinkLabel.OtherDeduction, _
            Convert.ToDouble(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value), _
            Me.TBLSalary, Me.TBLItemSalary_Income, Me.TBLItemSalary_Deduction, Me.TBLItemSalary_Other, Me.TBLLoanPayment, _
            Me.TBLLoanPaymentItem)
            If Me.TBLItemSalary_Other.Rows.Count > 0 Then
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Frm_a.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Frm_a.BackColorAll
                If IsNothing(Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value) = False Then
                    Filter = "[Primary Transaksi] = " + Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value.ToString + ""
                    Me.TBLItemSalary_Other.DefaultView.RowFilter = Filter
                    If Me.TBLItemSalary_Other.DefaultView.Count > 0 Then
                        For Each RV In Me.TBLItemSalary_Other.DefaultView
                            Row_a = RV.Row
                            Frm_a.dgvItem.Rows.Add()
                            RowDGV = Frm_a.dgvItem.Rows(Frm_a.dgvItem.Rows.Count - 1)
                            With RowDGV
                                .Cells(Frm_a.ColNo.Index).Value = ""
                                .Cells(Frm_a.ColSlipDesc.Index).Value = Row_a.SlipDescDescription
                                .Cells(Frm_a.ColComponentName.Index).Value = Row_a.Description
                                If Row_a.Type = ClsTransaksiSalary_Item.EnumType.Income Then
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Income"
                                Else
                                    .Cells(Frm_a.ColTipeKomponen.Index).Value = "Deduction"
                                End If
                                CboDGV = .Cells(Frm_a.ColTaxElement.Index)
                                .Cells(Frm_a.ColTaxElement.Index).Value = CboDGV.Items.Item(Row_a.Tax_Component)
                                .Cells(Frm_a.ColExpenseAccount.Index).Value = Row_a.Account_Code
                                .Cells(Frm_a.ColAmount.Index).Value = ObjCD.DigitGrouping(Row_a.Amount, Me.Def_Prec)
                                If Row_a.Override = 1 Then
                                    .Cells(Frm_a.ColOverride.Index).Value = True
                                Else
                                    .Cells(Frm_a.ColOverride.Index).Value = False
                                End If
                                .Cells(Frm_a.ColOverrideReason.Index).Value = Row_a.Override_Reason
                                .Cells(Frm_a.ColOriginalValue.Index).Value = ObjCD.DigitGrouping(Row_a.Original_Value, Me.Def_Prec)
                                .Cells(Frm_a.ColRemarks.Index).Value = Row_a.Remarks
                                .Cells(Frm_a.ColPrimarySlipDesc.Index).Value = Row_a.Primary_Slip_Desc
                                .Cells(Frm_a.ColIdentifierKomponen.Index).Value = ""
                                .Cells(Frm_a.ColPrimaryAkunExpense.Index).Value = Row_a.Primary_Expense_Account
                                .Cells(Frm_a.ColExplaination.Index).Value = Row_a.Explaination
                            End With
                        Next
                    End If
                End If
                Me.TBLItemSalary_Other.DefaultView.RowFilter = ""
                Me.ObjCD.SetNumber(Frm_a.dgvItem)
                Frm_a.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Frm_a.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Frm_a.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm_a.Dispose()
        End Try
    End Sub

    Private Sub LblIncomeTax_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblIncomeTax.Click
        
    End Sub

    Private Sub Lbl_Loan_Installment_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Loan_Installment.LinkClicked
        Dim Frm_a As New FrmTransaksiPayment_Loan_Payment
        Try
            If Me.ApakahRowItemSalaryKosong() = True Then
                Exit Sub
            End If
            If CDec(Me.TxtLoanInstallment.Text) <> 0 Then
                Frm_a.mPrimarySalaryFromPayment = Me.dgvItem.Item(Me.ColEmp_PrimarySalary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value
                Frm_a.mDef_Prec_Payment = Me.Def_Prec
                Frm_a.BackColorAll = Me.BackColorAll
                Frm_a.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm_a.Dispose()
        End Try
    End Sub

    Public Sub GetNilaiPrimarySalary(ByRef vPrimaryMin As Double, ByRef vPrimaryMax As Double)
        Dim RowDGV As DataGridViewRow
        Try
            Me.dgvItem.AllowUserToAddRows = False
            If Me.dgvItem.Rows.Count > 0 Then
                For Each RowDGV In Me.dgvItem.Rows
                    If RowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value > vPrimaryMax Then
                        vPrimaryMax = Convert.ToDouble(RowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value)
                    End If
                    If RowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value < vPrimaryMin Then
                        vPrimaryMin = Convert.ToDouble(RowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value)
                    End If
                Next
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 1 Then
                    Me.dgvItem.AllowUserToAddRows = False
                    If Me.dgvItem.Rows.Count > 0 Then
                        Me.GetNilaiPrimarySalary(PrimaryMin, PrimaryMax)
                        Me.Obj.BuatJournalViewPayment(PrimaryMin, PrimaryMax, Me.mTBLItem_Employee_d, Me.TBLHasil, Me.Def_Prec, Me)
                    End If
                    Me.dgvItem.AllowUserToAddRows = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAttendance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAttendance.LinkClicked
        Try
            If Me.LblAttendance.Text = "Show Detail" Then
                Me.SplitContainer4.Panel2Collapsed = False
                Me.LblAttendance.Text = "Hide Detail"
            Else
                Me.SplitContainer4.Panel2Collapsed = True
                Me.LblAttendance.Text = "Show Detail"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblIncomeTax_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblIncomeTax.LinkClicked
        Dim k As New FrmTrans_Gaji_IncomeTax2
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim row As DseTransaksiSalary.tbltrans_salary2Row
        Dim DTAItemTax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_itemtaxTableAdapter

        Try
            If Me.ApakahRowItemSalaryKosong() = True Then
                Exit Sub
            End If
            If Me.dgvItem.SelectedCells.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTAItemTax.Connection.ConnectionString = Core_Module.GetConstring

                DTA.FillBy_Primary(TBL, rowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value)
                If TBL.Rows.Count > 0 Then
                    row = TBL.Rows(0)

                    DTAItemTax.FillBy_PrimaryTransaksi(Me.TBLItemTax, rowDGV.Cells(Me.ColEmp_PrimarySalary.Index).Value)
                    k.vTBLItemTax = Me.TBLItemTax
                    k.vEmployeeCode = rowDGV.Cells(Me.ColEmp_EmployeeCode.Index).Value
                    k.vTransactionDate = rowDGV.Cells(Me.ColEmp_SalaryDate.Index).Value
                    k.vEmployeeName = rowDGV.Cells(Me.ColEmp_Name.Index).Value
                    k.vPrimaryEmployee = rowDGV.Cells(Me.ColEmp_PrimaryEmployee.Index).Value
                    k.vNIK_Kary = row.A_NIK
                    k.vNPWP = row.A_NPWP
                    k.vStatusNPWP = row.A_Status_NPWP
                    k.vKodeNegaraDomisili = row.A_Kode_Negara_Domisili
                    k.vJenisKelamin = row.A_Jenis_Kelamin
                    k.vAlamatPajak1 = row.A_Alamat1
                    k.vAlamatPajak2 = row.A_Alamat2
                    k.vDepartement = Me.ObjKaryawan.GetDeptDescription(row.Primary_Employee)
                    k.vPosition = Me.ObjPosition.GetDesc(row.Primary_Position)

                    k.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            k = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = True Then
                    Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemarks.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.ReadOnly = True Then
                    Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
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

    Private Sub txtJournalDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDesc.GotFocus
        Me.TxtJournalDesc.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtJournalDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDesc.LostFocus
        Me.TxtJournalDesc.BackColor = Color.White
    End Sub

    Private Sub TxtDocumentNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDocumentNo.TextChanged
        Try
            Me.SetTxtJournal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtJournal()
        Dim ArrJournalDesc() As String
        Dim Date_JournalDesc As String = ""
        Dim JournalDesc_ As String = ""

        Try
            If Me.txtJournalNo.Text = Me.Buffer Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            ElseIf Me.txtJournalNo.Text = "" Then
                Me.txtJournalNo.Text = Me.TxtDocumentNo.Text
            End If

            'set journal description
            If Me.TxtJournalDesc.Text = "" Then
                Me.TxtJournalDesc.Text = "Salary Payment - " & Format(Me.DtpDate.Value, "MMM yyyy")
            End If

            If Me.TxtJournalDesc.Text.Contains(Format(Me.BufferTransPaymentDate, "MMM yyyy")) = True Then
                Me.TxtJournalDesc.Text = Me.TxtJournalDesc.Text.Replace(Format(Me.BufferTransPaymentDate, "MMM yyyy"), Format(Me.DtpDate.Value, "MMM yyyy"))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.ValueChanged
        Try
            'Me.Buffer = Me.TxtDocumentNo.Text
            'Me.SetTxtJournal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class