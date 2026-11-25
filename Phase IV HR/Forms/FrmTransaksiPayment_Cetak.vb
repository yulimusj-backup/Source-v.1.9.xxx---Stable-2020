Option Explicit On
Option Strict Off

Public Class FrmTransaksiPayment_Cetak
    Public mFrmTransPayment As FrmTransaksiPayment
    Public mPrimarySalaryMin As Double = -1
    Public mPrimarySalaryMax As Double = -1
    Public mNoSalary As String = ""
    Public mDefPrec As Double = 0
    Public mTanggalTransaksi As Date = #1/1/1950#
    Public mPeriodFrom As Date = #1/1/1950#
    Public mPeriodTo As Date = #1/1/1950#

    Private DSCetakPayment As New DseTransPayment_Cetak
    Private TblCetakPaymentDummy As New DseTransPayment_Cetak.TblCetakDataTable
    Private TblCetakPayment As DseTransPayment_Cetak.TblCetakDataTable = DSCetakPayment.TblCetak
    Private RowCetakPayment As DseTransPayment_Cetak.TblCetakRow
    Private ReportCetakPayment As New CrpTransaksiPayment_Cetak_PaymentDetail
    Private ParamSubTotalAccDesc As String = ""
    Private ParamSubTotalDebet As String = ""
    Private ParamSubTotalKredit As String = ""
    Dim objCd As New Core_Module.ClsCekData
    Dim ObjPayment As New HR.ClsTrans_Payment

    Private Sub FrmTransaksiPayment_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Me.isiDataTest()
            Me.refreshReport()
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub refreshReport()
        Dim rowDgv As Windows.Forms.DataGridViewRow
        Try
            Me.ParamSubTotalAccDesc = ""
            Me.ParamSubTotalDebet = ""
            Me.ParamSubTotalKredit = ""
            Me.TblCetakPayment.Clear()
            If Me.TblCetakPaymentDummy.Rows.Count > 0 Then
                ''get journal from dgv
                For Each rowDgv In Me.mFrmTransPayment.DGVJournalItem.Rows
                    With rowDgv
                        If Me.mFrmTransPayment.DGVJournalItem.Rows.IndexOf(rowDgv) <> 0 Then
                            Me.ParamSubTotalAccDesc += vbNewLine & .Cells(Me.mFrmTransPayment.ColJurnal_Account.Index).Value.ToString _
                            & " - " & .Cells(Me.mFrmTransPayment.ColJurnal_DescAccount.Index).Value.ToString
                            If Convert.ToDecimal(.Cells(Me.mFrmTransPayment.ColJurnal_Debet.Index).Value) <> "0" Then
                                Me.ParamSubTotalDebet += vbNewLine & .Cells(Me.mFrmTransPayment.ColJurnal_Debet.Index).Value.ToString
                            Else
                                Me.ParamSubTotalDebet += vbNewLine & "-"
                            End If
                            If Convert.ToDecimal(.Cells(Me.mFrmTransPayment.ColJurnal_Kredit.Index).Value) <> "0" Then
                                Me.ParamSubTotalKredit += vbNewLine & .Cells(Me.mFrmTransPayment.ColJurnal_Kredit.Index).Value.ToString
                            Else
                                Me.ParamSubTotalKredit += vbNewLine & "-"
                            End If
                        Else
                            Me.ParamSubTotalAccDesc = .Cells(Me.mFrmTransPayment.ColJurnal_Account.Index).Value.ToString _
                            & " - " & .Cells(Me.mFrmTransPayment.ColJurnal_DescAccount.Index).Value.ToString
                            If Convert.ToDecimal(.Cells(Me.mFrmTransPayment.ColJurnal_Debet.Index).Value) <> "0" Then
                                Me.ParamSubTotalDebet = .Cells(Me.mFrmTransPayment.ColJurnal_Debet.Index).Value.ToString
                            Else
                                Me.ParamSubTotalDebet = "-"
                            End If
                            If Convert.ToDecimal(.Cells(Me.mFrmTransPayment.ColJurnal_Kredit.Index).Value) <> "0" Then
                                Me.ParamSubTotalKredit = .Cells(Me.mFrmTransPayment.ColJurnal_Kredit.Index).Value.ToString
                            Else
                                Me.ParamSubTotalKredit = "-"
                            End If
                        End If
                    End With
                Next
                ''sort and copy
                Me.TblCetakPaymentDummy.DefaultView.RowFilter = ""
                Me.TblCetakPaymentDummy.DefaultView.Sort = "[Primary Salary], [Position Debet Kredit], [Description], [Account Description]"
                For i As Integer = 0 To Me.TblCetakPaymentDummy.DefaultView.Count - 1
                    Me.TblCetakPayment.ImportRow(Me.TblCetakPaymentDummy.DefaultView(i).Row)
                Next
                ''set params
                Me.ReportCetakPayment.SetDataSource(Me.DSCetakPayment)
                Me.ReportCetakPayment.SetParameterValue("DocumentNo", Me.mNoSalary)
                Me.ReportCetakPayment.SetParameterValue("TransactionDate", Me.mTanggalTransaksi)
                Me.ReportCetakPayment.SetParameterValue("PeriodFrom", Me.objCd.GetFirstDateofTheMonth(Me.mPeriodFrom))
                Me.ReportCetakPayment.SetParameterValue("PeriodTo", Me.objCd.GetLasDateOfTheMonth(Me.mPeriodTo))
                Me.ReportCetakPayment.SetParameterValue("SubTotalAccDesc", Me.ParamSubTotalAccDesc)
                Me.ReportCetakPayment.SetParameterValue("SubTotalDebet", Me.ParamSubTotalDebet)
                Me.ReportCetakPayment.SetParameterValue("SubTotalKredit", Me.ParamSubTotalKredit)
                Me.ReportCetakPayment.SetParameterValue(Me.ReportCetakPayment.Parameter_Format_Decimal.ParameterFieldName, Me.mDefPrec)
                Me.crv1.ReportSource = Me.ReportCetakPayment
                Me.crv1.ShowRefreshButton = False
                Me.crv1.Zoom(90)
                Me.crv1.Show()
                Me.crv1.Refresh()
            Else
                MsgBox("There isn't data.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub isiDataTest()
        Dim TBLHasil As New DseTransPayment.DataTable1DataTable
        Try
            Me.TblCetakPaymentDummy.Clear()
            Me.ObjPayment.BuatFormatDataSetCetakanPayment(Me.mPrimarySalaryMin, Me.mPrimarySalaryMax, _
            Me.mFrmTransPayment.mTBLItem_Employee_d, TBLHasil, Me.TblCetakPaymentDummy, Me.mFrmTransPayment.Def_Prec, Me.mFrmTransPayment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLHasil.Dispose()
        End Try
    End Sub

    ''rosadin test
    'Private Sub isiDataTest()
    '    Try
    '        Me.TblCetakPaymentDummy.Clear()

    '        ''jon
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000004/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Jon"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 4
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000004/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Jon"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 4
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000004/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Jon"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Loss Acc"
    '        Me.RowCetakPayment.Debet = 10000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 4
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000004/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Jon"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1810000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 4
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end jon

    '        ''iwan
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000001/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Iwan"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 1
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000001/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Iwan"
    '        Me.RowCetakPayment.Description = "PAYMENT"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1500000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 1
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000001/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Iwan"
    '        Me.RowCetakPayment.Description = "PAYMENT"
    '        Me.RowCetakPayment.Account_Description = "KK"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 500000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 1
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end iwan

    '        ''yudha
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000002/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Yudha"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2500000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 2
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000002/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Yudha"
    '        Me.RowCetakPayment.Description = "OVERTIME"
    '        Me.RowCetakPayment.Account_Description = "OT Expense"
    '        Me.RowCetakPayment.Debet = 750000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 2
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000002/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Yudha"
    '        Me.RowCetakPayment.Description = "PPh Pasal 21"
    '        Me.RowCetakPayment.Account_Description = "Pajak Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 75000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 2
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000002/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Yudha"
    '        Me.RowCetakPayment.Description = "KOPERASI"
    '        Me.RowCetakPayment.Account_Description = "Kop Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 2
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000002/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Yudha"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 3150000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 2
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end yudha

    '        ''bela
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000003/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Bela"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 3
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000003/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Bela"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 3
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000003/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Bela"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Gain Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 3
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000003/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Bela"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1775000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 3
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end bella

    '        ''abad
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000005/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Abad"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 5
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000005/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Abad"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 5
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000005/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Abad"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Gain Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 5
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000005/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Abad"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1775000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 5
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end abad

    '        ''sukoco
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000006/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Sukoco"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        Me.RowCetakPayment.Primary_Salary = 6
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow

    '        Me.RowCetakPayment.No_Document = "000006/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Sukoco"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 6
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000006/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Sukoco"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Gain Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 6
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000006/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Sukoco"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1775000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 6
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ''end sukoco

    '        ' ''andrew
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000007/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Andrew"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 7
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000007/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Andrew"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 7
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000007/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Andrew"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Gain Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 7
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000007/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Andrew"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1775000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 7
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ' ''end andrew  

    '        ' ''dudung
    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000008/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Dudung"
    '        Me.RowCetakPayment.Description = "BASIC SALARY"
    '        Me.RowCetakPayment.Account_Description = "Salary Expense"
    '        Me.RowCetakPayment.Debet = 2000000
    '        Me.RowCetakPayment.Kredit = 0
    '        Me.RowCetakPayment.Position_Debet_Kredit = 0
    '        Me.RowCetakPayment.Primary_Salary = 8
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000008/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Dudung"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 200000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 8
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000008/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Dudung"
    '        Me.RowCetakPayment.Description = "Loan Installment"
    '        Me.RowCetakPayment.Account_Description = "Gain Acc"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 25000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 8
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '        Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '        Me.RowCetakPayment.No_Document = "000008/SLR/01/14"
    '        Me.RowCetakPayment.Employee_Name = "Dudung"
    '        Me.RowCetakPayment.Description = "Payment"
    '        Me.RowCetakPayment.Account_Description = "BCA"
    '        Me.RowCetakPayment.Debet = 0
    '        Me.RowCetakPayment.Kredit = 1775000
    '        Me.RowCetakPayment.Position_Debet_Kredit = 1
    '        Me.RowCetakPayment.Primary_Salary = 8
    '        Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        ' ''end dudung  

    '        For i As Integer = 0 To 5
    '            Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '            Me.RowCetakPayment.No_Document = "00000" & i + 9 & "/SLR/01/14"
    '            Me.RowCetakPayment.Employee_Name = "Looping " & i + 1
    '            Me.RowCetakPayment.Description = "BASIC SALARY"
    '            Me.RowCetakPayment.Account_Description = "Salary Expense"
    '            Me.RowCetakPayment.Debet = 2000000
    '            Me.RowCetakPayment.Kredit = 0
    '            Me.RowCetakPayment.Position_Debet_Kredit = 0
    '            Me.RowCetakPayment.Primary_Salary = i + 9
    '            Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '            Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '            Me.RowCetakPayment.No_Document = "00000" & i + 9 & "/SLR/01/14"
    '            Me.RowCetakPayment.Employee_Name = "Looping " & i + 1
    '            Me.RowCetakPayment.Description = "Loan Installment"
    '            Me.RowCetakPayment.Account_Description = "Receivable Acc"
    '            Me.RowCetakPayment.Debet = 0
    '            Me.RowCetakPayment.Kredit = 200000
    '            Me.RowCetakPayment.Position_Debet_Kredit = 1
    '            Me.RowCetakPayment.Primary_Salary = i + 9
    '            Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '            Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '            Me.RowCetakPayment.No_Document = "00000" & i + 9 & "/SLR/01/14"
    '            Me.RowCetakPayment.Employee_Name = "Looping " & i + 1
    '            Me.RowCetakPayment.Description = "Loan Installment"
    '            Me.RowCetakPayment.Account_Description = "Gain Acc"
    '            Me.RowCetakPayment.Debet = 0
    '            Me.RowCetakPayment.Kredit = 25000
    '            Me.RowCetakPayment.Position_Debet_Kredit = 1
    '            Me.RowCetakPayment.Primary_Salary = i + 9
    '            Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)

    '            Me.RowCetakPayment = Me.TblCetakPaymentDummy.NewTblCetakRow
    '            Me.RowCetakPayment.No_Document = "00000" & i + 9 & "/SLR/01/14"
    '            Me.RowCetakPayment.Employee_Name = "Looping " & i + 1
    '            Me.RowCetakPayment.Description = "Payment"
    '            Me.RowCetakPayment.Account_Description = "BCA"
    '            Me.RowCetakPayment.Debet = 0
    '            Me.RowCetakPayment.Kredit = 1775000
    '            Me.RowCetakPayment.Position_Debet_Kredit = 1
    '            Me.RowCetakPayment.Primary_Salary = i + 9
    '            Me.TblCetakPaymentDummy.AddTblCetakRow(Me.RowCetakPayment)
    '        Next

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class