Option Explicit On
Option Strict Off

Public Class FrmMasterCurrency2_CekDefPrec
    Public PrimaryCurrency As Double = -1
    Public DefPrecBaru As Integer = 0

    Dim RetVal As Boolean = True
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TotalPhase As Integer = 9
    Dim MsgStr As String = "" 

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel data check?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.bgw1.IsBusy = True Then
                    Me.bgw1.CancelAsync()
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.Close()
                Else
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCurrency2_CekDefPrec_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.LblProgress.Text = ""
            Control.CheckForIllegalCrossThreadCalls = False
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            'cek ke po
            If Me.CekPO = False Then
                Me.RetVal = False
                Exit Sub
            End If

            'cek ke so
            If Me.CekSO = False Then
                Me.RetVal = False
                Exit Sub
            End If

            'cek ke quotation
            If Me.CekQuotation = False Then
                Me.RetVal = False
                Exit Sub
            End If

            'cek ke semua jurnal
            If Me.CekJournal = False Then
                Me.RetVal = False
                Exit Sub
            End If

            'cek salary
            If Me.PrimaryCurrency = 1 Then
                If Me.CekSalary = False Then
                    Me.RetVal = False
                    Exit Sub
                End If
            End If

            'monthly closing (hanya untuk IDR saja)

            'cek ke quick sr

            'cek ke quick pr

            'cek ke cndn

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekJournal() As Boolean
        Dim PrimaryJournalMax As Double = -1
        Dim DTA As New dseMasterCurrencyTableAdapters.DataTable4TableAdapter
        Dim TBL As New dseMasterCurrency.DataTable4DataTable
        Dim row As dseMasterCurrency.DataTable4Row 

        Dim RetVal As Nullable(Of Double)
        Dim StartPrimary, EndPrimary As Double
        Dim Span As Double = 1000
        Dim JmlData As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring 
            If Me.PrimaryCurrency = 1 Then
                RetVal = DTA.ScalarQuery_GetPrimaryMax1()
            Else
                RetVal = DTA.ScalarQuery_GetPrimaryMax(Me.PrimaryCurrency)
            End If
            If RetVal.HasValue = True Then
                If RetVal.Value > PrimaryJournalMax Then
                    PrimaryJournalMax = RetVal.Value
                End If
            End If

            If PrimaryJournalMax = -1 Then
                Return True
            Else
                'load data item Journal
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimaryJournalMax Then
                    EndPrimary = PrimaryJournalMax
                End If
                If Me.PrimaryCurrency = 1 Then
                    JmlData = DTA.ScalarQuery_GetJml1(PrimaryJournalMax).Value
                Else
                    JmlData = DTA.ScalarQuery_GetJml(PrimaryJournalMax, Me.PrimaryCurrency).Value
                End If

                'cek data item Journal
                Me.LblProgress.Text = "Phase 7 of " + Me.TotalPhase.ToString + " - Journal Header item check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimaryJournalMax
                    TBL.Rows.Clear()
                    If Me.PrimaryCurrency = 1 Then
                        DTA.FillBy1(TBL, EndPrimary, StartPrimary)
                    Else
                        DTA.Fill(TBL, EndPrimary, StartPrimary, Me.PrimaryCurrency)
                    End If

                    While TBL.Rows.Count > 0
                        row = TBL.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 7 of " + Me.TotalPhase.ToString + " - Journal Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 7 of " + Me.TotalPhase.ToString + " - Journal Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 7 of " + Me.TotalPhase.ToString + " - Journal Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.PrimaryCurrency = 1 Then
                            If Me.ObjCD.GetJmlPrecision(row.Debet) > Me.DefPrecBaru Then
                                MsgStr = "Inconsistent currency precission on Journal: " + row.Ref_Number
                                Return False
                            ElseIf Me.ObjCD.GetJmlPrecision(row.Kredit) > Me.DefPrecBaru Then
                                MsgStr = "Inconsistent currency precission on Journal: " + row.Ref_Number
                                Return False
                            End If
                        Else
                            If Me.ObjCD.GetJmlPrecision(row.Debet_Original) > Me.DefPrecBaru Then
                                MsgStr = "Inconsistent currency precission on Journal: " + row.Ref_Number
                                Return False
                            ElseIf Me.ObjCD.GetJmlPrecision(row.Credit_Original) > Me.DefPrecBaru Then
                                MsgStr = "Inconsistent currency precission on Journal: " + row.Ref_Number
                                Return False
                            End If
                        End If

                        TBL.Rows.Remove(row)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimaryJournalMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimaryJournalMax Then
                        EndPrimary = PrimaryJournalMax
                    End If
                End While
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekSalary() As Boolean
        Dim PrimarySalaryMax As Double = -1
        Dim DTA As New dseMasterCurrencyTableAdapters.DataTable5TableAdapter
        Dim TBL As New dseMasterCurrency.DataTable5DataTable
        Dim row As dseMasterCurrency.DataTable5Row
        Dim DTA1 As New dseMasterCurrencyTableAdapters.tbltrans_salaryTableAdapter
        Dim TBL1 As New dseMasterCurrency.tbltrans_salaryDataTable
        Dim row1 As dseMasterCurrency.tbltrans_salaryRow

        Dim RetVal As Nullable(Of Double)
        Dim StartPrimary, EndPrimary As Double
        Dim Span As Double = 1000
        Dim JmlData As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA1.ScalarQuery_GetPrimaryMax()
            If RetVal.HasValue = True Then
                If RetVal.Value > PrimarySalaryMax Then
                    PrimarySalaryMax = RetVal.Value
                End If
            End If
            If PrimarySalaryMax = -1 Then
                Return True
            Else
                'load data header Salary
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimarySalaryMax Then
                    EndPrimary = PrimarySalaryMax
                End If
                JmlData = DTA1.ScalarQuery_GetJml(PrimarySalaryMax).Value

                'cek data header Salary
                Me.LblProgress.Text = "Phase 8 of " + Me.TotalPhase.ToString + " - Salary Header data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimarySalaryMax
                    TBL1.Rows.Clear()
                    DTA1.Fill(TBL1, EndPrimary, StartPrimary)

                    While TBL1.Rows.Count > 0
                        row1 = TBL1.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 8 of " + Me.TotalPhase.ToString + " - Salary Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 8 of " + Me.TotalPhase.ToString + " - Salary Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 8 of " + Me.TotalPhase.ToString + " - Salary Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row1.Deduction) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Income) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Income_Tax) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Loan_Installment) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Penghasilan_Bruto) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Salary_Received) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row1.No_Transaksi
                            Return False
                        End If

                        TBL1.Rows.Remove(row1)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimarySalaryMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimarySalaryMax Then
                        EndPrimary = PrimarySalaryMax
                    End If
                End While

                'load data item Salary
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimarySalaryMax Then
                    EndPrimary = PrimarySalaryMax
                End If
                JmlData = DTA.ScalarQuery_GetJml(PrimarySalaryMax).Value

                'cek data item Salary
                Me.LblProgress.Text = "Phase 9 of " + Me.TotalPhase.ToString + " - Salary item data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimarySalaryMax
                    TBL.Rows.Clear()
                    DTA.Fill(TBL, EndPrimary, StartPrimary)

                    While TBL.Rows.Count > 0
                        row = TBL.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 9 of " + Me.TotalPhase.ToString + " - Salary Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 9 of " + Me.TotalPhase.ToString + " - Salary Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 9 of " + Me.TotalPhase.ToString + " - Salary Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row.Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Salary: " + row.No_Transaksi
                            Return False
                        End If

                        TBL.Rows.Remove(row)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimarySalaryMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimarySalaryMax Then
                        EndPrimary = PrimarySalaryMax
                    End If
                End While
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekQuotation() As Boolean
        Dim PrimaryQuotationMax As Double = -1
        Dim DTA As New dseMasterCurrencyTableAdapters.DataTable3TableAdapter
        Dim TBL As New dseMasterCurrency.DataTable3DataTable
        Dim row As dseMasterCurrency.DataTable3Row
        Dim DTA1 As New dseMasterCurrencyTableAdapters.tbltrans_quotationTableAdapter
        Dim TBL1 As New dseMasterCurrency.tbltrans_quotationDataTable
        Dim row1 As dseMasterCurrency.tbltrans_quotationRow

        Dim RetVal As Nullable(Of Double)
        Dim StartPrimary, EndPrimary As Double
        Dim Span As Double = 1000
        Dim JmlData As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA1.ScalarQuery_GetPrimaryMax(Me.PrimaryCurrency)
            If RetVal.HasValue = True Then
                If RetVal.Value > PrimaryQuotationMax Then
                    PrimaryQuotationMax = RetVal.Value
                End If
            End If
            If PrimaryQuotationMax = -1 Then
                Return True
            Else
                'load data header Quotation
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimaryQuotationMax Then
                    EndPrimary = PrimaryQuotationMax
                End If
                JmlData = DTA1.ScalarQuery_GetJml(Me.PrimaryCurrency, PrimaryQuotationMax).Value

                'cek data header Quotation
                Me.LblProgress.Text = "Phase 5 of " + Me.TotalPhase.ToString + " - Quotation Header data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimaryQuotationMax
                    TBL1.Rows.Clear()
                    DTA1.Fill(TBL1, EndPrimary, StartPrimary, Me.PrimaryCurrency)

                    While TBL1.Rows.Count > 0
                        row1 = TBL1.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 5 of " + Me.TotalPhase.ToString + " - Quotation Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 5 of " + Me.TotalPhase.ToString + " - Quotation Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 5 of " + Me.TotalPhase.ToString + " - Quotation Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row1.Sub_Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row1.No_Transaksi
                            Return False
                        End If

                        TBL1.Rows.Remove(row1)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimaryQuotationMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimaryQuotationMax Then
                        EndPrimary = PrimaryQuotationMax
                    End If
                End While

                'load data item Quotation
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimaryQuotationMax Then
                    EndPrimary = PrimaryQuotationMax
                End If
                JmlData = DTA.ScalarQuery_GetJml(PrimaryQuotationMax, Me.PrimaryCurrency).Value

                'cek data item Quotation
                Me.LblProgress.Text = "Phase 6 of " + Me.TotalPhase.ToString + " - Quotation item data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimaryQuotationMax
                    TBL.Rows.Clear()
                    DTA.Fill(TBL, EndPrimary, StartPrimary, Me.PrimaryCurrency)

                    While TBL.Rows.Count > 0
                        row = TBL.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 6 of " + Me.TotalPhase.ToString + " - Quotation Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 6 of " + Me.TotalPhase.ToString + " - Quotation Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 6 of " + Me.TotalPhase.ToString + " - Quotation Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row.Unit_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Disc_Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Nett_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on Quotation: " + row.No_Transaksi
                            Return False
                        End If

                        TBL.Rows.Remove(row)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimaryQuotationMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimaryQuotationMax Then
                        EndPrimary = PrimaryQuotationMax
                    End If
                End While
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekSO() As Boolean
        Dim PrimarySOMax As Double = -1
        Dim DTA As New dseMasterCurrencyTableAdapters.DataTable2TableAdapter
        Dim TBL As New dseMasterCurrency.DataTable2DataTable
        Dim row As dseMasterCurrency.DataTable2Row
        Dim DTA1 As New dseMasterCurrencyTableAdapters.tbltrans_soTableAdapter
        Dim TBL1 As New dseMasterCurrency.tbltrans_soDataTable
        Dim row1 As dseMasterCurrency.tbltrans_soRow

        Dim RetVal As Nullable(Of Double)
        Dim StartPrimary, EndPrimary As Double
        Dim Span As Double = 1000
        Dim JmlData As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA1.ScalarQuery_GetPrimaryMax(Me.PrimaryCurrency)
            If RetVal.HasValue = True Then
                If RetVal.Value > PrimarySOMax Then
                    PrimarySOMax = RetVal.Value
                End If
            End If
            If PrimarySOMax = -1 Then
                Return True
            Else
                'load data header SO
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimarySOMax Then
                    EndPrimary = PrimarySOMax
                End If
                JmlData = DTA1.ScalarQuery_GetJml(Me.PrimaryCurrency, PrimarySOMax).Value

                'cek data header SO
                Me.LblProgress.Text = "Phase 3 of " + Me.TotalPhase.ToString + " - SO Header data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimarySOMax
                    TBL1.Rows.Clear()
                    DTA1.Fill(TBL1, EndPrimary, StartPrimary, Me.PrimaryCurrency)

                    While TBL1.Rows.Count > 0
                        row1 = TBL1.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 3 of " + Me.TotalPhase.ToString + " - SO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 3 of " + Me.TotalPhase.ToString + " - SO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 3 of " + Me.TotalPhase.ToString + " - SO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row1.Sub_Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row1.No_Transaksi
                            Return False
                        End If

                        TBL1.Rows.Remove(row1)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimarySOMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimarySOMax Then
                        EndPrimary = PrimarySOMax
                    End If
                End While

                'load data item SO
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimarySOMax Then
                    EndPrimary = PrimarySOMax
                End If
                JmlData = DTA.ScalarQuery_GetJml(PrimarySOMax, Me.PrimaryCurrency).Value

                'cek data item SO
                Me.LblProgress.Text = "Phase 4 of " + Me.TotalPhase.ToString + " - SO item data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimarySOMax
                    TBL.Rows.Clear()
                    DTA.Fill(TBL, EndPrimary, StartPrimary, Me.PrimaryCurrency)

                    While TBL.Rows.Count > 0
                        row = TBL.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 4 of " + Me.TotalPhase.ToString + " - SO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 4 of " + Me.TotalPhase.ToString + " - SO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 4 of " + Me.TotalPhase.ToString + " - SO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row.Unit_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Disc_Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Nett_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on SO: " + row.No_Transaksi
                            Return False
                        End If

                        TBL.Rows.Remove(row)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimarySOMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimarySOMax Then
                        EndPrimary = PrimarySOMax
                    End If
                End While
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekPO() As Boolean
        Dim PrimaryPOMax As Double = -1
        Dim DTA As New dseMasterCurrencyTableAdapters.DataTable1TableAdapter
        Dim TBL As New dseMasterCurrency.DataTable1DataTable
        Dim row As dseMasterCurrency.DataTable1Row
        Dim DTA1 As New dseMasterCurrencyTableAdapters.tbltrans_poTableAdapter
        Dim TBL1 As New dseMasterCurrency.tbltrans_poDataTable
        Dim row1 As dseMasterCurrency.tbltrans_poRow

        Dim RetVal As Nullable(Of Double)
        Dim StartPrimary, EndPrimary As Double
        Dim Span As Double = 1000
        Dim JmlData As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA.ScalarQuery_GetPrimaryPOMax()
            If RetVal.HasValue = True Then
                If RetVal.Value > PrimaryPOMax Then
                    PrimaryPOMax = RetVal.Value
                End If
            End If
            If PrimaryPOMax = -1 Then
                Return True
            Else
                'load data header po
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimaryPOMax Then
                    EndPrimary = PrimaryPOMax
                End If
                JmlData = DTA1.ScalarQuery_GetJml(PrimaryPOMax, Me.PrimaryCurrency).Value

                'cek data header po
                Me.LblProgress.Text = "Phase 1 of " + Me.TotalPhase.ToString + " - PO Header data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimaryPOMax
                    TBL1.Rows.Clear()
                    DTA1.Fill(TBL1, Me.PrimaryCurrency, EndPrimary, StartPrimary)

                    While TBL1.Rows.Count > 0
                        row1 = TBL1.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 1 of " + Me.TotalPhase.ToString + " - PO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 1 of " + Me.TotalPhase.ToString + " - PO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 1 of " + Me.TotalPhase.ToString + " - PO Header data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row1.Sub_Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row1.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row1.Total) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row1.No_Transaksi
                            Return False
                        End If

                        TBL1.Rows.Remove(row1)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimaryPOMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimaryPOMax Then
                        EndPrimary = PrimaryPOMax
                    End If
                End While

                'load data item po
                StartPrimary = 0
                EndPrimary = StartPrimary + Span - 1
                If EndPrimary > PrimaryPOMax Then
                    EndPrimary = PrimaryPOMax
                End If
                JmlData = DTA.ScalarQuery_GetJml(PrimaryPOMax, Me.PrimaryCurrency).Value

                'cek data item po
                Me.LblProgress.Text = "Phase 2 of " + Me.TotalPhase.ToString + " - PO item data check"
                Me.prg1.Value = 0
                Me.prg1.Maximum = JmlData
                While EndPrimary <= PrimaryPOMax
                    TBL.Rows.Clear()
                    DTA.Fill(TBL, EndPrimary, StartPrimary, Me.PrimaryCurrency)

                    While TBL.Rows.Count > 0
                        row = TBL.Rows(0)
                        Me.prg1.Value += 1
                        If Me.prg1.Value > 25 Then
                            If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                                Me.LblProgress.Text = "Phase 2 of " + Me.TotalPhase.ToString + " - PO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                                Me.LblProgress.Text = "Phase 2 of " + Me.TotalPhase.ToString + " - PO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                            End If
                        Else
                            Me.LblProgress.Text = "Phase 2 of " + Me.TotalPhase.ToString + " - PO Item data check - Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString
                        End If

                        If Me.ObjCD.GetJmlPrecision(row.Unit_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Disc_Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Amount) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row.No_Transaksi
                            Return False
                        ElseIf Me.ObjCD.GetJmlPrecision(row.Nett_Price) > Me.DefPrecBaru Then
                            MsgStr = "Inconsistent currency precission on PO: " + row.No_Transaksi
                            Return False
                        End If

                        TBL.Rows.Remove(row)
                    End While

                    StartPrimary = EndPrimary + 1
                    If StartPrimary > PrimaryPOMax Then
                        Exit While
                    End If
                    EndPrimary = StartPrimary + Span - 1
                    If EndPrimary > PrimaryPOMax Then
                        EndPrimary = PrimaryPOMax
                    End If
                End While
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            If Me.RetVal = False Then
                MsgBox(Me.MsgStr, MsgBoxStyle.Information, "Information")
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Else
                Me.DialogResult = Windows.Forms.DialogResult.OK
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class