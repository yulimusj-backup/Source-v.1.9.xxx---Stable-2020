Option Explicit On
Option Strict Off

Imports Microsoft.VisualBasic.FileIO

Public Class FrmTransaksiSalesReturnBalanceManual_Import
    Public UserCode As String = ""
    Public BackCOlorAll As Color = Color.LightSeaGreen

    Dim FileToLoad As String = ""
    Dim ColumnNumber As Integer = 10
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSalesBalance As New ClsTransaksiSalesBalance
    Dim ObjClosing As New ClsTransaksiClosing
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ConvertionDate As Date = #1/1/1950#

    Dim StatusResult As EnumStatus = EnumStatus.IncorrectFileFormat

    Dim CustCode, CustName, InvoiceNo, Curr, ARAccount, Remarks, Exception, ARAccountDefault As String
    Dim InvDate, DueDate As Date
    Dim Amount, ExchangeRate As Decimal
    Dim CustCodeIdx, InvoiceNoIDX, CurIdx, ARAccountIdx, RemarksIdx, InvDateIdx, DueDateIdx, AmountIdx, ExchangeRateIdx As Integer

    Dim DTACust As New DseTransaksiSalesBalance_ImportTableAdapters.tblmaster_customerTableAdapter
    Dim TBLCust As New DseTransaksiSalesBalance_Import.tblmaster_customerDataTable
    Dim rowCust As DseTransaksiSalesBalance_Import.tblmaster_customerRow
    Dim DTACurr As New DseTransaksiSalesBalance_ImportTableAdapters.tblmaster_currencyTableAdapter
    Dim TBLCurr As New DseTransaksiSalesBalance_Import.tblmaster_currencyDataTable
    Dim rowCurr As DseTransaksiSalesBalance_Import.tblmaster_currencyRow
    Dim DTAAkun As New DseTransaksiSalesBalance_ImportTableAdapters.tblmaster_akunTableAdapter
    Dim TBLAkun As New DseTransaksiSalesBalance_Import.tblmaster_akunDataTable
    Dim rowAkun As DseTransaksiSalesBalance_Import.tblmaster_akunRow
    Dim DefPrec As Integer = 0
    Dim FinishLoad As Boolean = False
    Dim AdaException As Boolean = False

    Enum EnumStatus As Integer
        Done
        IncorrectFileFormat
    End Enum

    Private Sub btnFindFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFile.Click

        Try
            'cari file yang berisikan data piutang
            Me.OpenFileDialog1.FileName = ""
            If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.Rows.Clear()
                Me.SetForm(False)

                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgw1.RunWorkerAsync()
                Me.prgBar1.Style = ProgressBarStyle.Marquee
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please specify data to import.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.AdaException = True Then
                MsgBox("Please fix exception before start import process.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Start import?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.SetForm(False)
                Me.bgw2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesBalanceManual_Import_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboSeparator.SelectedIndex = 0

            Me.CustCodeIdx = 0
            Me.InvoiceNoIDX = 2
            Me.InvDateIdx = 3
            Me.DueDateIdx = 4
            Me.CurIdx = 5
            Me.AmountIdx = 6
            Me.ExchangeRateIdx = 7
            Me.ARAccountIdx = 8
            Me.RemarksIdx = 9

            Me.DTAAkun.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTACurr.Connection.ConnectionString = Me.DTAAkun.Connection.ConnectionString
            Me.DTACust.Connection.ConnectionString = Me.DTAAkun.Connection.ConnectionString

            Me.ConvertionDate = Me.ObjClosing.GetConvertionDate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Dim LineNumber As Double = 0
        Dim Progres As Integer = 0

        Try
            Me.ObjCD.UpdateThreading()
            Me.AdaException = False

            Me.TBLAkun.Rows.Clear()
            Me.TBLCurr.Rows.Clear()
            Me.TBLCust.Rows.Clear()

            Me.DTAAkun.Fill(Me.TBLAkun)
            Me.DTACurr.Fill(Me.TBLCurr)
            Me.DTACust.Fill(Me.TBLCust)

            'pastikan formatnya tepat: jumlah kolom tepat
            Me.FileToLoad = Me.OpenFileDialog1.FileName
            Me.TextBox1.Text = Me.FileToLoad

            Dim filename As String = Me.FileToLoad
            Dim fields As String()
            Dim delimiter As String = Me.cboSeparator.Text
            Using parser As New TextFieldParser(filename)
                parser.SetDelimiters(delimiter)
                While Not parser.EndOfData
                    LineNumber += 1
                    Me.LblProgress.Text = "Loading data: " + Me.ObjCD.DigitGrouping(LineNumber, 0)

                    ' Read in the fields for the current line
                    fields = parser.ReadFields
                    ' Add code here to use data in fields variable.
                    If LineNumber = 1 Then
                        If fields.Length <> ColumnNumber Then
                            Me.StatusResult = EnumStatus.IncorrectFileFormat
                            Exit Sub
                        End If
                    Else
                        Exception = ""
                        CustCode = fields(Me.CustCodeIdx).Trim
                        InvoiceNo = fields(Me.InvoiceNoIDX).Trim
                        If IsDate(fields(Me.InvDateIdx)) = True Then
                            InvDate = CDate(fields(Me.InvDateIdx))
                        Else
                            InvDate = #1/1/1950#
                        End If
                        If IsDate(fields(Me.DueDateIdx)) = True Then
                            DueDate = CDate(fields(Me.DueDateIdx))
                        Else
                            DueDate = #1/1/1950#
                        End If
                        Curr = fields(Me.CurIdx).Trim
                        ARAccount = fields(Me.ARAccountIdx).Trim
                        Remarks = fields(Me.RemarksIdx).Trim
                        If IsNumeric(fields(Me.AmountIdx)) = True Then
                            Amount = fields(Me.AmountIdx)
                        Else
                            Amount = 0
                        End If
                        If IsNumeric(fields(Me.ExchangeRateIdx)) = True Then
                            ExchangeRate = fields(Me.ExchangeRateIdx)
                        Else
                            ExchangeRate = 0
                        End If

                        'masukkan file ke dalam data tabel, sambil diperiksa
                        '1. Cust Code harus ada
                        Me.TBLCust.DefaultView.RowFilter = "[Customer Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(CustCode) + "'"
                        If Me.TBLCust.DefaultView.Count = 0 Then
                            If Exception = "" Then
                                Exception = "Customer not found"
                            Else
                                Exception = Exception + ", " + "Customer not found"
                            End If
                            CustName = ""
                        Else
                            rowCust = Me.TBLCust.DefaultView(0).Row
                            CustName = rowCust.Customer_Name
                        End If
                        '2. Invoice No tidak boleh kosong, tidak boleh duplikasi dengan baris yang sudah ada, tidak boleh duplikasi dengan data sales balance yang sudah ada
                        If InvoiceNo = "" Then
                            If Exception = "" Then
                                Exception = "Incorrect invoice number"
                            Else
                                Exception = Exception + ", " + "Incorrect invoice number"
                            End If
                        End If
                        If Me.GetIdx(InvoiceNo) <> -1 Then
                            If Exception = "" Then
                                Exception = "Invoice number duplication"
                            Else
                                Exception = Exception + ", " + "Invoice number duplication"
                            End If
                        Else
                            If Me.ObjSalesBalance.ApakahSudahAda(InvoiceNo) = True Then
                                If Exception = "" Then
                                    Exception = "Invoice number duplication"
                                Else
                                    Exception = Exception + ", " + "Invoice number duplication"
                                End If
                            End If
                        End If
                        '3. Invoice Date: format harus tepat dan nilainya < Convertion Date
                        If InvDate = #1/1/1950# Then
                            If Exception = "" Then
                                Exception = "Invalid invoice date"
                            Else
                                Exception = Exception + ", " + "Invalid invoice date"
                            End If
                        End If
                        If InvDate.Date >= Me.ConvertionDate.Date Then
                            If Exception = "" Then
                                Exception = "Invoice date must be < " + Format(Me.ConvertionDate, "dd MMM yyyy")
                            Else
                                Exception = Exception + ", " + "Invoice date must be < " + Format(Me.ConvertionDate, "dd MMM yyyy")
                            End If
                        End If
                        '4. Due date: format harus tepat
                        If DueDate = #1/1/1950# Then
                            If Exception = "" Then
                                Exception = "Invalid due date"
                            Else
                                Exception = Exception + ", " + "Invalid due date"
                            End If
                        End If
                        '5. Currency harus ada
                        Me.TBLCurr.DefaultView.RowFilter = "[Currency Code] = '" + Me.Curr + "'"
                        If Me.TBLCurr.DefaultView.Count = 0 Then
                            If Exception = "" Then
                                Exception = "Currency not found"
                            Else
                                Exception = Exception + ", " + "Currency not found"
                            End If
                            DefPrec = 0
                            ARAccountDefault = ""
                        Else
                            rowCurr = Me.TBLCurr.DefaultView(0).Row
                            DefPrec = rowCurr.Default_Precision
                            ARAccountDefault = rowCurr.AR_Account
                        End If
                        '6. Amount harus negative, digit grouping sesuai dengan currency
                        Amount = Me.ObjCD.DigitGrouping(Amount, Me.DefPrec)
                        If Amount >= 0 Then
                            If Exception = "" Then
                                Exception = "Specify negative number for amount data"
                            Else
                                Exception = Exception + ", " + "Specify negative number for amount data"
                            End If
                        End If
                        '7. AR account harus sesuai dengan akun ar currency, jika tidak maka akun tersebut currencynya harus sama, tipenya harus account payable, dan harus ada di master akun
                        If ARAccount <> ARAccountDefault Then
                            Me.TBLAkun.DefaultView.RowFilter = "[Account Code] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(ARAccount) + "'"
                            If Me.TBLAkun.DefaultView.Count = 0 Then
                                If Exception = "" Then
                                    Exception = "AR Account not found"
                                Else
                                    Exception = Exception + ", " + "AR Account not found"
                                End If
                            Else
                                rowAkun = Me.TBLAkun.DefaultView(0).Row
                                If rowAkun.Tipe_Perkiraan2 <> ClsMasterAkun.EnumTipeAkun2.Account_Receivable Then
                                    If Exception = "" Then
                                        Exception = "Incorrect AR account type"
                                    Else
                                        Exception = Exception + ", " + "Incorrect AR account type"
                                    End If
                                End If
                                If rowAkun.Currency_Code <> Curr Then
                                    If Exception = "" Then
                                        Exception = "Inconsistent currency code"
                                    Else
                                        Exception = Exception + ", " + "Inconsistent currency code"
                                    End If
                                End If
                            End If
                        End If

                        If Exception <> "" Then
                            Me.AdaException = True
                        End If

                        '8. Tambahkan baris dgv
                        If Progres + 1 = 99 Then
                            Progres = 0
                        Else
                            Progres += 1
                        End If
                        Me.FinishLoad = False
                        Me.bgw1.ReportProgress(Progres)
                        While Me.FinishLoad = False
                            Threading.Thread.Sleep(50)
                        End While
                    End If
                End While
            End Using

            Me.StatusResult = EnumStatus.Done
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIdx(ByVal vInvoiceNumber As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColInvoiceNo.Index).Value = vInvoiceNumber Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub bgw1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgw1.ProgressChanged
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.Rows.Add()

            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.Amount, Me.DefPrec)
            rowDGV.Cells(Me.ColARAccount.Index).Value = Me.ARAccount
            rowDGV.Cells(Me.ColCurrency.Index).Value = Me.Curr
            rowDGV.Cells(Me.ColCustName.Index).Value = Me.CustName
            rowDGV.Cells(Me.ColCustomerCode.Index).Value = Me.CustCode
            rowDGV.Cells(Me.ColDueDate.Index).Value = Me.DueDate
            rowDGV.Cells(Me.ColException.Index).Value = Me.Exception
            rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(Me.ExchangeRate, 4)
            rowDGV.Cells(Me.ColInvoiceDate.Index).Value = Me.InvDate
            rowDGV.Cells(Me.ColInvoiceNo.Index).Value = Me.InvoiceNo
            rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColRemarks.Index).Value = Me.Remarks

            If rowDGV.Cells(Me.ColException.Index).Value.ToString.Trim <> "" Then
                rowDGV.DefaultCellStyle.BackColor = Color.MistyRose
            End If

            Me.FinishLoad = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.ColInvoiceDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColDueDate.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.prgBar1.Style = ProgressBarStyle.Continuous
            Me.LblProgress.Text = ""
            Me.SetForm(True)

            If Me.StatusResult = EnumStatus.IncorrectFileFormat Then
                MsgBox("Incorrect file format, there should be " + Me.ColumnNumber.ToString + " columns in file.", _
                               MsgBoxStyle.Information, "Information")
            Else
                If Me.AdaException = False Then
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
                    MsgBox("Data loaded.", MsgBoxStyle.Information, "Information")
                Else
                    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.dgvItem.DefaultCellStyle.BackColor
                    MsgBox("Data loaded. There exception detected, please fix before import.", MsgBoxStyle.Information, "Information")
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Dim Frm As FrmTransaksiSalesBalanceManual
        Dim rowDGV As DataGridViewRow

        Try
            Me.ObjCD.UpdateThreading()

            Me.prgBar1.Value = 0
            Me.prgBar1.Maximum = Me.dgvItem.RowCount
            For Each rowDGV In Me.dgvItem.Rows
                Me.prgBar1.Value += 1
                Me.LblProgress.Text = "Data " + Me.prgBar1.Value.ToString + " of " + Me.prgBar1.Maximum.ToString

                Frm = New FrmTransaksiSalesBalanceManual
                Frm.UserCode = Me.UserCode

                Frm.Import_Amount = rowDGV.Cells(Me.ColAmount.Index).Value
                Frm.Import_ARAccount = rowDGV.Cells(Me.ColARAccount.Index).Value
                Frm.Import_Curr = rowDGV.Cells(Me.ColCurrency.Index).Value
                Frm.Import_CustCode = rowDGV.Cells(Me.ColCustomerCode.Index).Value
                Frm.Import_CustName = rowDGV.Cells(Me.ColCustName.Index).Value
                Frm.Import_DueDate = rowDGV.Cells(Me.ColDueDate.Index).Value
                Frm.Import_ExchangeRate = rowDGV.Cells(Me.ColExchangeRate.Index).Value
                Frm.Import_InvDate = rowDGV.Cells(Me.ColInvoiceDate.Index).Value
                Frm.Import_InvoiceNo = rowDGV.Cells(Me.ColInvoiceNo.Index).Value
                Frm.Import_Remarks = rowDGV.Cells(Me.ColRemarks.Index).Value

                Frm.ShowDialog()

                Frm.Dispose()
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Try
            Me.prgBar1.Value = 0
            Me.LblProgress.Text = "..."
            Me.SetForm(True)

            Me.TextBox1.Text = ""
            Me.dgvItem.Rows.Clear()

            MsgBox("Data import completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetForm(ByVal vEnable As Boolean)
        Try
            Me.ControlBox = vEnable
            Me.btnFindFile.Enabled = vEnable
            Me.btnStart.Enabled = vEnable
            Me.cboSeparator.Enabled = vEnable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class