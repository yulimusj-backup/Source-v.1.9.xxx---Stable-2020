Option Explicit On
Option Strict Off

Public Class FrmTransaksiClosing
    Public UserCode As String = ""
    Dim Obj As New ClsTransaksiClosing
    Dim ExitStr As Boolean = False

    Dim ObjClosingInv As New ClsTransaksiMonthlyClosing
    Dim ObjJournal As New ClsTransaksiJournal
    Dim ObjStockInOut As New ClsTransaksiStockInOut
    Dim ObjStockBundling As New ClsTransaksiStockBundling
    Dim ObjStockTransfer As New ClsTransaksiStockTransfer


    Private Sub FrmTransaksiClosing_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If ExitStr = False Then
                If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiClosing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.Obj.ApakahBolehUbahConvertionDate(Me.ObjStockInOut, Me.ObjStockTransfer, Me.ObjStockBundling, Me.ObjClosingInv) = False Then
                Me.dtpConvertion.Enabled = False
            End If

            'cek apakah payroll convertion date bisa di ubah
            If Me.ApakahBolehUbahPayrollConvertionDate = False Then
                Me.DtpConventionDate_HR.Enabled = False
            End If

            Me.Obj.LoadData()
            Me.SalinKeField()
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahBolehUbahPayrollConvertionDate() As Boolean
        Try
            'tidak bisa ubah payroll convertion date, jika sudah ada transaksi berikut:
            '1. sudah ada loan
            If Me.Obj.GetJmlLoan > 0 Then
                Return False
            End If
            '2. sudah ada loan payment
            If Me.Obj.GetJmlLoanPayment > 0 Then
                Return False
            End If
            '3. sudah ada tax begining balance
            If Me.Obj.GetJmlTaxBeginingBalance > 0 Then
                Return False
            End If
            '4. sudah ada salary calculation
            If Me.Obj.GetJmlSalaryCalculation > 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.dtpConvertion.Value = .ConvertionDate
                Me.dtpClosingDate.Value = .ClosingDate
                Me.DtpConventionDate_HR.Value = .Payroll_ConventionDate
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObject()
        Try
            With Me.Obj
                .ConvertionDate = Me.dtpConvertion.Value.Date
                .ClosingDate = Me.dtpClosingDate.Value.Date
                .Payroll_ConventionDate = Me.DtpConventionDate_HR.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SalinKeObject()
                Me.Obj.SaveData(Me.UserCode)
                Me.ExitStr = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ExitStr = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExplainClosingDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExplainClosingDate.Click
        Dim ij As New FrmTransaksiClosing_Explain

        Try
            ij.TxtExplain.Text = "Defines journal locking date." + vbCrLf + _
            "Every journal dated before or on the closing date is locked. User can't edit or delete the respective journal. " + vbCrLf + _
            "User can't create new journal dated before or on closing date."
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnExplain_ConvertionDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExplain_ConvertionDate.Click
        Dim ij As New FrmTransaksiClosing_Explain
        Dim StatusMonthlyClosing As String = ""
        Dim StatusJournal As String = ""
        Dim StatusInventoryBeginingBalance As String = ""
        Dim StatusStockInOut As String = ""
        Dim StatusStockBundling As String = ""
        Dim StatusStockTransfer As String = ""

        Try
            If Me.ObjClosingInv.GetJml > 0 Then
                StatusMonthlyClosing = " (Status: found)"
            Else
                StatusMonthlyClosing = " (Status: not found)"
            End If
            If Me.ObjJournal.GetJml() > 0 Then
                StatusJournal = " (Status: found)"
            Else
                StatusJournal = " (Status: not found)"
            End If
            If Me.ObjStockInOut.GetJml_BeginingBalance > 0 Then
                StatusInventoryBeginingBalance = " (Status: found)"
                StatusStockInOut = " (Status: found)"
            Else
                StatusInventoryBeginingBalance = " (Status: not found)"
                If Me.ObjStockInOut.GetJml > 0 Then
                    StatusStockInOut = " (Status: found)"
                Else
                    StatusStockInOut = " (Status: not found)"
                End If
            End If
            If Me.ObjStockTransfer.GetJml() > 0 Then
                StatusStockTransfer = " (Status: found)"
            Else
                StatusStockTransfer = " (Status: not found)"
            End If
            If Me.ObjStockBundling.GetJml() > 0 Then
                StatusStockBundling = " (Status: found)"
            Else
                StatusStockBundling = " (Status: not found)"
            End If

            ij.TxtExplain.Text = "Defines SMESS start date." + vbCrLf + _
            "This date represent SMESS start date, every transaction must be date after this date. The transaction dated before this date must be consider as beginning balance." + vbCrLf + _
            "This date will be locked if there is one of the following transaction:" + vbCrLf + _
            "1. Monthly Closing" + StatusMonthlyClosing + vbCrLf + _
            "2. Journal" + StatusJournal + vbCrLf + _
            "3. Inventory Begining Balance" + StatusInventoryBeginingBalance + vbCrLf + _
            "4. Stock In/Out Transaction" + StatusStockInOut + vbCrLf + _
            "5. Stock Bundling Transaction" + StatusStockBundling + vbCrLf + _
            "6. Stock Transfer Transaction" + StatusStockTransfer

            ij.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnExplain_PayrollConvertionDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExplain_PayrollConvertionDate.Click
        Dim ij As New FrmTransaksiClosing_Explain
        Dim StatusLoan As String = ""
        Dim StatusLoanPayment As String = ""
        Dim StatusSalaryCalculation As String = ""
        Dim StatusTaxBeginingBalance As String = ""

        Try
            'load
            If Me.Obj.GetJmlLoan > 0 Then
                StatusLoan = " (Status: found)"
            Else
                StatusLoan = " (Status: not found)"
            End If
            'loan payment
            If Me.Obj.GetJmlLoanPayment() > 0 Then
                StatusLoanPayment = " (Status: found)"
            Else
                StatusLoanPayment = " (Status: not found)"
            End If
            'tax begining balance
            If Me.Obj.GetJmlTaxBeginingBalance > 0 Then
                StatusTaxBeginingBalance = " (Status: found)"
            Else
                StatusTaxBeginingBalance = " (Status: not found)"
            End If
            'salary calculation
            If Me.Obj.GetJmlSalaryCalculation() > 0 Then
                StatusSalaryCalculation = " (Status: found)"
            Else
                StatusSalaryCalculation = " (Status: not found)"
            End If

            ij.TxtExplain.Text = "Defines SMESS payroll start date." + vbCrLf + _
            "This date represent SMESS payroll start date, every payroll transaction must be date after this date. The transaction dated before this date must be consider as beginning balance." + vbCrLf + _
            "This date will be locked if there is one of the following transaction:" + vbCrLf + _
            "1. Loan" + StatusLoan + vbCrLf + _
            "2. Loan Payment" + StatusLoanPayment + vbCrLf + _
            "3. Tax Begining Balance" + StatusTaxBeginingBalance + vbCrLf + _
            "4. Salary Calculation" + StatusSalaryCalculation 

            ij.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class