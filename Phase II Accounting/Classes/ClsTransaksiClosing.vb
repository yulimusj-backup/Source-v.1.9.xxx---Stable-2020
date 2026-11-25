Option Explicit On
Option Strict Off

Public Class ClsTransaksiClosing
    Private mConvertionDate As Date
    Private mClosingDate As Date
    Private mPayroll_ConventionDate As Date
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date
    Private mPrimary As Double = -1

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property ConvertionDate() As Date
        Get
            Return Me.mConvertionDate
        End Get
        Set(ByVal value As Date)
            mConvertionDate = value
        End Set
    End Property

    Public Property ClosingDate() As Date
        Get
            Return Me.mClosingDate
        End Get
        Set(ByVal value As Date)
            mClosingDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property Payroll_ConventionDate() As Date
        Get
            Return mPayroll_ConventionDate
        End Get
        Set(ByVal value As Date)
            mPayroll_ConventionDate = value
        End Set
    End Property

    Public Function GetClosingDate() As Date
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetClosingDate.HasValue = True Then
                Return DTA.ScalarQuery_GetClosingDate.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetConvertionDate() As Date
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                Return DTA.ScalarQuery_GetConvertionDate.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPayrollConvertionDate() As Date
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPayrollConventionDate()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub LoadData()
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim TBL As New DseTransaksiClosing.tbltrans_accounting_closingDataTable
        Dim row As DseTransaksiClosing.tbltrans_accounting_closingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL)
            row = TBL.Rows(0)
            Me.mClosingDate = row.Closing_Date
            Me.mConvertionDate = row.Convertion_Date
            Me.mPayroll_ConventionDate = row.Convertion_Date_Payroll
            Me.mModifiedBy = row.Modified_By
            Me.mModifiedDate = row.Modified_Date
            Me.mPrimary = row.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub SaveData(ByVal vUserCode As String)
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_accounting_closingTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Update(Me.mClosingDate, Me.mConvertionDate, vUserCode, Now, Me.mPayroll_ConventionDate, Me.mPrimary)
            Me.mModifiedDate = DTA.ScalarQuery_GetModifiedDate(Me.mPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function ApakahBolehUbahConvertionDate(ByVal vObjStockInOut As ClsTransaksiStockInOut, ByVal vObjStockTransfer As ClsTransaksiStockTransfer, _
    ByVal vObjStockBundling As ClsTransaksiStockBundling, ByVal vObjInvClosing As ClsTransaksiMonthlyClosing) As Boolean
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh melakukan perubahan data jika:
            '1. sudah ada jurnal yang dimasukkan
            If DTA.ScalarQuery_GetJml.HasValue = True Then
                If DTA.ScalarQuery_GetJml.Value = 0 Then
                    'Return True
                Else
                    Return False
                End If
            End If

            'jika sudah ada closing bulan 0
            If vObjInvClosing.GetJml > 0 Then
                Return False
            End If

            'jika sudah ada transaksi beginning balance
            If vObjStockInOut.GetJml_BeginingBalance > 0 Then
                Return False
            End If

            'jika sudah ada transaksi stock in/out, bundling, transfer
            If vObjStockInOut.GetJml > 0 Then
                Return False
            End If
            If vObjStockBundling.GetJml > 0 Then
                Return False
            End If
            If vObjStockTransfer.GetJml > 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlLoan() As Double
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_loanTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlLoanPayment() As Double
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlLoanPayment
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlSalaryCalculation() As Double
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_salaryTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlTaxBeginingBalance() As Double
        Dim DTA As New DseTransaksiClosingTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
