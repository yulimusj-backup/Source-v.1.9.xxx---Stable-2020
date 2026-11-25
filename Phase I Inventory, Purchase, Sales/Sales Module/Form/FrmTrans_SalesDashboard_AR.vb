Option Explicit On
Option Strict Off

Imports Core_Module
Partial Public Class FrmTrans_SalesDashboard
    Dim ObjInvoiceBalance As New Inventory_Module.ClsTransaksiSalesBalance
    Dim TBLSalesBalance As New Inventory_Module.DseTransaksiSalesBalance.balance_invoice_viewDataTable
    Dim FrmSales As New Inventory_Module.FrmTransaksiJournal_Sales
    Dim TBLSRBalance As New Inventory_Module.DseTransaksiSalesBalance.balance_sr_viewDataTable

    Private Sub btnAR_Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAR_Filter.Click
        Try
            Me.dgvException.AllowUserToAddRows = False
            Me.UpdateAR()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvException.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvException.SelectionChanged
        Dim DefPrec As Integer = 0
        Dim TotalAR, TotalAROverDue As Decimal
        Dim CurCode As String = ""

        Try
            If Me.dgvException.SelectedCells.Count = 0 Then
                Me.dgvSalesBalance.Rows.Clear()
                Me.LblCreditLimit.Text = ": -"
                Me.LblTotalARAbove.Text = ": -"
                Me.LblTotalOverDue.Text = ": -"
            Else
                Me.LoadDetailCustomer(Me.dgvException.Item(Me.ColExc_PrimaryCust.Index, Me.dgvException.SelectedCells(0).RowIndex).Value)
                'isikan data credit limit
                DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
                CurCode = Me.ObjCur.GetCurrencyCode(1)
                Me.LblCreditLimit.Text = ": " + CurCode + " " + _
                ObjCd.DigitGrouping(Me.ObjCustomer.GetCreditLimit(Me.dgvException.Item(Me.ColExc_PrimaryCust.Index, Me.dgvException.SelectedCells(0).RowIndex).Value), DefPrec)
                TotalAR = 0
                TotalAROverDue = 0
                For Each rowDGV In Me.dgvSalesBalance.Rows
                    TotalAR += CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                    If rowDGV.DefaultCellStyle.ForeColor = Color.Red Then
                        TotalAROverDue += CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value)
                    End If
                Next
                Me.LblTotalARAbove.Text = ": " + CurCode + " " + Me.ObjCd.DigitGrouping(TotalAR, DefPrec)
                Me.LblTotalOverDue.Text = ": " + CurCode + " " + Me.ObjCd.DigitGrouping(TotalAROverDue, DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDetailCustomer(ByVal vPrimaryCustomer As Double)
        Dim i As Double = 0
        Dim AdaForeignCurrency As Boolean = False
        Try
            Me.ObjInvoiceBalance.IsiDGVBalanceTanpaJurnal(Me.dgvSalesBalance, Me.TBLSalesBalance, Me.FrmSales, Me.ObjCd, _
            vPrimaryCustomer, Me.ObjCur.GetDecimal_DefaultCurrency, Me.TBLSRBalance)

            'buang yg belum sampai due date
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSalesBalance.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            While i <= Me.dgvSalesBalance.Rows.Count - 1
                rowDGV = Me.dgvSalesBalance.Rows(i)
                If CDate(rowDGV.Cells(Me.ColDueDateInvoice.Index).Value).Date < Today Then
                    rowDGV.DefaultCellStyle.ForeColor = Color.Red
                End If
                i += 1
            End While

            For Each rowDGV In Me.dgvSalesBalance.Rows
                If rowDGV.Cells(Me.ColPrimaryCurrencyBalanceSales.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                    AdaForeignCurrency = True
                    Exit For
                End If
            Next

            If AdaForeignCurrency = True Then
                Me.ColExhangeRateBalance.Visible = True
                Me.ColBalanceSales.Visible = True
            Else
                Me.ColExhangeRateBalance.Visible = False
                Me.ColBalanceSales.Visible = False
            End If
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvException_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvException.Sorted
        Try
            Me.ObjCd.SetNumber(Me.dgvException, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateAR()
        Dim CreditLimit As Decimal = 0
        Dim PrimaryCust As Double
        Dim OverDue As Boolean = False
        Dim OverLimit As Boolean = False
        Dim DTA1 As New Inventory_Module.DseTransaksiSalesBalanceTableAdapters.tblmaster_customerTableAdapter
        Dim TBL1 As New Inventory_Module.DseTransaksiSalesBalance.tblmaster_customerDataTable
        Dim row1 As Inventory_Module.DseTransaksiSalesBalance.tblmaster_customerRow
        Dim Lanjut As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA1.Fill(TBL1)
            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvException.Rows.Clear()

            'ada 2 kelompok customer yang bermasalah:
            For Each row1 In TBL1.Rows
                If Me.TxtAR_CustCode.Text.Trim <> "" Then
                    If Me.TxtAR_CustCode.Text = row1.Customer_Code Then
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                Else
                    Lanjut = True
                End If
                If Lanjut = True Then
                    PrimaryCust = row1.Primary_Customer
                    OverDue = False
                    OverLimit = False

                    '1. yang punya ar diatas credit limit
                    CreditLimit = Me.ObjCustomer.GetCreditLimit(row1.Primary_Customer)
                    If CreditLimit <> 0 Then
                        If CreditLimit < Me.ObjSO.GetAR(PrimaryCust) Then
                            OverLimit = True
                        End If
                    End If

                    '2. yang punya ar overdue
                    If Me.ObjSO.ApakahAdaAROverdue(PrimaryCust, Today) = True Then
                        OverDue = True
                    End If

                    If OverDue = True Or OverLimit = True Then
                        Me.dgvException.Rows.Add()
                        rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.ColExc_AboveCreditLimit.Index).Value = OverLimit
                            .Cells(Me.ColExc_Customer.Index).Value = Me.ObjCustomer.GetCustName(PrimaryCust)
                            .Cells(Me.ColExc_No.Index).Value = Me.dgvException.Rows.Count
                            .Cells(Me.ColExc_OverdueInvoice.Index).Value = OverDue
                            .Cells(Me.ColExc_PrimaryCust.Index).Value = PrimaryCust
                        End With
                    End If
                End If
            Next

            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvException.ClearSelection()
            If Me.dgvException.Rows.Count > 0 Then
                Me.dgvException.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Sub TxtAR_CustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAR_CustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtAR_CustCode.Text) = True Then
                    Me.TxtAR_PrimaryCust.Text = Me.ObjCustomer.GetPrimary(Me.TxtAR_CustCode.Text)
                    Me.TxtAR_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtAR_PrimaryCust.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtAR_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtAR_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtAR_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtAR_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtAR_PrimaryCust.Text)
                        Me.TxtAR_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtAR_PrimaryCust.Text)
                    Else
                        Me.TxtAR_CustCode.Clear()
                        Me.TxtAR_TxtCustName.Clear()
                        Me.TxtAR_PrimaryCust.Text = "-1"
                        Me.TxtAR_CustCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtAR_CustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAR_CustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtAR_CustCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtAR_CustCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.TxtAR_CustCode.Text)
                    If PrimaryBaru <> Me.TxtAR_PrimaryCust.Text Then
                        Me.TxtAR_PrimaryCust.Text = PrimaryBaru
                        Me.TxtAR_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtAR_PrimaryCust.Text)
                    End If
                Else
                    If Me.TxtAR_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtAR_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtAR_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtAR_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtAR_PrimaryCust.Text)
                        Me.TxtAR_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtAR_PrimaryCust.Text)
                    Else
                        Me.TxtAR_CustCode.Clear()
                        Me.TxtAR_TxtCustName.Clear()
                        Me.TxtAR_PrimaryCust.Text = "-1"
                        Me.TxtAR_CustCode.Focus()
                    End If
                End If
            Else
                Me.TxtAR_CustCode.Clear()
                Me.TxtAR_TxtCustName.Clear()
                Me.TxtAR_PrimaryCust.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnAR_FindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAR_FindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtAR_CustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtAR_CustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtAR_PrimaryCust.Text = i.PrimaryCustomer
                Me.TxtAR_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtAR_PrimaryCust.Text)
                Me.TxtAR_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtAR_PrimaryCust.Text)
            Else
                Me.TxtAR_CustCode.Clear()
                Me.TxtAR_TxtCustName.Clear()
                Me.TxtAR_PrimaryCust.Text = "-1"
                Me.TxtAR_CustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class
