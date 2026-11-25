Option Explicit On
Option Strict Off

Partial Public Class FrmMainSMESS
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj_AOPBalance As New AccModule.ClsTransaksiPurchaseBalance
    Dim TBL_AOP_Balance As New AccModule.DseTransaksiPurchaseBalance.balance_purchase_viewDataTable
    Dim ObjInvoiceBalance As New AccModule.ClsTransaksiSalesBalance
    Dim TBLSalesBalance As New AccModule.DseTransaksiSalesBalance.balance_invoice_viewDataTable
    Dim TBLSRBalance As New AccModule.DseTransaksiSalesReturnBalance.balance_sr_viewDataTable
    Dim ij As New AccModule.FrmTransaksiJournal_Purchase
    Dim FrmSales As New AccModule.FrmTransaksiJournal_Sales
    Dim ObjCust As New Core_Module.clsMasterCustomer

    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim DefPrec As Integer = 0
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim Buffer As String = ""
    Dim Buffer1 As String = ""

    Private Sub LoadStockIn()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexSupplier(ByVal vPrimarySupplier As Double) As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvRekap.Rows
                If rowDGV.Cells(Me.ColRekap_Primary.Index).Value = vPrimarySupplier Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub LoadRekapSupplier()
        Dim rowDGV As DataGridViewRow
        Dim rowDGV1 As DataGridViewRow
        Dim i As Double = 0

        Try
            If Me.dgvPurchaseBalance.ColumnCount = 0 Then
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvRekap.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvRekap.ClearSelection()
            Me.dgvRekap.Rows.Clear()
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.ColRekap_Code.HeaderText = "Supplier Code"
            Me.ColRekap_Nama.HeaderText = "Supplier Name"
            Me.ColCreateCollectionLetter.Visible = False
            Me.ColCreateCollectionLetter.HeaderText = "Collection Letter"
            Me.ColCreateCollectionLetter.DefaultCellStyle.NullValue = ""

            Me.Obj_AOPBalance.IsiDGVBalance(Me.DgvRekapAP, Me.TBL_AOP_Balance, ij, Me.ObjCD, _
            Me.TxtReminderAPAR_FilterPrimary.Text, -1, Me.DefPrec)

            'buang yg belum sampai due date
            i = 0
            If Me.chkShowOverDue.Checked = True Then
                While i <= Me.DgvRekapAP.Rows.Count - 1
                    rowDGV = Me.DgvRekapAP.Rows(i)
                    If CDate(rowDGV.Cells(Me.ColDueDatePurchase.Index).Value).Date > Me.dtpOverDue.Value.Date Then
                        Me.DgvRekapAP.Rows.Remove(rowDGV)
                    Else
                        'If Me.chkAPARReminder_Filter.Checked = True Then
                        '    If CDbl(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value) <> CDbl(Me.TxtReminderAPAR_FilterPrimary.Text) Then
                        '        Me.DgvRekapAP.Rows.Remove(rowDGV)
                        '    Else
                        '        i += 1
                        '    End If
                        'Else
                        i += 1
                        'End If
                    End If
                End While
                'Else
                '    'buang yang tidak termasuk supplier yang ditunjuk
                '    If Me.chkAPARReminder_Filter.Checked = True Then
                '        While i <= Me.DgvRekapAP.Rows.Count - 1
                '            rowDGV = Me.DgvRekapAP.Rows(i)
                '            If CDbl(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value) <> CDbl(Me.TxtReminderAPAR_FilterPrimary.Text) Then
                '                Me.DgvRekapAP.Rows.Remove(rowDGV)
                '            Else
                '                i += 1
                '            End If
                '        End While
                '    End If
            End If


            'tambahkan ke dalam dgv rekap
            For Each rowDGV In Me.DgvRekapAP.Rows
                i = Me.GetIndexSupplier(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value)

                If i = -1 Then
                    'tambahkan baris baru
                    Me.dgvRekap.Rows.Add()
                    rowDGV1 = Me.dgvRekap.Rows(Me.dgvRekap.Rows.Count - 1)
                    With rowDGV1
                        .Cells(Me.ColRekap_Amount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColBalancePurchase.Index).Value), Me.DefPrec)
                        .Cells(Me.ColRekap_Primary.Index).Value = rowDGV.Cells(Me.ColPrimarySupplier.Index).Value
                        .Cells(Me.ColRekap_Code.Index).Value = Me.ObjSupplier.GetSupplierCode(.Cells(Me.ColRekap_Primary.Index).Value)
                        .Cells(Me.ColRekap_Nama.Index).Value = Me.ObjSupplier.GetSupplierName(.Cells(Me.ColRekap_Primary.Index).Value)
                    End With
                Else
                    'tambahkan amount saja
                    rowDGV1 = Me.dgvRekap.Rows(i)
                    With rowDGV1
                        .Cells(Me.ColRekap_Amount.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColRekap_Amount.Index).Value) + _
                        CDec(rowDGV.Cells(Me.ColBalancePurchase.Index).Value), Me.DefPrec)
                    End With
                End If
            Next

            'sort
            If Me.dgvRekap.Rows.Count > 0 Then
                Me.dgvRekap.Sort(Me.ColRekap_Code, System.ComponentModel.ListSortDirection.Ascending)
            End If

            Me.dgvRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvRekap.Rows.Count > 0 Then
                Me.dgvRekap.Rows(0).Selected = True
                Me.LoadDetailSupplier(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, Me.dgvRekap.SelectedRows(0).Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LoadRekapCust()
        Dim rowDGV As DataGridViewRow
        Dim rowDGV1 As DataGridViewRow
        Dim i As Double = 0

        Try
            If Me.dgvSalesBalance.ColumnCount = 0 Then
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvRekap.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvRekap.ClearSelection()
            Me.dgvRekap.Rows.Clear()
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.ColRekap_Code.HeaderText = "Customer Code"
            Me.ColRekap_Nama.HeaderText = "Customer Name"
            Me.ColCreateCollectionLetter.Visible = True

            Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvRekapAR, Me.TBLSalesBalance, Me.FrmSales, Me.ObjCD, Me.TxtReminderAPAR_FilterPrimary.Text, -1, Me.DefPrec, Me.TBLSRBalance)

            'buang yg belum sampai due date
            i = 0
            If Me.chkShowOverDue.Checked = True Then
                While i <= Me.dgvRekapAR.Rows.Count - 1
                    rowDGV = Me.dgvRekapAR.Rows(i)
                    If CDate(rowDGV.Cells(Me.ColDueDateInvoice.Index).Value).Date > Me.dtpOverDue.Value.Date Then
                        Me.dgvRekapAR.Rows.Remove(rowDGV)
                    Else
                        'If Me.chkAPARReminder_Filter.Checked = True Then
                        '    If CDbl(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value) <> CDbl(Me.TxtReminderAPAR_FilterPrimary.Text) Then
                        '        Me.DgvRekapAP.Rows.Remove(rowDGV)
                        '    Else
                        '        i += 1
                        '    End If
                        'Else
                        i += 1
                        'End If
                    End If
                End While
                'Else
                '    If Me.chkAPARReminder_Filter.Checked = True Then
                '        While i <= Me.dgvRekapAR.Rows.Count - 1
                '            rowDGV = Me.dgvRekapAR.Rows(i)
                '            If CDbl(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value) <> CDbl(Me.TxtReminderAPAR_FilterPrimary.Text) Then
                '                Me.DgvRekapAP.Rows.Remove(rowDGV)
                '            Else
                '                i += 1
                '            End If
                '        End While
                '    End If
            End If

            'tambahkan ke dalam dgv rekap
            For Each rowDGV In Me.dgvRekapAR.Rows
                i = Me.GetIndexSupplier(rowDGV.Cells(Me.ColPrimaryCust1.Index).Value)

                If i = -1 Then
                    'tambahkan baris baru
                    Me.dgvRekap.Rows.Add()
                    rowDGV1 = Me.dgvRekap.Rows(Me.dgvRekap.Rows.Count - 1)
                    With rowDGV1
                        .Cells(Me.ColRekap_Amount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value), Me.DefPrec)
                        .Cells(Me.ColRekap_Primary.Index).Value = rowDGV.Cells(Me.ColPrimaryCust1.Index).Value
                        .Cells(Me.ColRekap_Code.Index).Value = Me.ObjCust.GetCustomerCode(.Cells(Me.ColRekap_Primary.Index).Value)
                        .Cells(Me.ColRekap_Nama.Index).Value = Me.ObjCust.GetCustName(.Cells(Me.ColRekap_Primary.Index).Value)
                    End With
                Else
                    'tambahkan amount saja
                    rowDGV1 = Me.dgvRekap.Rows(i)
                    With rowDGV1
                        .Cells(Me.ColRekap_Amount.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColRekap_Amount.Index).Value) + _
                        CDec(rowDGV.Cells(Me.ColBalanceSales.Index).Value), Me.DefPrec)
                    End With
                End If
            Next

            'sort
            If Me.dgvRekap.Rows.Count > 0 Then
                Me.dgvRekap.Sort(Me.ColRekap_Code, System.ComponentModel.ListSortDirection.Ascending)
            End If

            Me.dgvRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvRekap.Rows.Count > 0 Then
                Me.dgvRekap.Rows(0).Selected = True
                Me.LoadDetailCustomer(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, Me.dgvRekap.SelectedRows(0).Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LoadRekap()
        Try
            If Me.rboAP.Checked = True Then
                Me.LoadRekapSupplier()
                Me.dgvPurchaseBalance.BringToFront()
            Else
                Me.LoadRekapCust()
                Me.dgvSalesBalance.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDetailSupplier(ByVal vPrimarySupplier As Double)
        Dim i As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim AdaForeignCurrency As Boolean = False

        Try
            Me.Obj_AOPBalance.IsiDGVBalance(Me.dgvPurchaseBalance, Me.TBL_AOP_Balance, ij, Me.ObjCD, _
            vPrimarySupplier, -1, Me.DefPrec)
            If Me.chkShowOverDue.Checked = True Then
                While i <= Me.dgvPurchaseBalance.Rows.Count - 1
                    rowDGV = Me.dgvPurchaseBalance.Rows(i)
                    If CDate(rowDGV.Cells(Me.ColDueDatePurchase.Index).Value).Date > Me.dtpOverDue.Value.Date Then
                        Me.dgvPurchaseBalance.Rows.Remove(rowDGV)
                    Else
                        i += 1
                    End If
                End While
            End If

            For Each rowDGV In Me.dgvPurchaseBalance.Rows
                If rowDGV.Cells(Me.ColPrimaryCurrencyBalance.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                    AdaForeignCurrency = True
                    Exit For
                End If
            Next
            If AdaForeignCurrency = True Then
                Me.ColExchangeRateBalance.Visible = True
            Else
                Me.ColExchangeRateBalance.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDetailCustomer(ByVal vPrimaryCustomer As Double)
        Dim i As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim AdaForeignCurrency As Boolean = False

        Try
            Me.ObjInvoiceBalance.IsiDGVBalance(Me.dgvSalesBalance, Me.TBLSalesBalance, Me.FrmSales, Me.ObjCD, _
            vPrimaryCustomer, -1, Me.DefPrec, Me.TBLSRBalance)

            'buang yg belum sampai due date
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSalesBalance.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If Me.chkShowOverDue.Checked = True Then
                While i <= Me.dgvSalesBalance.Rows.Count - 1
                    rowDGV = Me.dgvSalesBalance.Rows(i)
                    If CDate(rowDGV.Cells(Me.ColDueDateInvoice.Index).Value).Date > Me.dtpOverDue.Value.Date Then
                        Me.dgvSalesBalance.Rows.Remove(rowDGV)
                    Else
                        i += 1
                    End If
                End While
            End If

            For Each rowDGV In Me.dgvSalesBalance.Rows
                If rowDGV.Cells(Me.ColPrimaryCurrencyBalanceSales.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                    AdaForeignCurrency = True
                    Exit For
                End If
            Next

            If AdaForeignCurrency = True Then
                Me.ColExhangeRateBalance.Visible = True
            Else
                Me.ColExhangeRateBalance.Visible = False
            End If
            Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BayarAP(ByVal vPrimarySupplier As Double, ByVal vTipeDokumentAcuan As AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vPrimaryDokumentAcuan As Double)
        Dim ij As AccModule.FrmTransaksiJournal_Purchase
        Dim j As Windows.Forms.Form
        Dim Ketemu As Boolean = False

        Try
            'cari anak dari main form
            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Purchase)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Ketemu = True
                        Exit For
                    End If
                End If
            Next

            'jika tidak ada maka bikin ada baru
            If Ketemu = False Then
                ij = New AccModule.FrmTransaksiJournal_Purchase
                ij.UserCode = FrmLogin.UsernameTextBox.Text
                ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
                ' Set the caption to be unique.
                ij.BackColorAll = Core_Module.GetSMESSColor
                ij.MdiParent = Me.ParentForm
                ij.Show()
            End If

            'pay
            If IsNothing(ij) = False Then
                ij.Pay(vPrimarySupplier, vTipeDokumentAcuan, vPrimaryDokumentAcuan)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BayarAR(ByVal vPrimaryCust As Double, ByVal vTipeDokumentAcuan As AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance, _
    ByVal vPrimaryDokumentAcuan As Double)
        Dim ij As AccModule.FrmTransaksiJournal_Sales
        Dim j As Windows.Forms.Form
        Dim Ketemu As Boolean = False

        Try
            'cari anak dari main form
            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Sales)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Ketemu = True
                        Exit For
                    End If
                End If
            Next

            'jika tidak ada maka bikin ada baru
            If Ketemu = False Then
                ij = New AccModule.FrmTransaksiJournal_Sales
                ij.UserCode = FrmLogin.UsernameTextBox.Text
                ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
                ' Set the caption to be unique.
                ij.BackColorAll = Core_Module.GetSMESSColor
                ij.MdiParent = Me.ParentForm
                ij.Show()
            End If

            'pay
            If IsNothing(ij) = False Then
                ij.Pay(vPrimaryCust, vTipeDokumentAcuan, vPrimaryDokumentAcuan)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BikinAOP()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvRekap_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvRekap.CellClick
        Dim ij As New AccModule.FrmTransaksiTagihan2
        Try
            If e.ColumnIndex = Me.ColRekap_Print.Index Then
                If Me.rboAR.Checked = True Then
                    Me.PrintAR_All(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, e.RowIndex).Value)
                Else
                    Me.PrintAP_All(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, e.RowIndex).Value)
                End If
            ElseIf e.ColumnIndex = Me.ColCreateCollectionLetter.Index Then
                If Me.dgvRekap.Rows.Count > 0 Then
                    ij.BackColorAll = Me.BackColorAll
                    ij.UserCode = FrmLogin.UsernameTextBox.Text
                    ij.listOfCust.Add(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, e.RowIndex).Value)
                    ij.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub dgvRekap_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRekap.SelectionChanged
        Try
            If Me.dgvRekap.SelectedRows.Count = 0 Then
                Me.dgvPurchaseBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvPurchaseBalance.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvPurchaseBalance.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                Me.dgvSalesBalance.Rows.Clear()
                Me.dgvPurchaseBalance.Rows.Clear()

                Me.dgvPurchaseBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvSalesBalance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Else
                If Me.rboAP.Checked = True Then
                    Me.LoadDetailSupplier(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, Me.dgvRekap.SelectedRows(0).Index).Value)
                ElseIf Me.rboAR.Checked = True Then
                    Me.LoadDetailCustomer(Me.dgvRekap.Item(Me.ColRekap_Primary.Index, Me.dgvRekap.SelectedRows(0).Index).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboAP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboAP.CheckedChanged
        Try
            If Me.rboAP.Checked = True Then
                Me.chkAPARReminder_Filter.Text = "Supplier Code"
                Me.chkAPARReminder_Filter.Checked = False
                Me.LoadRekap()
                Me.cboPrintOption.Visible = True
                Me.cboPrintOption.TabStop = True
                Me.llbCreateCollectionLetter.Visible = False
                Me.ColCreateCollectionLetter.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboAR_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboAR.CheckedChanged
        Try
            If Me.rboAR.Checked = True Then
                Me.chkAPARReminder_Filter.Text = "Customer Code"
                Me.chkAPARReminder_Filter.Checked = False
                Me.LoadRekap()
                Me.cboPrintOption.Visible = False
                Me.cboPrintOption.TabStop = False
                Me.llbCreateCollectionLetter.Visible = True
                Me.ColCreateCollectionLetter.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PnlReminder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles PnlReminder.GotFocus
        Try
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.chkAPARReminder_Filter.Checked = False
                Me.LoadRekap()
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
            Me.LoadRekap()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowOverDue_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowOverDue.CheckedChanged
        Try
            If Me.chkShowOverDue.Checked = True Then
                Me.dtpOverDue.Enabled = True
            Else
                Me.dtpOverDue.Enabled = False
            End If
            Me.LoadRekap()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpOverDue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpOverDue.Enter
        Try
            Me.Buffer = Me.dtpOverDue.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpOverDue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpOverDue.Leave
        Try
            If CDate(Me.Buffer).Date <> Me.dtpOverDue.Value.Date Then
                If Me.dtpOverDue.Enabled = True Then
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPurchaseBalance_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPurchaseBalance.CellContentClick
        Dim rowDGV As DataGridViewRow

        Try
            If Me.Col_Purchase_Pay.Index = e.ColumnIndex Then
                rowDGV = Me.dgvPurchaseBalance.Rows(e.RowIndex)
                Me.BayarAP(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value, rowDGV.Cells(Me.ColTipeDocumentAcuan_Purchase.Index).Value, _
                rowDGV.Cells(Me.Col_PrimaryDocumentAcuanPurchase.Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesBalance_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSalesBalance.CellContentClick
        Dim rowDGV As DataGridViewRow

        Try
            If Me.ColPayInv.Index = e.ColumnIndex Then
                rowDGV = Me.dgvSalesBalance.Rows(e.RowIndex)
                Me.BayarAR(rowDGV.Cells(Me.ColPrimarySupplier.Index).Value, rowDGV.Cells(Me.ColTipeDocumentAcuan_Sales.Index).Value, _
                rowDGV.Cells(Me.ColPrimaryDocumenAcuan_Sales.Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAPARReminder_Filter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAPARReminder_Filter.CheckedChanged
        Try
            If Me.chkAPARReminder_Filter.Checked = False Then
                Me.TxtReminderAPAR_FilterCode.Text = ""
                Me.TxtReminderAPAR_FilterDesc.Text = ""
                Me.TxtReminderAPAR_FilterPrimary.Text = "-1"

                Me.TxtReminderAPAR_FilterCode.Enabled = False
                Me.TxtReminderAPAR_FilterDesc.Enabled = False
                Me.BtnReminderAPAR_Find.Enabled = False
            Else
                Me.TxtReminderAPAR_FilterCode.Enabled = True
                Me.TxtReminderAPAR_FilterDesc.Enabled = True
                Me.BtnReminderAPAR_Find.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReminderAPAR_FilterCode.Enter
        Try
            Me.Buffer = Me.TxtReminderAPAR_FilterCode.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReminderAPAR_FilterCode.GotFocus
        TxtReminderAPAR_FilterCode.BackColor = BackColorAll
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReminderAPAR_FilterCode.LostFocus
        TxtReminderAPAR_FilterCode.BackColor = Me.TxtReminderAPAR_FilterPrimary.BackColor
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtReminderAPAR_FilterCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.rboAP.Checked = False Then
                    Me.TxtReminderAPAR_FilterCode_Customer_KeyDown()
                Else
                    Me.TxtReminderAPAR_FilterCode_Supplier_KeyDown()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReminderAPAR_FilterCode.Leave
        Try
            If Me.rboAP.Checked = False Then
                Me.TxtReminderAPAR_FilterCode_Customer_Leave()
            Else
                Me.TxtReminderAPAR_FilterCode_Supplier_Leave()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnReminderAPAR_Find_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReminderAPAR_Find.Click
        Try
            If Me.rboAP.Checked = False Then
                Me.BtnReminderAPAR_FindCustomer()
            Else
                Me.BtnReminderAPAR_FindSupplier()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Customer_KeyDown()
        Try
            If Me.ObjCustomer.ApakahSudahAda(Me.TxtReminderAPAR_FilterCode.Text) = True Then
                Me.TxtReminderAPAR_FilterPrimary.Text = Me.ObjCustomer.GetPrimary(Me.TxtReminderAPAR_FilterCode.Text)
                Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjCustomer.GetCustName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.llbCreateCollectionLetter.Focus()
                Me.LoadRekap()
            Else
                'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                Me.BtnReminderAPAR_Find.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Customer_Leave()
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtReminderAPAR_FilterCode.Text) = True Then
                    Me.TxtReminderAPAR_FilterPrimary.Text = Me.ObjCustomer.GetPrimary(Me.TxtReminderAPAR_FilterCode.Text)
                    Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjCustomer.GetCustName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Else
                    If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtReminderAPAR_FilterCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtReminderAPAR_FilterPrimary.Text = i.PrimaryCustomer
                        Me.TxtReminderAPAR_FilterCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtReminderAPAR_FilterPrimary.Text)
                        Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjCustomer.GetCustName(Me.TxtReminderAPAR_FilterPrimary.Text)
                    Else
                        Me.TxtReminderAPAR_FilterCode.Clear()
                        Me.TxtReminderAPAR_FilterDesc.Clear()
                        Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
                    End If
                End If
            Else
                Me.TxtReminderAPAR_FilterCode.Clear()
                Me.TxtReminderAPAR_FilterDesc.Clear()
                Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
            End If

            If Me.TxtReminderAPAR_FilterCode.Text <> Me.Buffer Then
                Me.LoadRekap()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnReminderAPAR_FindCustomer()
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                i.txtCari.Text = Me.TxtReminderAPAR_FilterCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtReminderAPAR_FilterPrimary.Text = i.PrimaryCustomer
                Me.TxtReminderAPAR_FilterCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjCustomer.GetCustName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.llbCreateCollectionLetter.Focus()
            Else
                Me.TxtReminderAPAR_FilterCode.Clear()
                Me.TxtReminderAPAR_FilterDesc.Clear()
                Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
                Me.TxtReminderAPAR_FilterCode.Focus()
            End If

            Me.LoadRekap()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Supplier_KeyDown()
        Try
            If Me.ObjSupplier.ApakahSudahAda(Me.TxtReminderAPAR_FilterCode.Text) = True Then
                Me.TxtReminderAPAR_FilterPrimary.Text = Me.ObjSupplier.GetPrimary(Me.TxtReminderAPAR_FilterCode.Text)
                Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.llbCreateCollectionLetter.Focus()
                Me.LoadRekap()
            Else
                'ada 2 Supplier dengan kode yang sama atau tidak ada Supplier yang dimaksud
                Me.BtnReminderAPAR_Find.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReminderAPAR_FilterCode_Supplier_Leave()
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.TxtReminderAPAR_FilterCode.Text) = True Then
                    Me.TxtReminderAPAR_FilterPrimary.Text = Me.ObjSupplier.GetPrimary(Me.TxtReminderAPAR_FilterCode.Text)
                    Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Else
                    If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtReminderAPAR_FilterCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtReminderAPAR_FilterPrimary.Text = i.PrimarySupplier
                        Me.TxtReminderAPAR_FilterCode.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtReminderAPAR_FilterPrimary.Text)
                        Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtReminderAPAR_FilterPrimary.Text)
                    Else
                        Me.TxtReminderAPAR_FilterCode.Clear()
                        Me.TxtReminderAPAR_FilterDesc.Clear()
                        Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
                    End If
                End If
            Else
                Me.TxtReminderAPAR_FilterCode.Clear()
                Me.TxtReminderAPAR_FilterDesc.Clear()
                Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
            End If

            If Me.TxtReminderAPAR_FilterCode.Text <> Me.Buffer Then
                Me.LoadRekap()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnReminderAPAR_FindSupplier()
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtReminderAPAR_FilterCode.Text <> "" Then
                i.txtCari.Text = Me.TxtReminderAPAR_FilterCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtReminderAPAR_FilterPrimary.Text = i.PrimarySupplier
                Me.TxtReminderAPAR_FilterCode.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.TxtReminderAPAR_FilterDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtReminderAPAR_FilterPrimary.Text)
                Me.llbCreateCollectionLetter.Focus()
            Else
                Me.TxtReminderAPAR_FilterCode.Clear()
                Me.TxtReminderAPAR_FilterDesc.Clear()
                Me.TxtReminderAPAR_FilterPrimary.Text = "-1"
                Me.TxtReminderAPAR_FilterCode.Focus()
            End If

            Me.LoadRekap()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chkAPARReminder_Filter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAPARReminder_Filter.Enter
        Try
            Me.Buffer = Me.chkAPARReminder_Filter.Checked.ToString
            Me.Buffer1 = Me.TxtReminderAPAR_FilterPrimary.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAPARReminder_Filter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAPARReminder_Filter.Leave
        Try
            If Me.Buffer <> Me.chkAPARReminder_Filter.Checked.ToString Then
                If Me.Buffer1 <> Me.TxtReminderAPAR_FilterPrimary.Text Then
                    Me.LoadRekap()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrintAP_All(ByVal vPrimarySupp As Double)
        Dim frmLaporan As New AccModule.FrmLaporanReminder
        Dim isAllSupplier As Boolean = True
        Dim tbl As New AccModule.DseTransaksiPurchaseBalance.balance_purchase_viewDataTable
        Dim objApReport As New AccModule.ClsTransaksiPurchaseBalance
        Dim TipeCetakan As AccModule.ClsTransaksiPurchaseBalance.EnumCetakanInvoiceNo = AccModule.ClsTransaksiPurchaseBalance.EnumCetakanInvoiceNo.AOPNo
        Dim IsAll As Boolean = False

        Try
            If vPrimarySupp = -1 Then
                IsAll = True
            End If
            frmLaporan.Text = "Accounts Payable Due Report"
            Me.Cursor = Cursors.WaitCursor
            If Me.cboPrintOption.SelectedIndex = 0 Then
                TipeCetakan = AccModule.ClsTransaksiPurchaseBalance.EnumCetakanInvoiceNo.SupplierInv
            Else
                TipeCetakan = AccModule.ClsTransaksiPurchaseBalance.EnumCetakanInvoiceNo.AOPNo
            End If
            If objApReport.CetakLaporanAPDueDate(Me.chkShowOverDue.Checked, Me.dtpOverDue.Value, IsAll, _
                    vPrimarySupp, tbl, 1, frmLaporan.CrvReport, TipeCetakan) = True Then
                frmLaporan.ShowDialog()
                frmLaporan.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tbl.Dispose()
            frmLaporan.Dispose()
            Me.Cursor = Cursors.Default
            objApReport = Nothing
        End Try
    End Sub

    Private Sub PrintAR_All(ByVal vPrimaryCust As Double)
        Dim frmLaporan As New AccModule.FrmLaporanReminder
        Dim isAllCustomer As Boolean = True
        Dim tbl As New AccModule.DseTransaksiSalesBalance.balance_invoice_viewDataTable
        Dim objArReport As New AccModule.ClsTransaksiSalesBalance
        Dim IsAll As Boolean = False

        Try
            If vPrimaryCust = -1 Then
                IsAll = True
            End If
            Me.Cursor = Cursors.WaitCursor
            frmLaporan.Text = "Accounts Receivable Due Report"
            If objArReport.CetakLaporanARDueDate(Me.chkShowOverDue.Checked, Me.dtpOverDue.Value, IsAll, _
                    vPrimaryCust, tbl, 1, frmLaporan.CrvReport) = True Then
                frmLaporan.ShowDialog()
                frmLaporan.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tbl.Dispose()
            frmLaporan.Dispose()
            Me.Cursor = Cursors.Default
            objArReport = Nothing
        End Try
    End Sub

    Private Sub llbPrint_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbPrint.LinkClicked

        Try
            If Me.chkAPARReminder_Filter.Checked = False Then
                If Me.rboAR.Checked = True Then
                    Me.PrintAR_All(-1)
                Else
                    Me.PrintAP_All(-1)
                End If
            Else
                If Me.TxtReminderAPAR_FilterPrimary.Text = "" Or Me.TxtReminderAPAR_FilterPrimary.Text = "-1" Then
                    If Me.rboAR.Checked = True Then
                        MsgBox("Please specify supplier.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Please specify customer.", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TxtReminderAPAR_FilterPrimary.Focus()
                    Exit Sub
                End If
                    If Me.rboAR.Checked = True Then
                        Me.PrintAR_All(Me.TxtReminderAPAR_FilterPrimary.Text)
                    Else
                        Me.PrintAP_All(Me.TxtReminderAPAR_FilterPrimary.Text)
                    End If
                End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCreateCollectionLetter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCreateCollectionLetter.LinkClicked
        Dim ij As New AccModule.FrmTransaksiTagihan2
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvRekap.Rows.Count > 0 Then
                ij.UserCode = FrmLogin.UsernameTextBox.Text
                ij.BackColorAll = Me.BackColorAll
                For Each rowDGV In Me.dgvRekap.Rows
                    ij.listOfCust.Add(rowDGV.Cells(Me.ColRekap_Primary.Index).Value)
                Next
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
