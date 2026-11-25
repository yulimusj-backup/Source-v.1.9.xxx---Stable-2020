Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS_Payment
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public ParentForm_1 As FrmTrans_POS
    Public Obj As ClsTrans_POS

    Dim FontSubTotal As New Font("Microsoft Sans Serif", 35, FontStyle.Regular)
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim ObjPayment As New ClsMast_PaymentOption
    Dim ObjSession As New ClsTrans_Session
    Dim Buffer As String = ""
    Dim BufferGrandTotal As Decimal = 0
    Dim mrOK As Boolean = False

    Dim DefPrec As Integer
    Dim PrimaryPaymentOption As Double = -1
    Dim TBLPayment As New DseMast_Payment_Option.tblmaster_payment_option1_3DataTable
    Dim DTAPayment As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option1_3TableAdapter
    Dim MultiPayment As Boolean = False
    Dim EraseRowIndex As Integer = 0
    Dim IsMultiPayment As Boolean = False

    Private Sub SetStartPosition()
        Try
            'isikan data total sales

            'isikan data tipe payment - pada cbo dan pada dgv

            'kosongkan data
            Me.TxtAccountNumber.Clear()
            Me.TxtApprovalCode.Clear()
            Me.TxtTotalPayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
            Me.TxtTotalPayment2.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrec)
            Me.TxtChange.Text = Me.ObjCD.DigitGrouping(Me.TxtTotalSales.Text, Me.DefPrec)
            Me.TxtChange2.Text = Me.ObjCD.DigitGrouping(Me.TxtTotalSales.Text, Me.DefPrec)
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_Payment_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.mrOK = False Then

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_Payment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim Hasil As New MsgBoxResult
        Try
            If ObjSession.GetStatus(Me.ParentForm_1.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.ParentForm_1.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If e.KeyCode = Keys.F1 Then
                Hasil = MsgBox("Do you want to reset payment?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    'Me.SetStartPosition()
                    Me.spc.Panel2Collapsed = True
                    Me.spc.Panel1Collapsed = False
                    Me.cboPaymentTipe.Focus()
                    Me.MultiPayment = False
                    Me.LblKeterangan.Text = "F1 - Single Payment, F2 - Multiple Payment"
                    '-----
                    Me.TxtTotalPayment.Text = "0"
                    Me.TxtChange.Text = "0"
                    Me.TxtRemarks.Text = ""
                    Me.TxtAccountNumber.Text = ""
                    Me.TxtApprovalCode.Text = ""
                    Me.TxtPrimaryPayment.Text = "-1"
                    Me.cboPaymentTipe.DataSource = Nothing
                    Me.cboPaymentTipe.Items.Clear()
                    Me.PrimaryPaymentOption = ObjGeneralSet.GetPrimaryPayment()
                    ObjPayment.IsiCBOPaymentOption(Me.cboPaymentTipe, Me.TBLPayment, DTAPayment, Me.PrimaryPaymentOption)
                    '-----
                    Me.HitungHeaderPayment(Me.ParentForm_1.Obj, Me.ParentForm_1)
                    Me.TxtTotalPayment.Focus()
                    Me.IsMultiPayment = False
                End If
            ElseIf e.KeyCode = Keys.F2 Then
                Hasil = MsgBox("Do you want to reset payment?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    'Me.SetStartPosition()
                    Me.spc.Panel2Collapsed = False
                    Me.spc.Panel1Collapsed = True
                    Me.dgvItem.Focus()
                    Me.MultiPayment = True
                    Me.LblKeterangan.Text = "F1 - Single Payment, F2 - Multiple Payment, F3 - Input New Payment"
                    '-----
                    Me.TxtTotalPayment2.Text = "0"
                    Me.TxtChange2.Text = "0"
                    Me.dgvItem.Rows.Clear()
                    '-----
                    Me.cboPaymentTipe.DataSource = Nothing
                    Me.cboPaymentTipe.Items.Clear()
                    Me.PrimaryPaymentOption = ObjGeneralSet.GetPrimaryPayment()
                    ObjPayment.IsiCBOPaymentOption(Me.cboPaymentTipe, Me.TBLPayment, DTAPayment, Me.PrimaryPaymentOption)
                    Me.cboPaymentTipe.SelectedIndex = 0
                    '-----
                    Me.HitungHeaderPayment(Me.ParentForm_1.Obj, Me.ParentForm_1)
                    Me.TxtTotalPayment.Focus()
                    Me.SesuaikanHeight(True)
                    Me.IsMultiPayment = True
                End If
            ElseIf e.KeyCode = Keys.F3 Then
                If Me.spc.Panel1Collapsed = True Then
                    Dim ij As New FrmTrans_POS_Payment_Input
                    ij.ParentForm_2 = Me
                    ij.ShowDialog()
                End If
            ElseIf e.KeyCode = Keys.F5 Then
                Me.btnOK.PerformClick()
            ElseIf e.KeyCode = Keys.F6 Then
                Me.btnCancel.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungHeaderPayment(ByVal vObjPOS As ClsTrans_POS, ByVal i As FrmTrans_POS)
        Try
            With vObjPOS
                Me.TxtTotalSales.Text = ObjCD.DigitGrouping(.Total, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.TxtDiscString.Text = .DiscStr
                Me.TxtDisc.Text = ObjCD.DigitGrouping(.Discount, Me.DefPrec, ClsCekData.EnumTipeRounding.Middlepoint)
                Me.TxtSubTotal.Text = ObjCD.DigitGrouping(.Subtotal, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.TxtTaxPercentage.Text = ObjCD.DigitGrouping(.TaxPersen, 2, Me.ParentForm_1.RoundingTipe)
                Me.TxtTaxAmount.Text = ObjCD.DigitGrouping(.TaxAmount, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.TxtRounding.Text = ObjCD.DigitGrouping(.Rounding, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.TxtBankCharge.Text = ObjCD.DigitGrouping(.BankCharge, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(.GrandTotal, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Me.Label3.Font = Me.FontSubTotal
                Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                'cek approval number & approval code 
                If ObjPayment.GetFillApprovalCode(Me.TxtPrimaryPayment.Text) = 1 Then
                    Me.TxtAccountNumber.Enabled = True
                    Me.TxtApprovalCode.Enabled = True
                    Me.TxtAccountNumber.Clear()
                    Me.TxtApprovalCode.Clear()
                Else
                    Me.TxtAccountNumber.Enabled = False
                    Me.TxtApprovalCode.Enabled = False
                    Me.TxtAccountNumber.Clear()
                    Me.TxtApprovalCode.Clear()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_Payment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DefPrec = Me.ParentForm_1.DefPrec
            Me.mrOK = False
            Me.spc.Panel2Collapsed = True
            Me.spc.Panel1Collapsed = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MultiPayment = False
            '-----
            Me.PrimaryPaymentOption = Me.ObjGeneralSet.GetPrimaryPayment()
            ObjPayment.IsiCBOPaymentOption(Me.cboPaymentTipe, Me.TBLPayment, DTAPayment, Me.PrimaryPaymentOption, True)
            If ObjPayment.GetFillApprovalCode(Me.TxtPrimaryPayment.Text) > 0 Then
                Me.TxtAccountNumber.Enabled = True
                Me.TxtApprovalCode.Enabled = True
            Else
                Me.TxtAccountNumber.Enabled = False
                Me.TxtApprovalCode.Enabled = False
            End If
            'input header Payment
            Me.HitungHeaderPayment(Me.ParentForm_1.Obj, Me.ParentForm_1)
            Me.TxtTotalPayment.Focus()
            Me.IsMultiPayment = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalPayment.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtTotalPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalPayment.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtTotalPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungChargeSinglePayment()
        Dim Angka As Decimal = 0
        Try
            Angka = Me.TxtTotalPayment.Text - Me.TxtGrandTotal.Text
            If Angka <= 0 Then
                Angka = 0
            End If
            Me.TxtChange.Text = ObjCD.DigitGrouping(Angka, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalPayment.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtTotalPayment, sender, e, 0, Me.ObjCD)
            'Hitung Change
            Me.HitungChargeSinglePayment()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SesuaikanHeight(Optional ByVal vEnable As Boolean = False)
        Try
            If vEnable = False Then
                If Me.MultiPayment = False Then
                    Me.Width = 585
                    Me.Height = 450
                Else
                    Me.Width = 585
                    Me.Height = 507
                End If
            Else
                Me.Width = 585
                Me.Height = 512
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPaymentTipe_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPaymentTipe.GotFocus
        Try
            Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPaymentTipe_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPaymentTipe.SelectedIndexChanged
        Dim row As DseMast_Payment_Option.tblmaster_payment_option1_3Row
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Try
            If Me.cboPaymentTipe.SelectedIndex = -1 Then
                Me.TxtPrimaryPayment.Text = "-1"
                Me.TxtAccountNumber.Enabled = False
                Me.TxtApprovalCode.Enabled = False
                Me.SesuaikanHeight(False)
            Else
                row = Me.TBLPayment.DefaultView(Me.cboPaymentTipe.SelectedIndex).Row
                Me.TxtPrimaryPayment.Text = row.primary
                If ObjPayment.GetFillApprovalCode(Me.TxtPrimaryPayment.Text) > 0 Then
                    Me.TxtApprovalCode.Enabled = True
                    Me.TxtAccountNumber.Enabled = True
                    Me.SesuaikanHeight(True)
                Else
                    Me.TxtAccountNumber.Enabled = False
                    Me.TxtApprovalCode.Enabled = False
                    Me.SesuaikanHeight(False)
                End If
                'Hitung Ulang bank charge
                If ObjPayment.GetUserPrecentage(Me.TxtPrimaryPayment.Text) <> 0 Then
                    Me.ParentForm_1.Obj.BankCharge = ObjCD.DigitGrouping((Me.ParentForm_1.Obj.TaxAmount + Me.ParentForm_1.Obj.Subtotal) * ObjPayment.GetUserPrecentage(Me.TxtPrimaryPayment.Text) / 100, DefPrec, Me.ParentForm_1.Obj.Rounding)
                    vGrandTotalBelumRounding = Me.ParentForm_1.Obj.Subtotal + Me.ParentForm_1.Obj.TaxAmount + Me.ParentForm_1.Obj.BankCharge
                    vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
                    Me.ParentForm_1.Obj.GrandTotal = vGrandTotalRounding
                    Me.ParentForm_1.Obj.Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    If ObjPayment.GetBankChargeOption(Me.TxtPrimaryPayment.Text) = 1 Then
                        Me.TxtBankCharge.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.BankCharge, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Else
                        Me.TxtBankCharge.Text = 0
                    End If
                    Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.GrandTotal, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.TxtRounding.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.Rounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                Else
                    Me.ParentForm_1.Obj.BankCharge = ObjCD.DigitGrouping(ObjPayment.GetBankChargeNominal(Me.TxtPrimaryPayment.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    vGrandTotalBelumRounding = Me.ParentForm_1.Obj.Subtotal + Me.ParentForm_1.Obj.TaxAmount + Me.ParentForm_1.Obj.BankCharge
                    vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
                    Me.ParentForm_1.Obj.GrandTotal = vGrandTotalRounding
                    Me.ParentForm_1.Obj.Rounding = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.ParentForm_1.RoundingTipe)
                    If ObjPayment.GetBankChargeOption(Me.TxtPrimaryPayment.Text) = 1 Then
                        Me.TxtBankCharge.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.BankCharge, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Else
                        Me.TxtBankCharge.Text = 0
                    End If
                    Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.GrandTotal, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.TxtRounding.Text = ObjCD.DigitGrouping(Me.ParentForm_1.Obj.Rounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahDataPaymentBenar() As Boolean
        Dim RowDGV As DataGridViewRow
        Dim JmlWithdrawal As Integer = 0
        Dim ListPayWithdrawal As New List(Of Double)
        Dim ListPayBiasa As New List(Of Double)
        Dim PrimaryPayWD As Double = -1
        Dim PrimaryPayBiasa As Double = -1
        Dim TotalPayBiasa As Decimal = 0
        Dim TotalPayWD As Decimal = 0
        Try
            'Cek Payment tidak boleh kurang dari grand Total
            If Me.MultiPayment = False Then
                If CDec(Me.TxtTotalPayment.Text) < CDec(Me.TxtGrandTotal.Text) Then
                    MsgBox("Total payment is not enough", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                    If ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.TxtRemarks.Focus()
                        Return False
                    End If
                End If
                If String.IsNullOrEmpty(Me.TxtAccountNumber.Text) = False Then
                    If ObjCD.CekDataKosong(Me.TxtAccountNumber.Text, False) = False Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.TxtAccountNumber.Focus()
                        Return False
                    End If
                End If
                If String.IsNullOrEmpty(Me.TxtApprovalCode.Text) = False Then
                    If ObjCD.CekDataKosong(Me.TxtApprovalCode.Text, False) = False Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.TxtApprovalCode.Focus()
                        Return False
                    End If
                End If
            Else
                If CDec(Me.TxtTotalPayment2.Text) < CDec(Me.TxtGrandTotal.Text) Then
                    MsgBox("Total payment is not enough", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.MultiPayment = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    'cek ga boleh withdrawal semua 
                    For Each RowDGV In Me.dgvItem.Rows
                        If ObjPayment.GetTypePayment(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value) = 2 Then
                            JmlWithdrawal += 1
                        End If
                    Next
                    If Me.dgvItem.Rows.Count = JmlWithdrawal Then
                        MsgBox("you can not enter all data payments type cash withdrawal", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'Cek Jumlah pembayaran selain withdrawal mencukupi kaga
                    For Each RowDGV In Me.dgvItem.Rows
                        If RowDGV.Cells(Me.ColPrimaryPayment.Index).Value <> -1 Then
                            If ObjPayment.GetTipePayment(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value) = ClsMast_PaymentOption.EnumTipePayment.Cash_Withdrawal Then
                                If ListPayWithdrawal.Contains(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value) = False Then
                                    ListPayWithdrawal.Add(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value)
                                End If
                            Else
                                If ListPayBiasa.Contains(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value) = False Then
                                    ListPayBiasa.Add(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value)
                                End If
                            End If
                        Else
                            If ListPayBiasa.Contains(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value) = False Then
                                ListPayBiasa.Add(RowDGV.Cells(Me.ColPrimaryPayment.Index).Value)
                            End If
                        End If
                    Next
                    'Ambil Total payment withdrawal 
                    For Each PrimaryPayWD In ListPayWithdrawal
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Cells(Me.ColPrimaryPayment.Index).Value = PrimaryPayWD Then
                                TotalPayWD += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                                Exit For
                            End If
                        Next
                    Next
                    'Ambil Total payment biasa
                    For Each PrimaryPayBiasa In ListPayBiasa
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Cells(Me.ColPrimaryPayment.Index).Value = PrimaryPayBiasa Then
                                TotalPayBiasa += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                                Exit For
                            End If
                        Next
                    Next
                    'cek Total Payment biasa nya cukup untuk bayar transaksi nya ga
                    If CDec(TotalPayBiasa) - CDec(Me.TxtGrandTotal.Text) < 0 Then
                        MsgBox("Total payments other than withdrawal is not enough", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPayBiasa = Nothing
            ListPayWithdrawal = Nothing
        End Try
    End Function

    Private Function GetIndexItemInput(ByVal vPrimaryPayment As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryPayment.Index).Value = vPrimaryPayment Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColPaymentAmount.Index Or _
            e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Me.BufferGrandTotal = CDec(Me.TxtGrandTotal.Text) - CDec(Me.TxtBankCharge.Text)
            Else
                If ObjPayment.GetFillApprovalCode(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) > 0 Then
                    If e.ColumnIndex = Me.ColApprovalCode.Index Or _
                    e.ColumnIndex = Me.ColAccountNumber.Index Then
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim RowDGV As DataGridViewRow
        Dim Index As Integer = -1
        Dim vTotalPayment2 As Decimal = 0
        Dim vWithdrawal2 As Decimal = 0
        Dim vPaymentAmount2 As Decimal = 0
        Dim vBankCharge2 As Decimal = 0
        Dim ListString As New List(Of String)
        Dim CekListString As String = ""
        Dim ParamString As String = ""
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If ObjCD.CekDataKosong(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Data keterangan" _
                            + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColPaymentCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjPayment.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput(Me.ObjPayment.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value = Me.ObjPayment.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                If Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 0 Then
                                    Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Credit_Card_Payment"
                                ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 1 Then
                                    Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Debit_Card_Payment"
                                ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 2 Then
                                    Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash_Withdrawal"
                                ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 3 Then
                                    Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Voucher_Payment"
                                Else
                                    Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash"
                                End If
                                Me.dgvItem.Item(Me.ColPaymentBankCharge.Index, e.RowIndex).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetUserPrecentage(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) * (Me.ParentForm_1.Obj.Subtotal + Me.ParentForm_1.Obj.TaxAmount), DefPrec, Me.ParentForm_1.RoundingTipe)
                            Else
                                If e.RowIndex <> Index Then
                                    If Bahasa = EnumBahasa.Inggris Then
                                        MsgBox("You cannot input the same payment code", MsgBoxStyle.Information, "Information")
                                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                                        MsgBox("Anda tidak dapat menginput payment code yang sama", MsgBoxStyle.Information, "Informasi")
                                    End If
                                    Me.dgvItem.Item(Me.ColPaymentCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmMast_PaymentOption_Cari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.mPrimary)
                                If Index < 0 Then
                                    Me.dgvItem.Item(Me.ColPaymentCode.Index, e.RowIndex).Value = Me.ObjPayment.GetCode(ij.mPrimary)
                                    Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value = Me.ObjPayment.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                    If Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 0 Then
                                        Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Credit_Card_Payment"
                                    ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 1 Then
                                        Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Debit_Card_Payment"
                                    ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 2 Then
                                        Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash_Withdrawal"
                                    ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 3 Then
                                        Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Voucher_Payment"
                                    Else
                                        Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash"
                                    End If
                                    Me.dgvItem.Item(Me.ColPaymentBankCharge.Index, e.RowIndex).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetUserPrecentage(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) * (Me.ParentForm_1.Obj.Subtotal + Me.ParentForm_1.Obj.TaxAmount), DefPrec, Me.ParentForm_1.RoundingTipe)
                                    Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value = ij.mPrimary
                                Else
                                    If e.RowIndex <> Index Then
                                        If Bahasa = EnumBahasa.Inggris Then
                                            MsgBox("You cannot input the same payment code", MsgBoxStyle.Information, "Information")
                                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                                            MsgBox("Anda tidak dapat menginput payment code yang sama", MsgBoxStyle.Information, "Informasi")
                                        End If
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    Else
                                        Me.dgvItem.Item(Me.ColPaymentCode.Index, e.RowIndex).Value = Me.ObjPayment.GetCode(ij.mPrimary)
                                        Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value = Me.ObjPayment.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                        If Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 0 Then
                                            Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Credit_Card_Payment"
                                        ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 1 Then
                                            Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Debit_Card_Payment"
                                        ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 2 Then
                                            Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash_Withdrawal"
                                        ElseIf Me.ObjPayment.GetTypePayment(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) = 3 Then
                                            Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Voucher_Payment"
                                        Else
                                            Me.dgvItem.Item(Me.ColPaymentType.Index, e.RowIndex).Value = "Cash"
                                        End If
                                        Me.dgvItem.Item(Me.ColPaymentBankCharge.Index, e.RowIndex).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetUserPrecentage(Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value) * (Me.ParentForm_1.Obj.Subtotal + Me.ParentForm_1.Obj.TaxAmount), DefPrec, Me.ParentForm_1.RoundingTipe)
                                        Me.dgvItem.Item(Me.ColPrimaryPayment.Index, e.RowIndex).Value = ij.mPrimary
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvItem)
                    End If
                    'Hapus Baris Jika Code No DownGrade Kosong
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPaymentCode.Index, e.RowIndex).Value) = True Then
                            If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        End If
                    End If
                    '----------
                    'Cek Amount 
                    If e.ColumnIndex = Me.ColPaymentAmount.Index Then
                        If ObjCD.CekBilanganPositif(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please input positif number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.Buffer, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                            Exit Sub
                        End If
                        '----------
                        'Hitung Total Payment
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Index <> e.RowIndex Then
                                If RowDGV.Cells(Me.ColPaymentType.Index).Value <> "Cash_Withdrawal" Then
                                    vTotalPayment2 += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                                End If
                            End If
                        Next
                        If .Cells(Me.ColPaymentType.Index).Value <> "Cash_Withdrawal" Then
                            vPaymentAmount2 = .Cells(Me.ColPaymentAmount.Index).Value
                        Else
                            vPaymentAmount2 = 0
                        End If
                        Me.TxtTotalPayment2.Text = ObjCD.DigitGrouping(vTotalPayment2 + vPaymentAmount2, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'Hitung Withdrawal
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Index <> e.RowIndex Then
                                If RowDGV.Cells(Me.ColPaymentType.Index).Value = "Cash_Withdrawal" Then
                                    vWithdrawal2 += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                                End If
                            End If
                        Next
                        If .Cells(Me.ColPaymentType.Index).Value = "Cash_Withdrawal" Then
                            vPaymentAmount2 = .Cells(Me.ColPaymentAmount.Index).Value
                        Else
                            vPaymentAmount2 = 0
                        End If
                        Me.TxtWithdrawal.Text = ObjCD.DigitGrouping(vWithdrawal2 + vPaymentAmount2, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'Hitung BankCharge dan GrandTotal
                        If Me.ObjPayment.GetBankChargeOption(.Cells(Me.ColPrimaryPayment.Index).Value) = 1 Then
                            If Me.ObjPayment.GetUserPrecentage(.Cells(Me.ColPrimaryPayment.Index).Value) <> 0 Then
                                .Cells(Me.ColPaymentBankCharge.Index).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetUserPrecentage(.Cells(Me.ColPrimaryPayment.Index).Value) / 100 * (Me.Obj.Subtotal + Me.Obj.TaxAmount), DefPrec, Me.ParentForm_1.RoundingTipe)
                            Else
                                .Cells(Me.ColPaymentBankCharge.Index).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetBankChargeNominal(.Cells(Me.ColPrimaryPayment.Index).Value), DefPrec, Me.ParentForm_1.RoundingTipe)
                            End If
                        Else
                            .Cells(Me.ColPaymentBankCharge.Index).Value = 0
                        End If
                        For Each RowDGV In Me.dgvItem.Rows
                            vBankCharge2 += RowDGV.Cells(Me.ColPaymentBankCharge.Index).Value
                        Next
                        Me.TxtBankCharge.Text = ObjCD.DigitGrouping(CDec(vBankCharge2), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(CDec(Me.BufferGrandTotal) + CDec(Me.TxtBankCharge.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'Hitung Rounding
                        Dim vGrandTotalBelumRounding As Decimal = 0
                        Dim vGrandTotalRounding As Decimal =0
                        vGrandTotalBelumRounding = Me.TxtGrandTotal.Text
                        vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
                        Me.TxtRounding.Text = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'Hitung Change
                        Me.TxtChange2.Text = ObjCD.DigitGrouping(CDec(Me.TxtTotalPayment2.Text) - CDec(Me.TxtGrandTotal.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'update objek
                        Me.SalinKeObjek()
                        Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                    End If
                    'Fill Approval Code
                    If e.ColumnIndex = Me.ColAccountNumber.Index Then
                        If ObjPayment.GetFillApprovalCode(Row.Cells(Me.ColPrimaryPayment.Index).Value) = 1 Then
                            If ObjCD.CekDataKosong(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                                MsgBox("Data keterangan" _
                                + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = ""
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColApprovalCode.Index Then
                        If ObjPayment.GetFillApprovalCode(Row.Cells(Me.ColPrimaryPayment.Index).Value) = 1 Then
                            If ObjCD.CekDataKosong(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                                MsgBox("Data keterangan" _
                                + " tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = ""
                            Exit Sub
                        End If
                    End If
                    'Cek Kombinasi CodePayment + Approval Code + Account Code, pastikan tidak ada yang sama
                    ListString.Clear()
                    ParamString = ""
                    CekListString = ""
                    If e.ColumnIndex = Me.ColAccountNumber.Index Or e.ColumnIndex = Me.ColApprovalCode.Index Then
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Index <> e.RowIndex Then
                                If ListString.Contains(RowDGV.Cells(Me.ColPaymentCode.Index).Value + RowDGV.Cells(Me.ColApprovalCode.Index).Value + RowDGV.Cells(Me.ColAccountNumber.Index).Value) = False Then
                                    ListString.Add(RowDGV.Cells(Me.ColPaymentCode.Index).Value + RowDGV.Cells(Me.ColApprovalCode.Index).Value + RowDGV.Cells(Me.ColAccountNumber.Index).Value)
                                End If
                            End If
                        Next
                        ParamString = Me.dgvItem.Item(Me.ColPaymentCode.Index, e.RowIndex).Value.ToString + Me.dgvItem.Item(Me.ColApprovalCode.Index, e.RowIndex).Value.ToString + Me.dgvItem.Item(Me.ColAccountNumber.Index, e.RowIndex).Value.ToString
                        For Each CekListString In ListString
                            If CekListString = ParamString Then
                                MsgBox("Please input another payment code", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        Next
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColPaymentCode.Index Then
                            If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            End With
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
            Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim vCRV As New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Try
            If ObjSession.GetStatus(Me.ParentForm_1.TxtPrimarySession.Text) = ClsTrans_Session.EnumStatusSession.Closed Then
                MsgBox("Session ID : " + ObjSession.GetSessionID(Me.ParentForm_1.TxtPrimarySession.Text).ToString + " is closed." _
                , MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.mrOK = True
            If Me.mrOK = True Then
                If Me.ApakahDataPaymentBenar() = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'salin ke objek
                        Me.SalinKeObjek()
                        If Me.ParentForm_1.Obj.UpdateDataHeader(Me.ParentForm_1.UserCode) = True Then
                            If Me.MultiPayment = False Then
                                Me.ParentForm_1.ObjItem.SimpanItemPaymentSingle(Me.ParentForm_1.Obj.Primary, Me)
                            Else
                                Me.ParentForm_1.ObjItem.SimpanItemPaymentMultiple(Me.ParentForm_1.Obj.Primary, Me.dgvItem, Me)
                            End If
                            Me.HapusTBLNextYangTralhDisave(Me.ParentForm_1.TxtPrimary.Text, Me.ParentForm_1.TBLNext)
                            'Cetak Receipt
                            Me.ParentForm_1.Obj.CetakPOSUkuranStngahLetter(vCRV, Me.ParentForm_1.UserCode, Me.ParentForm_1.Obj.Primary, True)
                            '----
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.ParentForm_1.TxtQty.Text = "1"
                            Me.ParentForm_1.TxtPrimaryCodeNo.Text = "-1"
                            Me.ParentForm_1.TxtPrimaryMember.Text = "-1"
                            Me.ParentForm_1.TxtMember.Clear()
                            Me.ParentForm_1.TxtCodeNo.Clear()
                            Me.ParentForm_1.dtpTanggalTransaksi.Value = Now
                            Me.ParentForm_1.TxtPrimary.Text = "-1"
                            Me.ParentForm_1.Obj.ResetObject(Me.ParentForm_1.dgvItem)
                            Me.ParentForm_1.TxtIdNumber.Text = Me.ParentForm_1.Obj.GetNoBaru(Me.ParentForm_1.dtpTanggalTransaksi.Value, Me.ParentForm_1.TxtPrimarySession.Text)
                            Me.ParentForm_1.TxtDisplay.Clear()
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HapusTBLNextYangTralhDisave(ByVal vPrimaryPOS As Double, ByVal vTBLNext As DseTrans_POS.TBLNextDataTable)
        Dim Row As DseTrans_POS.TBLNextRow
        Dim i As Integer = 0
        Try
            If vTBLNext.Rows.Count > 0 Then
                While i <= vTBLNext.Rows.Count - 1
                    Row = vTBLNext.Rows(i)
                    If Row.PrimaryPOS = Me.ParentForm_1.TxtPrimary.Text Then
                        vTBLNext.Rows.RemoveAt(i)
                        Exit While
                    Else
                        i += 1
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            Me.ParentForm_1.Obj.DiscStr = Me.TxtDiscString.Text
            Me.ParentForm_1.Obj.Discount = Me.TxtDisc.Text
            Me.ParentForm_1.Obj.Subtotal = Me.TxtSubTotal.Text
            Me.ParentForm_1.Obj.Rounding = Me.TxtRounding.Text
            Me.ParentForm_1.Obj.BankCharge = Me.TxtBankCharge.Text
            Me.ParentForm_1.Obj.GrandTotal = Me.TxtGrandTotal.Text
            If Me.IsMultiPayment = False Then
                Me.ParentForm_1.Obj.TotalPayment = Me.TxtTotalPayment.Text
                Me.ParentForm_1.Obj.Change = Me.TxtChange.Text
            Else
                Me.ParentForm_1.Obj.TotalPayment = CDec(Me.TxtTotalPayment2.Text) + CDec(Me.TxtWithdrawal.Text)
                Me.ParentForm_1.Obj.Change = Me.TxtChange2.Text
            End If
            Me.ParentForm_1.Obj.IsVoid = 0
            Me.ParentForm_1.Obj.Remarks = Me.TxtRemarks.Text
            If Me.MultiPayment = True Then
                Me.ParentForm_1.Obj.IsSinglePayment = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItem.KeyDown
        Dim RowDGV As DataGridViewRow
        Dim TotalPayment As Decimal = 0
        Dim Withdrawal As Decimal = 0
        Try
            If e.KeyData = Keys.Delete Then
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    Me.TxtBankCharge.Text = ObjCD.DigitGrouping(CDec(Me.TxtBankCharge.Text) - CDec(Me.dgvItem.Item(Me.ColPaymentBankCharge.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(CDec(Me.TxtGrandTotal.Text) - CDec(Me.dgvItem.Item(Me.ColPaymentBankCharge.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    'Hitung Rounding 
                    Dim vGrandTotalBelumRounding As Decimal = 0
                    Dim vGrandTotalRounding As Decimal = 0
                    vGrandTotalBelumRounding = Me.TxtGrandTotal.Text
                    vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
                    Me.TxtRounding.Text = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                    '----------
                    Me.dgvItem.Rows.RemoveAt(Me.dgvItem.SelectedCells(0).RowIndex)
                    Me.ObjCD.SetNumber(Me.dgvItem, True)
                    'Hitung kembalian 
                    If Me.dgvItem.Rows.Count > 0 Then
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Cells(Me.ColPaymentType.Index).Value <> "Cash_Withdrawal" Then
                                TotalPayment += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                            End If
                        Next
                        Me.TxtTotalPayment2.Text = ObjCD.DigitGrouping(TotalPayment, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        For Each RowDGV In Me.dgvItem.Rows
                            If RowDGV.Cells(Me.ColPaymentType.Index).Value = "Cash_Withdrawal" Then
                                Withdrawal += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
                            End If
                        Next
                        Me.TxtWithdrawal.Text = ObjCD.DigitGrouping(Withdrawal, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        Me.TxtChange2.Text = ObjCD.DigitGrouping(CDec(Me.TxtTotalPayment2.Text) - CDec(Me.TxtGrandTotal.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
                        'update objek
                        Me.SalinKeObjek()
                        Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
                    Else
                        'Balikin data header ke posisi awal
                        Me.cboPaymentTipe.DataSource = Nothing
                        Me.cboPaymentTipe.Items.Clear()
                        Me.PrimaryPaymentOption = ObjGeneralSet.GetPrimaryPayment()
                        ObjPayment.IsiCBOPaymentOption(Me.cboPaymentTipe, Me.TBLPayment, DTAPayment, Me.PrimaryPaymentOption)
                        Me.cboPaymentTipe.SelectedIndex = 0
                        '-----
                        Me.TxtTotalPayment2.Text = 0
                        Me.TxtWithdrawal.Text = 0
                        Me.TxtChange2.Text = 0
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Dim RowDGV As DataGridViewRow
        Dim Total As Decimal = 0
        Try
            For Each RowDGV In Me.dgvItem.Rows
                Total += RowDGV.Cells(Me.ColPaymentAmount.Index).Value
            Next
            Me.TxtTotalPayment2.Text = ObjCD.DigitGrouping(Total, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            'Hitung Kembali Semua Pembayaran
            Me.ParentForm_1.SalinKeObjekUntukPayment()
            Me.HitungHeaderPayment(Me.ParentForm_1.Obj, Me.ParentForm_1)
            If Me.TxtTotalPayment2.Text = "0" Then
                Me.TxtChange2.Text = "0"
            Else
                Me.TxtChange2.Text = ObjCD.DigitGrouping(Total - Me.TxtGrandTotal.Text, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccountNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountNumber.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtAccountNumber.Enabled = True Then
                    If Me.TxtAccountNumber.ReadOnly = False Then
                        TxtAccountNumber.BackColor = Me.BackColorAll
                    Else
                        TxtAccountNumber.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccountNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccountNumber.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtAccountNumber.Enabled = True Then
                    If Me.TxtAccountNumber.ReadOnly = False Then
                        TxtAccountNumber.BackColor = Color.White
                    Else
                        TxtAccountNumber.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalPayment.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtTotalPayment.Enabled = True Then
                    If Me.TxtTotalPayment.ReadOnly = False Then
                        TxtTotalPayment.BackColor = Me.BackColorAll
                    Else
                        TxtTotalPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
                Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTotalPayment.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtTotalPayment.Enabled = True Then
                    If Me.TxtTotalPayment.ReadOnly = False Then
                        TxtTotalPayment.BackColor = Color.White
                    Else
                        TxtTotalPayment.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovalCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovalCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtApprovalCode.Enabled = True Then
                    If Me.TxtApprovalCode.ReadOnly = False Then
                        TxtApprovalCode.BackColor = Me.BackColorAll
                    Else
                        TxtApprovalCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtApprovalCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtApprovalCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtApprovalCode.Enabled = True Then
                    If Me.TxtApprovalCode.ReadOnly = False Then
                        TxtApprovalCode.BackColor = Color.White
                    Else
                        TxtApprovalCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Me.BackColorAll
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
                Me.Label3.Text = "Change: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtChange.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Color.White
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDiscString.Enabled = True Then
                    If Me.TxtDiscString.ReadOnly = False Then
                        Me.TxtDiscString.BackColor = Me.BackColorAll
                    Else
                        Me.TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDiscString.Enabled = True Then
                    If Me.TxtDiscString.ReadOnly = False Then
                        Me.TxtDiscString.BackColor = Color.White
                    Else
                        Me.TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.Enter
        Try
            Me.Buffer = Me.TxtDiscString.Text
            Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscString.KeyDown
        Try
            ObjCD.Disc_KeyDown(Me.TxtDiscString, sender, e)
            Me.TxtDisc.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscString.Text, Me.TxtTotalSales.Text, DefPrec, _
            Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscString.KeyUp
        Try
            ObjCD.Disc_KeyUp(Me.TxtDiscString, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.Leave
        Dim vGrandTotalBelumRounding As Decimal = 0
        Dim vGrandTotalRounding As Decimal = 0
        Try
            ObjCD.Disc_Leave(Me.TxtDiscString, Me.TxtDiscString.Text, sender, e, 2, ObjCD, Me.Buffer)
            Me.TxtDisc.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscString.Text, Me.TxtTotalSales.Text, DefPrec, _
            Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
            Me.TxtSubTotal.Text = ObjCD.DigitGrouping(Me.TxtTotalSales.Text - Me.TxtDisc.Text, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
            vGrandTotalBelumRounding = CDec(Me.TxtSubTotal.Text) + CDec(Me.TxtTaxAmount.Text) + CDec(Me.TxtBankCharge.Text)
            vGrandTotalRounding = ObjCD.GetRoundingValue(vGrandTotalBelumRounding, DefPrec, Me.ParentForm_1.RoundingTipe, Me.ParentForm_1.RoundingValue, Me.ObjCD)
            Me.TxtGrandTotal.Text = ObjCD.DigitGrouping(vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            Me.TxtRounding.Text = ObjCD.DigitGrouping(vGrandTotalBelumRounding - vGrandTotalRounding, Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            'Hitung lagi Change
            If Me.MultiPayment = False Then
                Me.TxtChange.Text = ObjCD.DigitGrouping(CDec(Me.TxtTotalPayment.Text) - CDec(Me.TxtGrandTotal.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            Else
                Me.TxtChange2.Text = ObjCD.DigitGrouping(CDec(Me.TxtTotalPayment2.Text) - CDec(Me.TxtGrandTotal.Text), Me.ParentForm_1.DefPrec, Me.ParentForm_1.RoundingTipe)
            End If
            Me.Label3.Font = Me.FontSubTotal
            Me.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_1.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class