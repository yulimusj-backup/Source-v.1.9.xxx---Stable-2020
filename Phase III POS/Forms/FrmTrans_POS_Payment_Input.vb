Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS_Payment_Input
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public ParentForm_2 As FrmTrans_POS_Payment

    Dim ObjCD As New ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjGeneralSet As New ClsTrans_POS_GeneralSetting
    Dim ObjPayment As New ClsMast_PaymentOption
    Dim mrOK As Boolean = False

    Dim DefPrec As Integer
    Dim PrimaryPaymentOption As Double = -1
    Dim TBLPayment As New DseMast_Payment_Option.tblmaster_payment_option1_3DataTable
    Dim DTAPayment As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option1_3TableAdapter

    Private Sub FrmTrans_POS_Payment_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtTotalPayment.Text = "0"
            Me.TxtAccountNumber.Clear()
            Me.TxtApprovalCode.Clear()
            Me.TxtRemarks.Clear()
            '-----
            Me.PrimaryPaymentOption = ObjGeneralSet.GetPrimaryPayment()
            ObjPayment.IsiCBOPaymentOption(Me.cboPaymentTipe, Me.TBLPayment, DTAPayment, Me.PrimaryPaymentOption)
            If ObjPayment.GetFillApprovalCode(Me.PrimaryPaymentOption) > 0 Then
                Me.TxtAccountNumber.Enabled = True
                Me.TxtApprovalCode.Enabled = True
            Else
                Me.TxtAccountNumber.Enabled = False
                Me.TxtApprovalCode.Enabled = False
            End If
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

    Private Sub SalinKeObjek()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'If Me.TxtPrimaryPayment.Text = "-1" Then
            'MsgBox("Please input payment type", MsgBoxStyle.Information, "Information")
            'Me.cboPaymentTipe.Focus()
            'Return False
            'End If
            If ObjCD.CekBilanganPositif(Me.TxtTotalPayment.Text) = False Then
                MsgBox("Please enter positif number", MsgBoxStyle.Information, "Information")
                Me.TxtTotalPayment.Focus()
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexItem(ByVal vPrimaryPay As Double) As Double
        Dim RowDGV As DataGridViewRow
        Try
            For Each RowDGV In Me.ParentForm_2.dgvItem.Rows
                If RowDGV.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value = vPrimaryPay Then
                    Return RowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim index As Double = -1
        Dim rowDGV As DataGridViewRow
        Try
            If ApakahDataItemBenar() = True Then
                'cek apakah sudah ada item jenis pekerjaan sama yang masuk
                index = Me.GetIndexItem(Me.TxtPrimaryPayment.Text)
                If index < 0 Then
                    Me.ParentForm_2.dgvItem.Rows.Add()
                    rowDGV = Me.ParentForm_2.dgvItem.Rows(Me.ParentForm_2.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ParentForm_2.Col_No.Index).Value = ""
                        .Cells(Me.ParentForm_2.ColPaymentCode.Index).Value = Me.cboPaymentTipe.Text
                        .Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value = Me.TxtPrimaryPayment.Text
                        If Me.ObjPayment.GetTypePayment(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) = 0 Then
                            .Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Credit_Card_Payment"
                        ElseIf Me.ObjPayment.GetTypePayment(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) = 1 Then
                            .Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Debit_Card_Payment"
                        ElseIf Me.ObjPayment.GetTypePayment(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) = 2 Then
                            .Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Cash_Withdrawal"
                        ElseIf Me.ObjPayment.GetTypePayment(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) = 3 Then
                            .Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Voucher_Payment"
                        Else
                            .Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Cash"
                        End If
                        If Me.ObjPayment.GetBankChargeOption(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) = 1 Then
                            If Me.ObjPayment.GetUserPrecentage(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) <> 0 Then
                                .Cells(Me.ParentForm_2.ColPaymentBankCharge.Index).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetUserPrecentage(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value) / 100 * (Me.ParentForm_2.Obj.Subtotal + Me.ParentForm_2.Obj.TaxAmount), DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
                            Else
                                .Cells(Me.ParentForm_2.ColPaymentBankCharge.Index).Value = ObjCD.DigitGrouping(Me.ObjPayment.GetBankChargeNominal(.Cells(Me.ParentForm_2.ColPrimaryPayment.Index).Value), DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
                            End If
                        Else
                            .Cells(Me.ParentForm_2.ColPaymentBankCharge.Index).Value = 0
                        End If
                        .Cells(Me.ParentForm_2.ColPaymentAmount.Index).Value = ObjCD.DigitGrouping(Me.TxtTotalPayment.Text, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
                        .Cells(Me.ParentForm_2.ColAccountNumber.Index).Value = Me.TxtAccountNumber.Text
                        .Cells(Me.ParentForm_2.ColApprovalCode.Index).Value = Me.TxtApprovalCode.Text
                        .Cells(Me.ParentForm_2.ColRemarks.Index).Value = Me.TxtRemarks.Text
                    End With
                    'set number
                    Me.ObjCD.SetNumber(Me.ParentForm_2.dgvItem)
                    index = Me.ParentForm_2.dgvItem.Rows.Count - 1
                Else
                    MsgBox("You cannot input this data, because data has already used in multiple payment", MsgBoxStyle.Information, "Information")
                    Me.cboPaymentTipe.Focus()
                    Exit Sub
                End If

                'salin ke Objek yang berkaitan hitung bank charge
                Me.SalinKeObjekNHitungUlang()
                Me.SalinKeField()
                'Kembali ulangi proses salin untuk mengupdate nilai withdrawal yang baru di input 
                Me.SalinKeObjekNHitungUlang()
                Me.SalinKeField()
                Me.ParentForm_2.Label3.Text = "Total: " + ObjCur.GetCurrencyCode(ObjGeneralSet.GetPrimaryCurr()) + ". " + ObjCD.DigitGrouping(Me.ParentForm_2.TxtGrandTotal.Text.ToString, Me.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)

                '-----Scrolling DGV-----
                If Me.ParentForm_2.dgvItem.FirstDisplayedScrollingRowIndex < index Then
                    While Me.ParentForm_2.dgvItem.Rows(index).Displayed = False _
                    Or Me.ParentForm_2.dgvItem.DisplayedRowCount(True) <> Me.ParentForm_2.dgvItem.DisplayedRowCount(False)
                        Me.ParentForm_2.dgvItem.FirstDisplayedScrollingRowIndex = Me.ParentForm_2.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                Else
                    While Me.ParentForm_2.dgvItem.Rows(index).Displayed = False
                        Me.ParentForm_2.dgvItem.FirstDisplayedScrollingRowIndex = _
                        Me.ParentForm_2.dgvItem.FirstDisplayedScrollingRowIndex - 1
                    End While
                End If

                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HitungTotalPaymentDGV() As Decimal
        Dim RowDGV As DataGridViewRow
        Dim Retval As Decimal = 0
        Try
            For Each RowDGV In Me.ParentForm_2.dgvItem.Rows
                If RowDGV.Cells(Me.ParentForm_2.ColPaymentType.Index).Value <> "Cash_Withdrawal" Then
                    Retval += RowDGV.Cells(Me.ParentForm_2.ColPaymentAmount.Index).Value
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function HitungTotalWithdrawalDGV() As Decimal
        Dim RowDGV As DataGridViewRow
        Dim Retval As Decimal = 0
        Try
            For Each RowDGV In Me.ParentForm_2.dgvItem.Rows
                If RowDGV.Cells(Me.ParentForm_2.ColPaymentType.Index).Value = "Cash_Withdrawal" Then
                    Retval += RowDGV.Cells(Me.ParentForm_2.ColPaymentAmount.Index).Value
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function HitungBankChargeDGV() As Decimal
        Dim RowDGV As DataGridViewRow
        Dim Retval As Decimal = 0
        Try
            For Each RowDGV In Me.ParentForm_2.dgvItem.Rows
                Retval += RowDGV.Cells(Me.ParentForm_2.ColPaymentBankCharge.Index).Value
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjekNHitungUlang()
        Dim vGrandTotalRounding As Decimal = 0
        Dim vTotalPaymentMinusWithdrawal As Decimal = 0
        Try
            Me.ParentForm_2.Obj.BankCharge = Me.HitungBankChargeDGV()
            Me.ParentForm_2.Obj.GrandTotal = ObjCD.DigitGrouping(CDec(Me.ParentForm_2.TxtSubTotal.Text) + CDec(Me.ParentForm_2.TxtTaxAmount.Text) + Me.ParentForm_2.Obj.BankCharge, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            vGrandTotalRounding = ObjCD.GetRoundingValue(Me.ParentForm_2.Obj.GrandTotal, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe, Me.ParentForm_2.ParentForm_1.RoundingValue, Me.ObjCD)
            Me.ParentForm_2.Obj.GrandTotal = vGrandTotalRounding
            Me.ParentForm_2.Obj.Rounding = ObjCD.DigitGrouping(vGrandTotalRounding - Me.ParentForm_2.Obj.GrandTotal, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.Obj.TotalPayment = ObjCD.DigitGrouping(CDec(Me.HitungTotalPaymentDGV) + CDec(Me.ParentForm_2.TxtWithdrawal.Text), Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            vTotalPaymentMinusWithdrawal = CDec(Me.ParentForm_2.Obj.TotalPayment) - CDec(Me.ParentForm_2.TxtWithdrawal.Text)
            If vTotalPaymentMinusWithdrawal <= 0 Then
                vTotalPaymentMinusWithdrawal = 0
            End If
            Me.ParentForm_2.Obj.Change = ObjCD.DigitGrouping(vTotalPaymentMinusWithdrawal - vGrandTotalRounding, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.Obj.IsVoid = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            Me.ParentForm_2.TxtBankCharge.Text = ObjCD.DigitGrouping(Me.ParentForm_2.Obj.BankCharge, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.TxtGrandTotal.Text = ObjCD.DigitGrouping(Me.ParentForm_2.Obj.GrandTotal, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.TxtRounding.Text = ObjCD.DigitGrouping(Me.ParentForm_2.Obj.Rounding, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.TxtTotalPayment2.Text = ObjCD.DigitGrouping(CDec(Me.ParentForm_2.Obj.TotalPayment) - CDec(Me.HitungTotalWithdrawalDGV), Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            If Me.ParentForm_2.TxtTotalPayment2.Text <= 0 Then
                Me.ParentForm_2.TxtTotalPayment2.Text = 0
            End If
            Me.ParentForm_2.TxtWithdrawal.Text = ObjCD.DigitGrouping(CDec(Me.HitungTotalWithdrawalDGV), Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
            Me.ParentForm_2.TxtChange2.Text = ObjCD.DigitGrouping(Me.ParentForm_2.Obj.Change, Me.ParentForm_2.ParentForm_1.DefPrec, Me.ParentForm_2.ParentForm_1.RoundingTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalPayment.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtTotalPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTotalPayment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTotalPayment.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtTotalPayment, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPaymentTipe_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPaymentTipe.SelectedIndexChanged
        Dim row As DseMast_Payment_Option.tblmaster_payment_option1_3Row
        Try
            If Me.cboPaymentTipe.SelectedIndex = -1 Then
                Me.TxtPrimaryPayment.Text = "-1"
                Me.TxtAccountNumber.Enabled = False
                Me.TxtApprovalCode.Enabled = False
            Else
                row = Me.TBLPayment.DefaultView(Me.cboPaymentTipe.SelectedIndex).Row
                Me.TxtPrimaryPayment.Text = row.primary
                If ObjPayment.GetFillApprovalCode(Me.TxtPrimaryPayment.Text) > 0 Then
                    Me.TxtAccountNumber.Enabled = True
                    Me.TxtApprovalCode.Enabled = True
                Else
                    Me.TxtAccountNumber.Enabled = False
                    Me.TxtApprovalCode.Enabled = False
                End If
            End If
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
End Class