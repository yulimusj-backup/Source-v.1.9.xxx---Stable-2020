Option Explicit On
Option Strict Off

Public Class FrmTransaksiJournal_Acquisition_Revaluation_Other_AOP

    Public Frm As New FrmTransaksiJournal_Acquisition_Revaluation_Other
    Public BackColorAll As Color

    Dim ObjSupp As New Core_Module.clsMasterSupplier
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjAop As New ClsTransaksiAOP
    Dim ObjProduuct As New Core_Module.clsMasterProduct
    Dim OBjCekData As New Core_Module.ClsCekData
    Dim ObjAkun As New ClsMasterAkun

    Private Sub btnFindSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim FrmCariSupp As New Core_Module.FrmCariMasterSupplier
        Try
            FrmCariSupp.txtCari.Text = Me.txtSupplierCode.Text
            FrmCariSupp.BackColorAll = Me.BackColorAll
            FrmCariSupp.ShowDialog()
            If FrmCariSupp.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = FrmCariSupp.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupp.GetSupplierCode(FrmCariSupp.PrimarySupplier)
                Me.ObjJurnal.CariDataAop(Me.txtPrimarySupplier.Text, Me.CboAOPNumber)
                Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
                Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
            Else
                Me.txtPrimarySupplier.Text = -1
                Me.txtSupplierCode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        Me.txtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtSupplierCode.Text <> "" Then
                    If Me.ObjSupp.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                        Me.txtPrimarySupplier.Text = Me.ObjSupp.GetPrimary(Me.txtSupplierCode.Text)
                        Me.ObjJurnal.CariDataAop(Me.txtPrimarySupplier.Text, Me.CboAOPNumber)
                        Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
                        Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
                    Else
                        Me.btnFindSupplier.PerformClick()
                    End If
                Else
                    Me.txtPrimarySupplier.Text = -1
                    Me.txtSupplierCode.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.ObjSupp.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.ObjSupp.GetPrimary(Me.txtSupplierCode.Text)
                    Me.ObjJurnal.CariDataAop(Me.txtPrimarySupplier.Text, Me.CboAOPNumber)
                    Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
                    Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            Else
                Me.txtPrimarySupplier.Text = -1
                Me.txtSupplierCode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAOPNumber_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAOPNumber.SelectedValueChanged
        Try
            Me.TxtPrimaryAOP.Text = Me.ObjAop.GetPrimary(Me.CboAOPNumber.Text)
            Me.ObjJurnal.CariDataProduct(Me.TxtPrimaryAOP.Text, Me.cboCodeNo)
            Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
            Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            Me.txtPrimaryCodeNo.Text = Me.ObjProduuct.GetPrimary(Me.cboCodeNo.Text)
            Me.ObjJurnal.CariDataPO(Me.TxtPrimaryAOP.Text, Me.txtPrimaryCodeNo.Text, Me.cboPurchaseOrder)
            Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
            Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
            Me.TxtPrimaryExpenseAccount.Text = Me.ObjJurnal.GetPrimaryExpense(Me.txtPrimaryCodeNo.Text)
            Me.TxtExpenseAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryExpenseAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPurchaseOrder_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPurchaseOrder.SelectedValueChanged
        Try
            Me.txtPrimaryPO.Text = Me.ObjJurnal.GetPrimaryPO(Me.TxtPrimaryAOP.Text, Me.txtPrimaryCodeNo.Text)
            Me.TxtAtPrice.Text = Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetAtPrice(Me.TxtPrimaryAOP.Text, _
            Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        Me.TxtQtyperBatch.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtQtyperBatch.Text <= Me.ObjJurnal.GetQty(Me.TxtPrimaryAOP.Text, _
            Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) Then
                Me.TxtTotal.Text = Me.OBjCekData.DigitGrouping(Me.TxtQtyperBatch.Text * Me.TxtAtPrice.Text, 2)
            Else
                MsgBox("Qty value should not be more than " + Me.OBjCekData.DigitGrouping(Me.ObjJurnal.GetQty(Me.TxtPrimaryAOP.Text, _
                Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), 2).ToString, MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatch.Text = 0
                Me.TxtQtyperBatch.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0
        Try
            For Each rowDGV In Frm.dgvPayment.Rows
                DefPrec = ObjCurr.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(rowDGV.Cells(Frm.COlItemPay_PrimaryAkun.Index).Value))
                If rowDGV.Cells(Frm.COlItemPay_PrimaryAkun.Index).Value = Me.TxtPrimaryExpenseAccount.Text Then
                    rowDGV.Cells(Frm.COlItemPay_AmountOriginal.Index).Value = Me.OBjCekData.DigitGrouping(CDec(rowDGV.Cells(Frm.COlItemPay_AmountOriginal.Index).Value) + _
                    CDec(Me.TxtTotal.Text), DefPrec)
                    Frm.TransferKeAccountDetail()
                    Me.Close()
                    Exit Sub
                End If
            Next
            DefPrec = ObjCurr.GetDefaultPrecission(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAccount.Text))
            Frm.dgvPayment.Rows.Add(Me.TxtExpenseAccount.Text, Me.ObjAkun.GetDescription(Me.TxtExpenseAccount.Text), _
            ObjCurr.GetCurrencyCode(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAccount.Text)), Me.OBjCekData.DigitGrouping(Me.TxtTotal.Text, DefPrec), _
            Me.OBjCekData.DigitGrouping(ObjCurr.GetDefaultExhangeRate(Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryExpenseAccount.Text)), DefPrec), "", "", "", _
            Me.TxtTotal.Text, -1, -1, Me.TxtExpenseAccount.Text)

            Frm.TransferKeAccountDetail()

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        Me.txtSupplierCode.BackColor = Me.txtPrimarySupplier.BackColor
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.LostFocus
        Me.TxtQtyperBatch.BackColor = Me.txtPrimarySupplier.BackColor
    End Sub
End Class