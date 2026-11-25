Option Explicit On
Option Strict Off
Imports Core_Module, AccModule

Public Class FrmMast_PaymentOption
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim baru As Boolean
    Dim buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim Obj As New ClsMast_PaymentOption
    Dim ObjAkun As New ClsMasterAkun
    Dim Def_Prec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency_POS
    Dim TandaTab2Aktif As Boolean = False

    Private Sub CboChargeRate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboChargeRate.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.CboChargeRate.SelectedIndex = 0 Then
                    'fixed nominal rate per transaction
                    Me.CC_PercentageChargeLabel.Visible = False
                    Me.CC_PercentageChargeLabel2.Visible = False
                    Me.CC_PercentageLabel3.Visible = False
                    Me.CC_PercentageLabel4.Visible = False
                    Me.TxtBankChargePrecentage.Visible = False
                    Me.TxtBankChargePrecentage.TabStop = False
                    Me.TxtBankChargePrecentage.Enabled = False
                    Me.TxtBankCustChargePercentage.Visible = False
                    Me.TxtBankCustChargePercentage.TabStop = False
                    Me.TxtBankCustChargePercentage.Enabled = False
                    Me.ChkConsistentCharge.Enabled = False
                    Me.ChkConsistentCharge.Checked = True
                    Me.CC_NominalChargeLabel.Visible = True
                    Me.TxtBankChargeNominal.TabStop = True
                    Me.TxtBankChargeNominal.Visible = True
                    Me.TxtBankChargeNominal.Enabled = True
                    Me.TxtBankChargeNominal.Text = "0"
                    Me.TxtBankChargePrecentage.Text = "0.00"
                    Me.TxtBankCustChargePercentage.Text = "0.00"
                Else
                    'percetage of the total transaction
                    Me.CC_PercentageChargeLabel.Visible = True
                    Me.CC_PercentageChargeLabel2.Visible = True
                    Me.CC_PercentageLabel3.Visible = True
                    Me.CC_PercentageLabel4.Visible = True
                    Me.TxtBankChargePrecentage.Visible = True
                    Me.TxtBankChargePrecentage.TabStop = True
                    Me.TxtBankCustChargePercentage.Visible = True
                    Me.TxtBankCustChargePercentage.TabStop = True
                    If Me.ChkConsistentCharge.Checked = True Then
                        Me.TxtBankCustChargePercentage.Enabled = False
                        Me.TxtBankCustChargePercentage.TabStop = False
                        Me.TxtBankChargePrecentage.Enabled = True
                        Me.TxtBankChargePrecentage.TabStop = True
                    Else
                        Me.TxtBankCustChargePercentage.Enabled = True
                        Me.TxtBankCustChargePercentage.TabStop = True
                        Me.TxtBankChargePrecentage.Enabled = True
                        Me.TxtBankChargePrecentage.TabStop = True
                    End If
                    Me.ChkConsistentCharge.Enabled = True
                    Me.CC_NominalChargeLabel.Visible = False
                    Me.TxtBankChargeNominal.TabStop = False
                    Me.TxtBankChargeNominal.Visible = False
                    Me.TxtBankChargeNominal.Enabled = False
                    Me.TxtBankChargeNominal.Text = "0"
                    Me.TxtBankChargePrecentage.Text = "0.00"
                    Me.TxtBankCustChargePercentage.Text = "0.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_PaymentOption_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboBankPaymentJournal.SelectedIndex = 0
            Me.CboChargeRate.SelectedIndex = 1
            Me.CboBankChargeOption.SelectedIndex = 0
            Me.cboPaymentType.SelectedIndex = 0
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency_POS
            Me.Obj = New ClsMast_PaymentOption
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.TandaTab2Aktif = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CC_CboChargeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboBankChargeOption.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.CboBankChargeOption.SelectedIndex = 0 Then
                    Me.TxtBankCustChargePercentage.Text = "0.00"
                    Me.TxtBankChargePrecentage.Text = "0.00"
                    Me.TxtBankChargeNominal.Text = Me.ObjCD.DigitGrouping("0", Me.Def_Prec)
                    Me.CboChargeRate.SelectedIndex = 1

                    Me.TxtBankCustChargePercentage.TabStop = False
                    Me.TxtBankCustChargePercentage.Enabled = False
                    Me.TxtBankChargePrecentage.TabStop = False
                    Me.TxtBankChargePrecentage.Enabled = False
                    Me.TxtBankChargeNominal.TabStop = False
                    Me.TxtBankChargeNominal.Enabled = False
                    Me.CboChargeRate.Enabled = False
                    Me.CboChargeRate.TabStop = False
                    Me.ChkConsistentCharge.Enabled = False
                    Me.ChkConsistentCharge.TabStop = False
                    Me.ChkConsistentCharge.Checked = True
                Else
                    Me.TxtBankChargePrecentage.TabStop = True
                    Me.TxtBankChargePrecentage.Enabled = True
                    Me.CboChargeRate.Enabled = True
                    Me.CboChargeRate.TabStop = True
                    Me.ChkConsistentCharge.Enabled = True
                    Me.ChkConsistentCharge.TabStop = True
                    Me.ChkConsistentCharge.Checked = True
                    If Me.CboChargeRate.SelectedIndex = 0 Then
                        Me.TxtBankChargePrecentage.TabStop = False
                        Me.TxtBankChargePrecentage.Enabled = False
                        Me.TxtBankCustChargePercentage.TabStop = False
                        Me.TxtBankCustChargePercentage.Enabled = False
                        Me.TxtBankChargeNominal.TabStop = True
                        Me.TxtBankChargeNominal.Enabled = True
                    Else
                        If Me.ChkConsistentCharge.Checked = True Then
                            Me.TxtBankCustChargePercentage.TabStop = False
                            Me.TxtBankCustChargePercentage.Enabled = False
                        Else
                            Me.TxtBankCustChargePercentage.TabStop = True
                            Me.TxtBankCustChargePercentage.Enabled = True
                        End If
                        Me.TxtBankChargeNominal.TabStop = False
                        Me.TxtBankChargeNominal.Enabled = False
                    End If
                    Me.TxtBankCustChargePercentage.Text = "0.00"
                    Me.TxtBankChargePrecentage.Text = "0.00"
                    Me.TxtBankChargeNominal.Text = Me.ObjCD.DigitGrouping("0", Me.Def_Prec)
                    Me.CboChargeRate.SelectedIndex = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CC_ChkConsistentCharge_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkConsistentCharge.CheckedChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.ChkConsistentCharge.Checked = True Then
                    Me.TxtBankCustChargePercentage.Enabled = False
                    Me.TxtBankCustChargePercentage.TabStop = False
                    Me.TxtBankCustChargePercentage.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtBankCustChargePercentage.Text = Me.TxtBankChargePrecentage.Text
                    Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(Me.TxtBankChargePrecentage.Text, 2)
                Else
                    Me.TxtBankCustChargePercentage.Enabled = True
                    Me.TxtBankCustChargePercentage.TabStop = True
                    Me.TxtBankCustChargePercentage.BackColor = Me.TxtPrimary.BackColor
                    Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(0, 2)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPaymentType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPaymentType.SelectedIndexChanged
        Try
            If Me.cboPaymentType.SelectedIndex = 0 Then
                'Credit Card Payment
                Me.TabControl1.BringToFront()
                Me.TabControl1.SelectedIndex = 0
                Me.TandaTab2Aktif = False
                Me.KosongkanTab(2)

                Me.Txt_Acc_CostPayable.Visible = True
                Me.Txt_Acc_CostPayable.TabStop = True
                Me.btnFind_Acc_CostPayable.Visible = True
                Me.Lbl_CC_Cost.Visible = True
                Me.Lbl_CC_Cost_Payable.Visible = True

                Me.Txt_Acc_CreditCardCost.Visible = True
                Me.Txt_Acc_CreditCardCost.TabStop = True
                Me.btnFind_Acc_CreditCardCost.Visible = True
            ElseIf Me.cboPaymentType.SelectedIndex = 1 Then
                'Debit Card Payment
                Me.TabControl1.BringToFront()
                Me.TabControl1.SelectedIndex = 0
                Me.TandaTab2Aktif = False
                Me.KosongkanTab(2)

                Me.Txt_Acc_CostPayable.Visible = True
                Me.Txt_Acc_CostPayable.TabStop = True
                Me.btnFind_Acc_CostPayable.Visible = True

                Me.Txt_Acc_CreditCardCost.Visible = True
                Me.Txt_Acc_CreditCardCost.TabStop = True
                Me.btnFind_Acc_CreditCardCost.Visible = True

                Me.Lbl_CC_Cost.Visible = True
                Me.Lbl_CC_Cost_Payable.Visible = True
            ElseIf Me.cboPaymentType.SelectedIndex = 2 Then
                'Cash Withdrawal
                Me.TabControl1.BringToFront()
                Me.TabControl1.SelectedIndex = 0
                Me.TandaTab2Aktif = False
                Me.KosongkanTab(2)

                Me.Txt_Acc_CostPayable.Visible = False
                Me.Txt_Acc_CostPayable.TabStop = False
                Me.btnFind_Acc_CostPayable.Visible = False
                Me.btnFind_Acc_CostPayable.TabStop = False
                Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
                Me.Txt_Acc_CostPayable.Text = ""

                Me.Txt_Acc_CreditCardCost.Visible = False
                Me.Txt_Acc_CreditCardCost.TabStop = False
                Me.btnFind_Acc_CreditCardCost.Visible = False
                Me.btnFind_Acc_CreditCardCost.TabStop = False
                Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
                Me.Txt_Acc_CreditCardCost.Text = ""

                Me.Lbl_CC_Cost.Visible = False
                Me.Lbl_CC_Cost_Payable.Visible = False
            ElseIf Me.cboPaymentType.SelectedIndex = 3 Then
                'Voucher Payment
                Me.TabControl2.BringToFront()
                Me.TabControl2.SelectedIndex = 0
                Me.TandaTab2Aktif = True
                Me.KosongkanTab(1)
            End If

            If Me.TandaTab2Aktif = False Then
                Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.tbpDetails.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False)
            Else
                Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.tbpDetails.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_BankCash_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Txt_Acc_BankCash.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.Txt_Acc_BankCash.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.Txt_Acc_BankCash.Text) = True Then
                        Me.Txt_Acc_Primary_BankCash.Text = objCode.GetPrimary(Me.Txt_Acc_BankCash.Text)
                        Me.Txt_Acc_BankCash.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_BankCash.Text)
                        Me.Txt_Acc_CostPayable.Focus()
                    Else
                        Me.btnFind_Acc_BankCash.PerformClick()
                    End If
                Else
                    Me.btnFind_Acc_BankCash.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_BankCash_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Txt_Acc_BankCash.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_BankCash.Text <> "" Then
                If objCode.ApakahSudahAda(Me.Txt_Acc_BankCash.Text) = True Then
                    Me.Txt_Acc_Primary_BankCash.Text = objCode.GetPrimary(Me.Txt_Acc_BankCash.Text)
                    Me.Txt_Acc_BankCash.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_BankCash.Text)
                Else
                    If Me.Txt_Acc_BankCash.Text <> "" Then
                        i.txtCari.Text = Me.Txt_Acc_BankCash.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.Txt_Acc_Primary_BankCash.Text = i.mPrimaryAkun
                        Me.Txt_Acc_BankCash.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_BankCash.Text)
                    Else
                        Me.Txt_Acc_BankCash.Clear()
                        Me.Txt_Acc_Primary_BankCash.Text = "-1"
                        Me.Txt_Acc_BankCash.Focus()
                    End If
                End If
            Else
                Me.Txt_Acc_BankCash.Clear()
                Me.Txt_Acc_Primary_BankCash.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFind_Acc_BankCash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFind_Acc_BankCash.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_BankCash.Text <> "" Then
                i.txtCari.Text = Me.Txt_Acc_BankCash.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.Txt_Acc_Primary_BankCash.Text = i.mPrimaryAkun
                Me.Txt_Acc_BankCash.Text = ObjCode.GetAccountCode(Me.Txt_Acc_Primary_BankCash.Text)
                Me.Txt_Acc_CostPayable.Focus()
            Else
                Me.Txt_Acc_BankCash.Clear()
                Me.Txt_Acc_Primary_BankCash.Text = "-1"
                Me.Txt_Acc_BankCash.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_BankReceivables_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Txt_Acc_BankReceivables.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.Txt_Acc_BankReceivables.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.Txt_Acc_BankReceivables.Text) = True Then
                        Me.Txt_Acc_Primary_Bank_Receivable.Text = objCode.GetPrimary(Me.Txt_Acc_BankReceivables.Text)
                        Me.Txt_Acc_BankReceivables.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                        Me.Txt_Acc_BankCash.Focus()
                    Else
                        Me.btnFind_Acc_BankReceivables.PerformClick()
                    End If
                Else
                    Me.btnFind_Acc_BankReceivables.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_BankReceivables_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Txt_Acc_BankReceivables.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_BankReceivables.Text <> "" Then
                If objCode.ApakahSudahAda(Me.Txt_Acc_BankReceivables.Text) = True Then
                    Me.Txt_Acc_Primary_Bank_Receivable.Text = objCode.GetPrimary(Me.Txt_Acc_BankReceivables.Text)
                    Me.Txt_Acc_BankReceivables.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                Else
                    If Me.Txt_Acc_BankReceivables.Text <> "" Then
                        i.txtCari.Text = Me.Txt_Acc_BankReceivables.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.Txt_Acc_Primary_Bank_Receivable.Text = i.mPrimaryAkun
                        Me.Txt_Acc_BankReceivables.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                    Else
                        Me.Txt_Acc_BankReceivables.Clear()
                        Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1"
                        Me.Txt_Acc_BankReceivables.Focus()
                    End If
                End If
            Else
                Me.Txt_Acc_BankReceivables.Clear()
                Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFind_Acc_BankReceivables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFind_Acc_BankReceivables.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_BankReceivables.Text <> "" Then
                i.txtCari.Text = Me.Txt_Acc_BankReceivables.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.Txt_Acc_Primary_Bank_Receivable.Text = i.mPrimaryAkun
                Me.Txt_Acc_BankReceivables.Text = ObjCode.GetAccountCode(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                Me.Txt_Acc_BankCash.Focus()
            Else
                Me.Txt_Acc_BankReceivables.Clear()
                Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1"
                Me.Txt_Acc_BankReceivables.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_CostPayable_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Txt_Acc_CostPayable.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.Txt_Acc_CostPayable.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.Txt_Acc_CostPayable.Text) = True Then
                        Me.Txt_Acc_PrimaryCostPayable.Text = objCode.GetPrimary(Me.Txt_Acc_CostPayable.Text)
                        Me.Txt_Acc_CostPayable.Text = objCode.GetAccountCode(Me.Txt_Acc_PrimaryCostPayable.Text)
                        Me.Txt_Acc_CreditCardCost.Focus()
                    Else
                        Me.btnFind_Acc_CostPayable.PerformClick()
                    End If
                Else
                    Me.btnFind_Acc_CostPayable.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_CostPayable_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Txt_Acc_CostPayable.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_CostPayable.Text <> "" Then
                If objCode.ApakahSudahAda(Me.Txt_Acc_CostPayable.Text) = True Then
                    Me.Txt_Acc_PrimaryCostPayable.Text = objCode.GetPrimary(Me.Txt_Acc_CostPayable.Text)
                    Me.Txt_Acc_CostPayable.Text = objCode.GetAccountCode(Me.Txt_Acc_PrimaryCostPayable.Text)
                Else
                    If Me.Txt_Acc_CostPayable.Text <> "" Then
                        i.txtCari.Text = Me.Txt_Acc_CostPayable.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.Txt_Acc_PrimaryCostPayable.Text = i.mPrimaryAkun
                        Me.Txt_Acc_CostPayable.Text = objCode.GetAccountCode(Me.Txt_Acc_PrimaryCostPayable.Text)
                    Else
                        Me.Txt_Acc_CostPayable.Clear()
                        Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
                        Me.Txt_Acc_CostPayable.Focus()
                    End If
                End If
            Else
                Me.Txt_Acc_CostPayable.Clear()
                Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFind_Acc_CostPayable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFind_Acc_CostPayable.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_CostPayable.Text <> "" Then
                i.txtCari.Text = Me.Txt_Acc_CostPayable.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.Txt_Acc_PrimaryCostPayable.Text = i.mPrimaryAkun
                Me.Txt_Acc_CostPayable.Text = ObjCode.GetAccountCode(Me.Txt_Acc_PrimaryCostPayable.Text)
                Me.Txt_Acc_CreditCardCost.Focus()
            Else
                Me.Txt_Acc_CostPayable.Clear()
                Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
                Me.Txt_Acc_CostPayable.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_CreditCardCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Txt_Acc_CreditCardCost.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.Txt_Acc_CreditCardCost.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.Txt_Acc_CreditCardCost.Text) = True Then
                        Me.Txt_Acc_Primary_CreditCardCost.Text = objCode.GetPrimary(Me.Txt_Acc_CreditCardCost.Text)
                        Me.Txt_Acc_CreditCardCost.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_CreditCardCost.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.btnFind_Acc_CreditCardCost.PerformClick()
                    End If
                Else
                    Me.btnFind_Acc_CreditCardCost.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Acc_CreditCardCost_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Txt_Acc_CreditCardCost.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_CreditCardCost.Text <> "" Then
                If objCode.ApakahSudahAda(Me.Txt_Acc_CreditCardCost.Text) = True Then
                    Me.Txt_Acc_Primary_CreditCardCost.Text = objCode.GetPrimary(Me.Txt_Acc_CreditCardCost.Text)
                    Me.Txt_Acc_CreditCardCost.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_CreditCardCost.Text)
                Else
                    If Me.Txt_Acc_CreditCardCost.Text <> "" Then
                        i.txtCari.Text = Me.Txt_Acc_CreditCardCost.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.Txt_Acc_Primary_CreditCardCost.Text = i.mPrimaryAkun
                        Me.Txt_Acc_CreditCardCost.Text = objCode.GetAccountCode(Me.Txt_Acc_Primary_CreditCardCost.Text)
                    Else
                        Me.Txt_Acc_CreditCardCost.Clear()
                        Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
                        Me.Txt_Acc_CreditCardCost.Focus()
                    End If
                End If
            Else
                Me.Txt_Acc_CreditCardCost.Clear()
                Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFind_Acc_CreditCardCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnFind_Acc_CreditCardCost.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Acc_CreditCardCost.Text <> "" Then
                i.txtCari.Text = Me.Txt_Acc_CreditCardCost.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.Txt_Acc_Primary_CreditCardCost.Text = i.mPrimaryAkun
                Me.Txt_Acc_CreditCardCost.Text = ObjCode.GetAccountCode(Me.Txt_Acc_Primary_CreditCardCost.Text)
                Me.btnOK.Focus()
            Else
                Me.Txt_Acc_CreditCardCost.Clear()
                Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
                Me.Txt_Acc_CreditCardCost.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Vou_CostAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Txt_Vou_CostAccount.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.Txt_Vou_CostAccount.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.Txt_Vou_CostAccount.Text) = True Then
                        Me.Txt_Vou_PrimaryCostAccount.Text = objCode.GetPrimary(Me.Txt_Vou_CostAccount.Text)
                        Me.Txt_Vou_CostAccount.Text = objCode.GetAccountCode(Me.Txt_Vou_PrimaryCostAccount.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.Btn_Vou_FindCostAccount.PerformClick()
                    End If
                Else
                    Me.Btn_Vou_FindCostAccount.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub Txt_Vou_CostAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Txt_Vou_CostAccount.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Vou_CostAccount.Text <> "" Then
                If objCode.ApakahSudahAda(Me.Txt_Vou_CostAccount.Text) = True Then
                    Me.Txt_Vou_PrimaryCostAccount.Text = objCode.GetPrimary(Me.Txt_Vou_CostAccount.Text)
                    Me.Txt_Vou_CostAccount.Text = objCode.GetAccountCode(Me.Txt_Vou_PrimaryCostAccount.Text)
                Else
                    If Me.Txt_Vou_CostAccount.Text <> "" Then
                        i.txtCari.Text = Me.Txt_Vou_CostAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.Txt_Vou_PrimaryCostAccount.Text = i.mPrimaryAkun
                        Me.Txt_Vou_CostAccount.Text = objCode.GetAccountCode(Me.Txt_Vou_PrimaryCostAccount.Text)
                    Else
                        Me.Txt_Vou_CostAccount.Clear()
                        Me.Txt_Vou_PrimaryCostAccount.Text = "-1"
                        Me.Txt_Vou_CostAccount.Focus()
                    End If
                End If
            Else
                Me.Txt_Vou_CostAccount.Clear()
                Me.Txt_Vou_PrimaryCostAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub Btn_Vou_FindCostAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles Btn_Vou_FindCostAccount.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.Txt_Vou_CostAccount.Text <> "" Then
                i.txtCari.Text = Me.Txt_Vou_CostAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.Txt_Vou_PrimaryCostAccount.Text = i.mPrimaryAkun
                Me.Txt_Vou_CostAccount.Text = ObjCode.GetAccountCode(Me.Txt_Vou_PrimaryCostAccount.Text)
                Me.btnOK.Focus()
            Else
                Me.Txt_Vou_CostAccount.Clear()
                Me.Txt_Vou_PrimaryCostAccount.Text = "-1"
                Me.Txt_Vou_CostAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            If Me.TandaTab2Aktif = False Then
                Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.tbpDetails.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False)
            Else
                Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.tbpDetails.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True)
                Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True)
            End If
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.ChkConsistentCharge.Enabled = True
            If Me.ChkConsistentCharge.Checked = True Then
                Me.TxtBankCustChargePercentage.Enabled = False
                Me.TxtBankCustChargePercentage.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtBankCustChargePercentage.TabStop = False
            Else
                Me.TxtBankCustChargePercentage.Enabled = True
                Me.TxtBankCustChargePercentage.BackColor = Me.TxtPrimary.BackColor
                Me.TxtBankCustChargePercentage.TabStop = True
            End If
            If Me.CboBankChargeOption.SelectedIndex = 0 Then
                Me.ChkConsistentCharge.Enabled = False
                Me.CboChargeRate.Enabled = False
                Me.TxtBankCustChargePercentage.Enabled = False
                Me.TxtBankChargePrecentage.Enabled = False
            End If
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()
            Me.ChkFillApproval.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabControl1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabControl2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.tbpDetails.Controls, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.ChkConsistentCharge.Enabled = False
            Me.BackColorDefault()
            If Me.CboChargeRate.SelectedIndex = 0 Then
                Me.TxtBankChargeNominal.Visible = True
                Me.TxtBankChargeNominal.TabStop = True
                Me.TxtBankChargePrecentage.Visible = False
                Me.TxtBankCustChargePercentage.TabStop = False
                Me.TxtBankCustChargePercentage.Visible = False
                Me.TxtBankCustChargePercentage.TabStop = False
            Else
                Me.TxtBankChargeNominal.Visible = False
                Me.TxtBankChargeNominal.TabStop = False
                Me.TxtBankChargePrecentage.Visible = True
                Me.TxtBankCustChargePercentage.TabStop = True
                Me.TxtBankCustChargePercentage.Visible = True
                Me.TxtBankCustChargePercentage.TabStop = True
            End If
            Me.ChkFillApproval.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabControl2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage3.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.tbpDetails.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabControl1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabControl2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage3.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.tbpDetails.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanTab(ByVal TabYgMana As Integer)
        Try
            If TabYgMana = 1 Then
                Me.ObjCD.KosongkanControl(Me.TabControl1.Controls, True, True, True)
                Me.ObjCD.KosongkanControl(Me.TabPage2.Controls, True, True, True)
                Me.ObjCD.KosongkanControl(Me.tbpDetails.Controls, True, True, True)
                Me.CboBankPaymentJournal.SelectedIndex = 0
                Me.CboChargeRate.SelectedIndex = 1
                Me.CboBankChargeOption.SelectedIndex = 0

                Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1"
                Me.Txt_Acc_Primary_BankCash.Text = "-1"
                Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
                Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
                Me.TxtMinimumTransaction.Text = "0"
                Me.TxtPaymentDelay.Text = "0"
                Me.TxtBankChargeNominal.Text = "0"
                Me.TxtBankChargePrecentage.Text = ObjCD.DigitGrouping(0, 2)
                Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(0, 2)
                If Me.CboBankChargeOption.SelectedIndex = 0 Then
                    Me.ChkConsistentCharge.Enabled = False
                End If
            ElseIf TabYgMana = 2 Then
                Me.ObjCD.KosongkanControl(Me.TabControl1.Controls, True, True, True)
                Me.ObjCD.KosongkanControl(Me.TabPage3.Controls, True, True, True)
                Me.ObjCD.KosongkanControl(Me.TabPage4.Controls, True, True, True)
                Me.Txt_Vou_MinimumTrans.Text = "0"
                Me.Txt_Vou_MaximumAmount.Text = "0"
                Me.Txt_Vou_PrimaryCostAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabControl1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabControl2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage3.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.TabPage4.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.tbpDetails.Controls, True, True, True)

            Me.CboBankPaymentJournal.SelectedIndex = 0
            Me.CboChargeRate.SelectedIndex = 1
            Me.CboBankChargeOption.SelectedIndex = 0
            Me.cboPaymentType.SelectedIndex = 0

            Me.TxtPrimary.Text = "-1"
            Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1"
            Me.Txt_Acc_Primary_BankCash.Text = "-1"
            Me.Txt_Acc_Primary_CreditCardCost.Text = "-1"
            Me.Txt_Acc_PrimaryCostPayable.Text = "-1"
            Me.Txt_Vou_PrimaryCostAccount.Text = "-1"

            Me.Txt_Vou_MinimumTrans.Text = "0"
            Me.Txt_Vou_MaximumAmount.Text = "0"
            Me.TxtMinimumTransaction.Text = "0"
            Me.TxtPaymentDelay.Text = "0"
            Me.TxtBankChargeNominal.Text = "0"
            Me.TxtBankChargePrecentage.Text = ObjCD.DigitGrouping(0, 2)
            Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(0, 2)

            If Me.CboBankChargeOption.SelectedIndex = 0 Then
                Me.ChkConsistentCharge.Enabled = False
            End If
            Me.ChkFillApproval.Checked = False

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ListPrimaryAkun As New List(Of Double)
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.txtPaymentCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDesc)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data default precision harus non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.txtPaymentCode.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Duplication in payment code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Duplikasi payment code, silahkan ubah code", _
                        MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.txtPaymentCode.Focus()
                    Return False
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetCode(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.txtPaymentCode.Text Then
                    If Me.Obj.ApakahSudahAda(Me.txtPaymentCode.Text) = True Then
                        If Bahasa = EnumBahasa.Inggris Then
                            MsgBox("Duplication in payment code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi payment code, silahkan ubah code", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                        Me.txtPaymentCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If Bahasa = EnumBahasa.Inggris Then
                        If MsgBox("Other user have modified this data, reload latest data?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        If MsgBox("Beberapa data telah dimodifikasi, apakah mau mengambil kembali data terakhir?", _
                        MsgBoxStyle.YesNo, "Konfirm") = MsgBoxResult.Yes Then
                            Me.Obj.CariDetail(Me.Obj.Primary, Me.ObjCD, Me)
                            Me.SalinKeField()
                        End If
                    End If
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtDesc.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input description", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan description", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtDesc.Focus()
                Return False
            End If

            If String.IsNullOrEmpty(Me.TxtRemarks.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtRemarks.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtRemarks.Focus()
                    Return False
                End If
            End If

            If Me.cboPaymentType.SelectedIndex <> 3 Then
                If String.IsNullOrEmpty(Me.TxtBankName.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input bank name", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan bank name", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtBankName.Focus()
                    Return False
                End If
                If String.IsNullOrEmpty(Me.TxtBankAccountNumber.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input account number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan account number", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtBankAccountNumber.Focus()
                    Return False
                End If
            End If

            ' Account code tidak boleh kosong 
            If Me.cboPaymentType.SelectedIndex = 3 Then
                If Me.Txt_Vou_PrimaryCostAccount.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input cost account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan cost account", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TabControl2.SelectedIndex = 1
                    Me.Txt_Vou_CostAccount.Focus()
                    Return False
                End If
            ElseIf Me.cboPaymentType.SelectedIndex = 2 Then
                If Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input bank receivable account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan bank recevable account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_BankReceivables.Focus()
                    Return False
                End If
                If Me.Txt_Acc_Primary_BankCash.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input bank cash account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan bank cash account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_BankCash.Focus()
                    Return False
                End If
            Else
                If Me.Txt_Acc_Primary_Bank_Receivable.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input bank receivable account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan bank recevable account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_BankReceivables.Focus()
                    Return False
                End If
                If Me.Txt_Acc_Primary_BankCash.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input bank cash account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan bank cash account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_BankCash.Focus()
                    Return False
                End If
                If Me.Txt_Acc_PrimaryCostPayable.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input cost payable account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan cost payable account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_CostPayable.Focus()
                    Return False
                End If
                If Me.Txt_Acc_Primary_CreditCardCost.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input credit card cost account", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan credit card cost account", MsgBoxStyle.Information, "Information")
                    End If
                    Me.TabControl1.SelectedIndex = 1
                    Me.Txt_Acc_CreditCardCost.Focus()
                    Return False
                End If
            End If

            ' Cek Akun yang dipergunakan 
            If Me.cboPaymentType.SelectedIndex = 0 Or Me.cboPaymentType.SelectedIndex = 1 Then
                If ListPrimaryAkun.Contains(Me.Txt_Acc_Primary_Bank_Receivable.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Bank receivable account has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Bank receivable account sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                End If
                If ListPrimaryAkun.Contains(Me.Txt_Acc_Primary_BankCash.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Bank cash account has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Bank cash account sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_Primary_BankCash.Text)
                End If
                If ListPrimaryAkun.Contains(Me.Txt_Acc_PrimaryCostPayable.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Cost payable has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Cost payable sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_PrimaryCostPayable.Text)
                End If
                If ListPrimaryAkun.Contains(Me.Txt_Acc_Primary_CreditCardCost.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Credit card cost has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Credit card cost sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_Primary_CreditCardCost.Text)
                End If
            End If
            ''
            If Me.cboPaymentType.SelectedIndex = 2 Then
                If ListPrimaryAkun.Contains(Me.Txt_Acc_Primary_Bank_Receivable.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Bank receivable account has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Bank receivable account sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_Primary_Bank_Receivable.Text)
                End If
                If ListPrimaryAkun.Contains(Me.Txt_Acc_Primary_BankCash.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Bank cash account has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Bank cash account sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Acc_Primary_BankCash.Text)
                End If
            End If
            ''
            If Me.cboPaymentType.SelectedIndex = 3 Then
                If ListPrimaryAkun.Contains(Me.Txt_Vou_PrimaryCostAccount.Text) = True Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Cost account has been used in another account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Cost account sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.Txt_Vou_PrimaryCostAccount.Text)
                End If
            End If

            'cek number tidak boleh negatif 
            If Me.cboPaymentType.SelectedIndex <> 3 Then
                If ObjCD.CekBilanganPositif(Me.TxtMinimumTransaction.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TabControl1.SelectedIndex = 0
                    Me.TxtMinimumTransaction.Focus()
                    Return False
                End If
                'If ObjCD.CekBilanganPositif(Me.TxtPaymentDelay.Text) = False Then
                'If Bahasa = EnumBahasa.Inggris Then
                'MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                'ElseIf Bahasa = EnumBahasa.Indonesia Then
                'MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                'End If
                'Me.TabControl1.SelectedIndex = 0
                'Me.TxtPaymentDelay.Focus()
                'Return False
                'End If
                If Me.CboBankChargeOption.SelectedIndex <> 0 Then
                    If Me.CboChargeRate.SelectedIndex = 0 Then
                        If ObjCD.CekBilanganPositif(Me.TxtBankChargeNominal.Text) = False Then
                            If Bahasa = EnumBahasa.Inggris Then
                                MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            ElseIf Bahasa = EnumBahasa.Indonesia Then
                                MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                            End If
                            Me.TabControl1.SelectedIndex = 0
                            Me.TxtBankChargeNominal.Focus()
                            Return False
                        End If
                    ElseIf Me.CboChargeRate.SelectedIndex = 1 Then
                        If Me.ChkConsistentCharge.Checked = True Then
                            If ObjCD.CekBilanganPositif(Me.TxtBankChargePrecentage.Text) = False Then
                                If Bahasa = EnumBahasa.Inggris Then
                                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                                ElseIf Bahasa = EnumBahasa.Indonesia Then
                                    MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                                End If
                                Me.TabControl1.SelectedIndex = 0
                                Me.TxtBankChargePrecentage.Focus()
                                Return False
                            End If
                        Else
                            If ObjCD.CekBilanganPositif(Me.TxtBankChargePrecentage.Text) = False Then
                                If Bahasa = EnumBahasa.Inggris Then
                                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                                ElseIf Bahasa = EnumBahasa.Indonesia Then
                                    MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                                End If
                                Me.TabControl1.SelectedIndex = 0
                                Me.TxtBankChargePrecentage.Focus()
                                Return False
                            End If
                            If ObjCD.CekBilanganPositif(Me.TxtBankCustChargePercentage.Text) = False Then
                                If Bahasa = EnumBahasa.Inggris Then
                                    MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                                ElseIf Bahasa = EnumBahasa.Indonesia Then
                                    MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                                End If
                                Me.TabControl1.SelectedIndex = 0
                                Me.TxtBankCustChargePercentage.Focus()
                                Return False
                            End If
                        End If
                    End If
                End If
            Else
                If ObjCD.CekBilanganPositif(Me.Txt_Vou_MinimumTrans.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TabControl2.SelectedIndex = 0
                    Me.Txt_Vou_MinimumTrans.Focus()
                    Return False
                End If
                If ObjCD.CekBilanganPositif(Me.Txt_Vou_MaximumAmount.Text) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan nilai positif", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TabControl2.SelectedIndex = 0
                    Me.Txt_Vou_MaximumAmount.Focus()
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtBankName.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtBankName.Text, False) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtBankName.Focus()
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(Me.TxtBankAccountNumber.Text) = False Then
                If ObjCD.CekDataKosong(Me.TxtBankAccountNumber.Text, False) = False Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TxtBankAccountNumber.Focus()
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ListPrimaryAkun = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .Code = Me.txtPaymentCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text
                If Me.cboPaymentType.SelectedIndex = 0 Then
                    .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Credit_Card_Payment
                ElseIf Me.cboPaymentType.SelectedIndex = 1 Then
                    .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Debit_Card_Payment
                ElseIf Me.cboPaymentType.SelectedIndex = 2 Then
                    .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Cash_Withdrawal
                ElseIf Me.cboPaymentType.SelectedIndex = 3 Then
                    .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Voucher_Payment
                End If
                .BankName = Me.TxtBankName.Text
                .BankAccNumber = Me.TxtBankAccountNumber.Text
                If Me.cboPaymentType.SelectedIndex = 3 Then
                    .MinimumTransAmount = Me.Txt_Vou_MinimumTrans.Text
                Else
                    .MinimumTransAmount = Me.TxtMinimumTransaction.Text
                End If
                .PaymentDelay = Me.TxtPaymentDelay.Text
                If Me.CboBankPaymentJournal.SelectedIndex = 0 Then
                    .PaymentJurnal = ClsMast_PaymentOption.EnumPaymentJurnal.Automatic
                ElseIf Me.CboBankPaymentJournal.SelectedIndex = 1 Then
                    .PaymentJurnal = ClsMast_PaymentOption.EnumPaymentJurnal.Manual
                End If
                If Me.CboBankChargeOption.SelectedIndex = 0 Then
                    .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.No_Charge
                ElseIf Me.CboBankChargeOption.SelectedIndex = 1 Then
                    .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.Apply_Charge_For_Customer
                ElseIf Me.CboBankChargeOption.SelectedIndex = 2 Then
                    .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.Apply_Charge_As_Internal_Cost
                End If
                If Me.CboChargeRate.SelectedIndex = 0 Then
                    .ChangeRate = ClsMast_PaymentOption.EnumChargeRate.Fixed_Nominal_Rate
                ElseIf Me.CboChargeRate.SelectedIndex = 1 Then
                    .ChangeRate = ClsMast_PaymentOption.EnumChargeRate.Percentage
                End If
                If Me.CboBankChargeOption.SelectedIndex = 1 Then
                    If ChkConsistentCharge.Checked = True Then
                        .ConsistentChange = 1
                    Else
                        .ConsistentChange = 0
                    End If
                Else
                    If ChkConsistentCharge.Checked = True Then
                        .ConsistentChange = 1
                    Else
                        .ConsistentChange = 0
                    End If
                End If
                .BankChargeNominal = Me.TxtBankChargeNominal.Text
                .BankChargePercentage = Me.TxtBankChargePrecentage.Text
                If Me.CboChargeRate.SelectedIndex = 1 Then
                    If Me.ChkConsistentCharge.Checked = False Then
                        .CustChargePercentage = Me.TxtBankCustChargePercentage.Text
                    Else
                        .CustChargePercentage = Me.TxtBankChargePrecentage.Text
                    End If
                End If
                .PrimaryAccBankReceivable = Me.Txt_Acc_Primary_Bank_Receivable.Text
                .PrimaryAccBankCash = Me.Txt_Acc_Primary_BankCash.Text
                .PrimaryAccCostPayable = Me.Txt_Acc_PrimaryCostPayable.Text
                .PrimaryAccCreditCardCost = Me.Txt_Acc_Primary_CreditCardCost.Text
                .PrimaryAccVoucherCost = Me.Txt_Vou_PrimaryCostAccount.Text
                .MaxVoucherAmount = Me.Txt_Vou_MaximumAmount.Text
                If Me.ChkFillApproval.Checked = True Then
                    .chkFillApproval = True
                Else
                    .chkFillApproval = False
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.txtPaymentCode.Text = .Code
                Me.TxtDesc.Text = .Desc
                Me.TxtRemarks.Text = .Remarks
                If .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Credit_Card_Payment Then
                    Me.cboPaymentType.SelectedIndex = 0
                ElseIf .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Debit_Card_Payment Then
                    Me.cboPaymentType.SelectedIndex = 1
                ElseIf .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Cash_Withdrawal Then
                    Me.cboPaymentType.SelectedIndex = 2
                ElseIf .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Voucher_Payment Then
                    Me.cboPaymentType.SelectedIndex = 3
                End If
                Me.TxtBankName.Text = .BankName
                Me.TxtBankAccountNumber.Text = .BankAccNumber
                If .TipePayment = ClsMast_PaymentOption.EnumTipePayment.Voucher_Payment Then
                    Me.Txt_Vou_MinimumTrans.Text = ObjCD.DigitGrouping(.MinimumTransAmount, Me.Def_Prec)
                Else
                    Me.TxtMinimumTransaction.Text = ObjCD.DigitGrouping(.MinimumTransAmount, Me.Def_Prec)
                End If
                Me.TxtPaymentDelay.Text = ObjCD.DigitGrouping(.PaymentDelay, 0)
                If .PaymentJurnal = ClsMast_PaymentOption.EnumPaymentJurnal.Automatic Then
                    Me.CboBankPaymentJournal.SelectedIndex = 0
                ElseIf .PaymentJurnal = ClsMast_PaymentOption.EnumPaymentJurnal.Manual Then
                    Me.CboBankPaymentJournal.SelectedIndex = 1
                End If
                If .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.No_Charge Then
                    Me.CboBankChargeOption.SelectedIndex = 0
                ElseIf .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.Apply_Charge_For_Customer Then
                    Me.CboBankChargeOption.SelectedIndex = 1
                ElseIf .BankChargeOption = ClsMast_PaymentOption.EnumBankChargeOption.Apply_Charge_As_Internal_Cost Then
                    Me.CboBankChargeOption.SelectedIndex = 2
                End If
                If .ChangeRate = ClsMast_PaymentOption.EnumChargeRate.Fixed_Nominal_Rate Then
                    Me.CboChargeRate.SelectedIndex = 0
                ElseIf .ChangeRate = ClsMast_PaymentOption.EnumChargeRate.Percentage Then
                    Me.CboChargeRate.SelectedIndex = 1
                End If
                If .ChangeRate = ClsMast_PaymentOption.EnumChargeRate.Percentage Then
                    If .ConsistentChange = 1 Then
                        Me.ChkConsistentCharge.Checked = True
                    Else
                        Me.ChkConsistentCharge.Checked = False
                    End If
                Else
                    If .ConsistentChange = 1 Then
                        Me.ChkConsistentCharge.Checked = True
                    Else
                        Me.ChkConsistentCharge.Checked = False
                    End If
                End If
                Me.TxtBankChargeNominal.Text = ObjCD.DigitGrouping(.BankChargeNominal, Me.Def_Prec)
                Me.TxtBankChargePrecentage.Text = ObjCD.DigitGrouping(.BankChargePercentage, 2)
                Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(.CustChargePercentage, 2)
                Me.Txt_Acc_Primary_Bank_Receivable.Text = .PrimaryAccBankReceivable
                Me.Txt_Acc_BankReceivables.Text = ObjAkun.GetAccountCode(.PrimaryAccBankReceivable)
                Me.Txt_Acc_Primary_BankCash.Text = .PrimaryAccBankCash
                Me.Txt_Acc_BankCash.Text = ObjAkun.GetAccountCode(.PrimaryAccBankCash)
                Me.Txt_Acc_PrimaryCostPayable.Text = .PrimaryAccCostPayable
                Me.Txt_Acc_CostPayable.Text = ObjAkun.GetAccountCode(.PrimaryAccCostPayable)
                Me.Txt_Acc_Primary_CreditCardCost.Text = .PrimaryAccCreditCardCost
                Me.Txt_Acc_CreditCardCost.Text = ObjAkun.GetAccountCode(.PrimaryAccCreditCardCost)
                Me.Txt_Vou_PrimaryCostAccount.Text = .PrimaryAccVoucherCost
                Me.Txt_Vou_CostAccount.Text = ObjAkun.GetAccountCode(.PrimaryAccVoucherCost)
                Me.Txt_Vou_MaximumAmount.Text = ObjCD.DigitGrouping(.MaxVoucherAmount, Me.Def_Prec)
                If Me.cboPaymentType.SelectedIndex = 3 Then
                    Me.TandaTab2Aktif = True
                Else
                    Me.TandaTab2Aktif = False
                End If
                If .chkFillApproval = True Then
                    Me.ChkFillApproval.Checked = True
                Else
                    Me.ChkFillApproval.Checked = False
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                        Me.TxtPrimary.Text, Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.txtPaymentCode.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.txtPaymentCode.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Obj.ResetObject()
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.CboChargeRate.SelectedIndex = 1
            Me.TabControl1.BringToFront()
            Me.TabControl1.SelectedIndex = 0
            Me.TandaTab2Aktif = False
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.txtPaymentCode.Focus()
            Me.txtPaymentCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmMast_PaymentOption_Cari
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowDialog()

            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtPaymentCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPaymentCode.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True, Bahasa) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.txtPaymentCode.Focus()
                    Me.txtPaymentCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtPaymentCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtPaymentCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                 Bahasa, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, True, True, _
                        Bahasa, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            Me.NonAktifkanDetail()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtPaymentCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, _
                            Me.TxtPrimary.Text, Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinimumTransaction.Enter
        Try
            ObjCD.Number_Enter(Me.TxtMinimumTransaction, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMinimumTransaction.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtMinimumTransaction, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMinimumTransaction.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtMinimumTransaction, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinimumTransaction.Leave
        Try
            ObjCD.Number_Leave(Me.TxtMinimumTransaction, sender, e, Me.Def_Prec, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDelay.Enter
        Try
            ObjCD.Number_Enter(Me.TxtPaymentDelay, sender, e)
            Me.buffer = Me.TxtPaymentDelay.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentDelay.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtPaymentDelay, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentDelay.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtPaymentDelay, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDelay.Leave
        Try
            ObjCD.Number_Leave(Me.TxtPaymentDelay, sender, e, Me.Def_Prec, ObjCD)
            If ObjCD.CekNonNegatif(Me.TxtPaymentDelay.Text) = False Then
                Me.TxtPaymentDelay.Text = ObjCD.DigitGrouping(Me.buffer, 0, ClsCekData.EnumTipeRounding.Middlepoint)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargeNominal.Enter
        Try
            ObjCD.Number_Enter(Me.TxtBankChargeNominal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankChargeNominal.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtBankChargeNominal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankChargeNominal.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtBankChargeNominal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargeNominal.Leave
        Try
            ObjCD.Number_Leave(Me.TxtBankChargeNominal, sender, e, Me.Def_Prec, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargePrecentage.Enter
        Try
            ObjCD.Number_Enter(Me.TxtBankChargePrecentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankChargePrecentage.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtBankChargePrecentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankChargePrecentage.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtBankChargePrecentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargePrecentage.Leave
        Try
            ObjCD.Number_Leave(Me.TxtBankChargePrecentage, sender, e, 2, ObjCD)
            If Me.ChkConsistentCharge.Checked = True Then
                Me.TxtBankCustChargePercentage.Text = ObjCD.DigitGrouping(Me.TxtBankChargePrecentage.Text, 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankCustChargePercentage.Enter
        Try
            ObjCD.Number_Enter(Me.TxtBankCustChargePercentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankCustChargePercentage.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtBankCustChargePercentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBankCustChargePercentage.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtBankCustChargePercentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankCustChargePercentage.Leave
        Try
            ObjCD.Number_Leave(Me.TxtBankCustChargePercentage, sender, e, 2, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MaximumAmount.Enter
        Try
            ObjCD.Number_Enter(Me.Txt_Vou_MaximumAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Vou_MaximumAmount.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.Txt_Vou_MaximumAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Vou_MaximumAmount.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.Txt_Vou_MaximumAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MaximumAmount.Leave
        Try
            ObjCD.Number_Leave(Me.Txt_Vou_MaximumAmount, sender, e, Me.Def_Prec, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MinimumTrans.Enter
        Try
            ObjCD.Number_Enter(Me.Txt_Vou_MinimumTrans, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Vou_MinimumTrans.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.Txt_Vou_MinimumTrans, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Vou_MinimumTrans.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.Txt_Vou_MinimumTrans, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MinimumTrans.Leave
        Try
            ObjCD.Number_Leave(Me.Txt_Vou_MinimumTrans, sender, e, Me.Def_Prec, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Back Color 
    Private Sub txtPaymentCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtPaymentCode.Enabled = True Then
                    If Me.txtPaymentCode.ReadOnly = False Then
                        Me.txtPaymentCode.BackColor = Me.BackColorAll
                    Else
                        Me.txtPaymentCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtPaymentCode.Enabled = True Then
                    If Me.txtPaymentCode.ReadOnly = False Then
                        Me.txtPaymentCode.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.txtPaymentCode.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        Me.TxtDesc.BackColor = Me.BackColorAll
                    Else
                        Me.TxtDesc.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtDesc.Enabled = True Then
                    If Me.TxtDesc.ReadOnly = False Then
                        Me.TxtDesc.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtDesc.BackColor = Me.TxtReadOnly.BackColor
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
                        Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankName.Enabled = True Then
                    If Me.TxtBankName.ReadOnly = False Then
                        Me.TxtBankName.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBankName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankName.Enabled = True Then
                    If Me.TxtBankName.ReadOnly = False Then
                        Me.TxtBankName.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBankName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankAccountNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankAccountNumber.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankAccountNumber.Enabled = True Then
                    If Me.TxtBankAccountNumber.ReadOnly = False Then
                        Me.TxtBankAccountNumber.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBankAccountNumber.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankAccountNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankAccountNumber.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankAccountNumber.Enabled = True Then
                    If Me.TxtBankAccountNumber.ReadOnly = False Then
                        Me.TxtBankAccountNumber.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBankAccountNumber.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinimumTransaction.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinimumTransaction.Enabled = True Then
                    If Me.TxtMinimumTransaction.ReadOnly = False Then
                        Me.TxtMinimumTransaction.BackColor = Me.BackColorAll
                    Else
                        Me.TxtMinimumTransaction.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMinimumTransaction_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMinimumTransaction.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtMinimumTransaction.Enabled = True Then
                    If Me.TxtMinimumTransaction.ReadOnly = False Then
                        Me.TxtMinimumTransaction.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtMinimumTransaction.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDelay.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaymentDelay.Enabled = True Then
                    If Me.TxtPaymentDelay.ReadOnly = False Then
                        Me.TxtPaymentDelay.BackColor = Me.BackColorAll
                    Else
                        Me.TxtPaymentDelay.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDelay_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDelay.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtPaymentDelay.Enabled = True Then
                    If Me.TxtPaymentDelay.ReadOnly = False Then
                        Me.TxtPaymentDelay.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtPaymentDelay.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargeNominal.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankChargeNominal.Enabled = True Then
                    If Me.TxtBankChargeNominal.ReadOnly = False Then
                        Me.TxtBankChargeNominal.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBankChargeNominal.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargeNominal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargeNominal.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankChargeNominal.Enabled = True Then
                    If Me.TxtBankChargeNominal.ReadOnly = False Then
                        Me.TxtBankChargeNominal.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBankChargeNominal.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargePrecentage.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankChargePrecentage.Enabled = True Then
                    If Me.TxtBankChargePrecentage.ReadOnly = False Then
                        Me.TxtBankChargePrecentage.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBankChargePrecentage.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankChargePrecentage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankChargePrecentage.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankChargePrecentage.Enabled = True Then
                    If Me.TxtBankChargePrecentage.ReadOnly = False Then
                        Me.TxtBankChargePrecentage.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBankChargePrecentage.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankCustChargePercentage.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankCustChargePercentage.Enabled = True Then
                    If Me.TxtBankCustChargePercentage.ReadOnly = False Then
                        Me.TxtBankCustChargePercentage.BackColor = Me.BackColorAll
                    Else
                        Me.TxtBankCustChargePercentage.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBankCustChargePercentage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBankCustChargePercentage.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtBankCustChargePercentage.Enabled = True Then
                    If Me.TxtBankCustChargePercentage.ReadOnly = False Then
                        Me.TxtBankCustChargePercentage.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtBankCustChargePercentage.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_Primary_Bank_Receivable_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_Primary_Bank_Receivable.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_Primary_Bank_Receivable.Enabled = True Then
                    If Me.Txt_Acc_Primary_Bank_Receivable.ReadOnly = False Then
                        Me.Txt_Acc_Primary_Bank_Receivable.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Acc_Primary_Bank_Receivable.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_Primary_Bank_Receivable_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_Primary_Bank_Receivable.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_Primary_Bank_Receivable.Enabled = True Then
                    If Me.Txt_Acc_Primary_Bank_Receivable.ReadOnly = False Then
                        Me.Txt_Acc_Primary_Bank_Receivable.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Acc_Primary_Bank_Receivable.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_BankCash_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_BankCash.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_BankCash.Enabled = True Then
                    If Me.Txt_Acc_BankCash.ReadOnly = False Then
                        Me.Txt_Acc_BankCash.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Acc_BankCash.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_BankCash_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_BankCash.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_BankCash.Enabled = True Then
                    If Me.Txt_Acc_BankCash.ReadOnly = False Then
                        Me.Txt_Acc_BankCash.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Acc_BankCash.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_CostPayable_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_CostPayable.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_CostPayable.Enabled = True Then
                    If Me.Txt_Acc_CostPayable.ReadOnly = False Then
                        Me.Txt_Acc_CostPayable.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Acc_CostPayable.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_CostPayable_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_CostPayable.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_CostPayable.Enabled = True Then
                    If Me.Txt_Acc_CostPayable.ReadOnly = False Then
                        Me.Txt_Acc_CostPayable.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Acc_CostPayable.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_CreditCardCost_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_CreditCardCost.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_CreditCardCost.Enabled = True Then
                    If Me.Txt_Acc_CreditCardCost.ReadOnly = False Then
                        Me.Txt_Acc_CreditCardCost.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Acc_CreditCardCost.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_CreditCardCost_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_CreditCardCost.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_CreditCardCost.Enabled = True Then
                    If Me.Txt_Acc_CreditCardCost.ReadOnly = False Then
                        Me.Txt_Acc_CreditCardCost.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Acc_CreditCardCost.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MinimumTrans.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_MinimumTrans.Enabled = True Then
                    If Me.Txt_Vou_MinimumTrans.ReadOnly = False Then
                        Me.Txt_Vou_MinimumTrans.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Vou_MinimumTrans.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MinimumTrans_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MinimumTrans.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_MinimumTrans.Enabled = True Then
                    If Me.Txt_Vou_MinimumTrans.ReadOnly = False Then
                        Me.Txt_Vou_MinimumTrans.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Vou_MinimumTrans.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MaximumAmount.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_MaximumAmount.Enabled = True Then
                    If Me.Txt_Vou_MaximumAmount.ReadOnly = False Then
                        Me.Txt_Vou_MaximumAmount.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Vou_MaximumAmount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_MaximumAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_MaximumAmount.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_MaximumAmount.Enabled = True Then
                    If Me.Txt_Vou_MaximumAmount.ReadOnly = False Then
                        Me.Txt_Vou_MaximumAmount.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Vou_MaximumAmount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_BankReceivables_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_BankReceivables.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_BankReceivables.Enabled = True Then
                    If Me.Txt_Acc_BankReceivables.ReadOnly = False Then
                        Me.Txt_Acc_BankReceivables.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Acc_BankReceivables.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Acc_BankReceivables_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Acc_BankReceivables.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Acc_BankReceivables.Enabled = True Then
                    If Me.Txt_Acc_BankReceivables.ReadOnly = False Then
                        Me.Txt_Acc_BankReceivables.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Acc_BankReceivables.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_CostAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_CostAccount.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_CostAccount.Enabled = True Then
                    If Me.Txt_Vou_CostAccount.ReadOnly = False Then
                        Me.Txt_Vou_CostAccount.BackColor = Me.BackColorAll
                    Else
                        Me.Txt_Vou_CostAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Vou_CostAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Vou_CostAccount.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.Txt_Vou_CostAccount.Enabled = True Then
                    If Me.Txt_Vou_CostAccount.ReadOnly = False Then
                        Me.Txt_Vou_CostAccount.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.Txt_Vou_CostAccount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class