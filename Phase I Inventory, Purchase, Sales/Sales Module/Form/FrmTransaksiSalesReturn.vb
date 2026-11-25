Option Explicit On
Option Strict Off

Imports Core_Module
Imports Inventory_Module

Public Class FrmTransaksiSalesReturn
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0
    Public ObjDO As New ClsTransaksiDO

    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjSRItem As New ClsTransaksiSalesReturnItem
    Dim ObjSalesReturn As New ClsTransaksiSalesReturn
    Dim ObjStockInOut As New clsTransaksiStockInOut
    Dim OBJCL As New ClsTransaksiPurchaseClosing
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjCur As New clsMasterCurrency

    Dim Baru As Boolean
    Dim buffer As String = ""
    Dim DS As New dseTransaksiSalesReturn
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Dim DseAccount As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As DseMultiAccount.tblCurrencyTaxRow
    Dim ObjAkun As New Core_Module.ClsMasterAkun1

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Receivable
        Dim objMulti As New ClsMultiAccount
        Try

            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryAccount.Items.Clear()
            Me.tblAccount.Rows.Clear()
            Me.txtPrimaryAccount.Text = -1

            If Me.cboTipeSR.SelectedIndex = 0 Then
                Me.cboAccount.Enabled = False
                Me.cboPrimaryAccount.Enabled = False
                Me.txtPrimaryAccount.Text = -1
            Else
                Me.cboAccount.Enabled = True
                Me.cboPrimaryAccount.Enabled = True
                objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccount, Me.TxtPrimaryCurrency.Text)
                Me.cboAccount.DataSource = Me.tblAccount.DefaultView
                Me.cboAccount.DisplayMember = "Account Code"
                Me.cboPrimaryAccount.DataSource = Me.tblAccount.DefaultView
                Me.cboPrimaryAccount.DisplayMember = "Primary"
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If
           

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOCurrencyTax()
        '- isicbo customer
        Dim objMulti As New ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboCurrTax.DataSource = Nothing
            Me.cboCurrTax.Items.Clear()
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.tblCurrTax.Rows.Clear()
            Me.txtPrimaryCurrTax.Text = -1

            If Me.cboTipeSR.SelectedIndex = 0 Then
                Me.cboCurrTax.Enabled = False
                Me.cboPrimaryCurrTax.Enabled = False
                Me.txtPrimaryCurrTax.Text = -1
            Else
                Me.cboCurrTax.Enabled = True
                Me.cboPrimaryCurrTax.Enabled = True
                objMulti.IsiTBLCurrTax(Me.TxtPrimaryCurrency.Text, Me.tblCurrTax)
                If Me.tblCurrTax.Rows.Count > 0 Then
                    Me.cboCurrTax.DataSource = Me.tblCurrTax.DefaultView
                    Me.cboCurrTax.DisplayMember = "Currency Code"
                    Me.cboPrimaryCurrTax.DataSource = Me.tblCurrTax.DefaultView
                    Me.cboPrimaryCurrTax.DisplayMember = "Primary"
                    Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                    Me.lblCurrTax.Text = Me.cboCurrTax.Text

                End If
            End If
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCurrency.TextChanged
        Try
            If Me.TxtRemark.Enabled = True AndAlso _
            IsNumeric(Me.TxtPrimaryCurrency.Text) AndAlso _
            (Me.TxtPrimaryCurrency.Text) > 0 Then
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOArAkunByTipePerkiraan1()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If Me.TxtRemark.Enabled = True AndAlso IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutystamp.Text) Then
                    Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.HitungTotal()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub cboPrimaryCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutystamp.Text) Then
                    Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.HitungTotal()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrTax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrTax.TextChanged
        Dim PrimaryCurrTax As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryCurrTax.Text) Then
                PrimaryCurrTax = CDbl(Me.txtPrimaryCurrTax.Text)
            End If

            If Me.tblCurrTax.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblCurrTax.Rows.Count - 1
                    Me.rowCurrTax = Me.tblCurrTax.Rows(IntA)
                    With Me.rowCurrTax
                        If .Primary = PrimaryCurrTax Then
                            Me.cboPrimaryCurrTax.SelectedIndex = IntA
                            Me.cboCurrTax.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAccount.SelectedIndexChanged
        Try
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
        Dim PrimaryARAkun As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
            End If

            If Me.tblAccount.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblAccount.Rows.Count - 1
                    Me.rowAccount = Me.tblAccount.Rows(IntA)
                    With Me.rowAccount
                        If .Primary = PrimaryARAkun Then
                            Me.cboPrimaryAccount.SelectedIndex = IntA
                            Me.cboAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungTotal()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal, tax, TaxExchangeRate, TaxAmount, Total, dutyStamp, subTotalAfterDP As Decimal
        Dim Rounding As Integer = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecCurrency As Double = 0
        Dim DefPrec As Double = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.HitungTotal_middlePoint()
                Exit Sub
            End If
            'hitung pajak
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrec
            End If
            DefPrecCurrency = Me.ObjCur.GetDecimal_DefaultCurrency
            DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)

            If Not IsNumeric(Me.txtTax.Text) Then
                tax = 0
            Else
                tax = CDec(Me.txtTax.Text)
            End If

            If Not IsNumeric(Me.txtTaxExchangeRate.Text) Then
                TaxExchangeRate = 0
            Else
                TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
            End If

            If Not IsNumeric(Me.txtTaxAmount.Text) Then
                TaxAmount = 0
            Else
                TaxAmount = CDec(Me.txtTaxAmount.Text)
            End If

            If Not IsNumeric(Me.TxtDutystamp.Text) Then
                dutyStamp = 0
            Else
                dutyStamp = CDec(TxtDutystamp.Text)
            End If



            If Me.cboTipeSR.SelectedIndex = 1 Or Me.cboTipeSR.SelectedIndex = 2 Then
                'hitung subtotal
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    End If
                Next
                subTotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))

                If Me.chkTax.Checked = True Then
                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                    If Me.TxtPrimaryCurrency.Text = "1" Then
                        TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(tax * _
                        subTotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * tax * _
                                                  subTotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                        Else

                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(tax * _
                                                  subTotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                        End If
                    End If
                End If
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(dutyStamp, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecTax)
            Else
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(dutyStamp, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecTax)
            End If


            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    decTotal += CDec(Me.TxtDutystamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.TxtDutystamp.Text) Then
                        decTotal += CDec(Me.TxtDutystamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub HitungTotal_middlePoint()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal, tax, TaxExchangeRate, TaxAmount, Total, dutyStamp, subTotalAfterDP As Decimal
        Dim Rounding As Integer = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecCurrency As Double = 0
        Dim DefPrec As Double = 0

        Try
            'hitung pajak
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrec
            End If
            DefPrecCurrency = Me.ObjCur.GetDecimal_DefaultCurrency
            DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)

            If Not IsNumeric(Me.txtTax.Text) Then
                tax = 0
            Else
                tax = CDec(Me.txtTax.Text)
            End If

            If Not IsNumeric(Me.txtTaxExchangeRate.Text) Then
                TaxExchangeRate = 0
            Else
                TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
            End If

            If Not IsNumeric(Me.txtTaxAmount.Text) Then
                TaxAmount = 0
            Else
                TaxAmount = CDec(Me.txtTaxAmount.Text)
            End If

            If Not IsNumeric(Me.TxtDutystamp.Text) Then
                dutyStamp = 0
            Else
                dutyStamp = CDec(TxtDutystamp.Text)
            End If



            If Me.cboTipeSR.SelectedIndex = 1 Or Me.cboTipeSR.SelectedIndex = 2 Then
                'hitung subtotal
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    End If
                Next
                subTotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))

                If Me.chkTax.Checked = True Then
                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                    If Me.TxtPrimaryCurrency.Text = "1" Then
                        TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(tax * _
                        subTotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * tax * _
                                                  subTotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                        Else

                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(tax * _
                                                  subTotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                        End If
                    End If
                End If
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(dutyStamp, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecTax)
            Else
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(dutyStamp, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecTax)
            End If


            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    decTotal += CDec(Me.TxtDutystamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.TxtDutystamp.Text) Then
                        decTotal += CDec(Me.TxtDutystamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtDutystamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Leave
        Try
            '  Me.ObjCD.Number_Leave(Me.TxtDutystamp, sender, e, Me.ObjCur.GetDecimal_DefaultCurrency, Me.ObjCD)
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.ObjSalesReturn.ResetObject(Me.dgvItem)
            Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.GetNoBaru(Me.dtpDate.Value)
            Me.TxtPrimaryCurrency.Text = "1"
            Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            'default tipe sales return adalah ganti barang
            Me.cboTipeSR.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryStockIn.Text = "-1"
            Me.TxtPrimarySalesReturn.Text = "-1"
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryDO.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.TxtPaymentTerm.Text = "0"
            Me.TxtDutystamp.Text = 0
            Me.dgvItem.Rows.Clear()

            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()
            Me.cboTipeSR.SelectedIndex = 0
            Me.ChkInvoiceCancellation.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimarySalesReturn.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.chkTaxReturn.Enabled = True

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            cboTipeSR_SelectedIndexChanged(Nothing, Nothing)

            If Me.TxtPrimarySalesReturn.Text <> -1 Then
                If Me.ObjSalesReturn.ApakahHapusMengakibatkanMinus2(Me.TxtPrimarySalesReturn.Text, False) = True Then
                    Me.btnFindStockIn.Enabled = False
                    Me.TxtStockInNo.Enabled = False
                End If
            End If

            Me.dgvItem.ReadOnly = False
            Me.dtpInvoiceDate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.chkTaxReturn.Enabled = False

            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim Rounding As Integer = 0
        Dim DefPrecBaseCurrency As Integer = 0

        Try
            DefPrecBaseCurrency = Me.ObjCur.GetDecimal_DefaultCurrency
            With Me.ObjSalesReturn
                Me.TxtSalesReturnNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimarySalesReturn.Text = .Primary
                Me.txtPrimaryStockIn.Text = .PrimaryStockInSalesReturn
                Me.IsiDataStockIn()
                Me.TxtReasonOfReturn.Text = .ReturnReason
                Me.txtPrimaryDO.Text = .PrimaryDO
                Me.chkTaxReturn.Checked = .CetakNotaReturn
                Me.TxtNotaRetur.Text = .NoNotaRetur
                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency
                If .TipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak Then
                    Me.cboTipeSR.SelectedIndex = 1
                Else
                    Me.cboTipeSR.SelectedIndex = 0
                End If

                If .InvoiceCancellation = 1 Then
                    Me.cboTipeSR.SelectedIndex = 2
                    Me.ChkInvoiceCancellation.Checked = True
                Else
                    Me.ChkInvoiceCancellation.Checked = False
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                If .PrimaryCurrency <> -1 Then
                    Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    Rounding = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                End If

                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    Me.cboPrimaryAccount.Text = .PrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOCurrencyTax()
                If .PrimaryCurrTax > -1 Then
                    Me.txtPrimaryCurrTax.Text = .PrimaryCurrTax
                    Me.cboPrimaryCurrTax.Text = .PrimaryCurrTax
                    Me.lblCurrTax.Text = Me.ObjCur.GetCurrencyCode(.PrimaryCurrTax)
                Else
                    Me.cboCurrTax.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                End If

                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(.Dutystamp, Me.ObjCur.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCur.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Rounding))
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                Me.chkTax.Checked = .UseTax
                Me.txtTax.Text = ObjCD.DigitGrouping(.TaxPercent, 2)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Rounding))
                Me.TxtPaymentTerm.Text = Me.ObjCD.DigitGrouping(.PaymentTerm, 0)

                'Update NPWP 16 Digit
                Me.txtNoNPWP.Text = .NPWPNo
            End With
            Me.FormatDGV()
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjSalesReturn
                .ReturnReason = Me.TxtReasonOfReturn.Text
                .NoTransaksi = Me.TxtSalesReturnNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimaryStockInSalesReturn = Me.txtPrimaryStockIn.Text
                .Remarks = Me.TxtRemark.Text
                .PrimaryDO = Me.txtPrimaryDO.Text
                .CetakNotaReturn = Me.chkTaxReturn.Checked
                .NoNotaRetur = Me.TxtNotaRetur.Text
                If Me.cboTipeSR.SelectedIndex = 0 Then
                    .TipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Revisi_Invoice
                Else
                    .TipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
                End If
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text

                .Dutystamp = Me.TxtDutystamp.Text
                .SubTotal = Me.txtSubTotalAfterDP.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                .UseTax = Me.chkTax.Checked
                .Total = Me.TxtTotal.Text
                .PaymentTerm = Math.Round(CDbl(Me.TxtPaymentTerm.Text), 0, MidpointRounding.AwayFromZero)
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryCustomer = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                .TaxPercent = Me.txtTax.Text

                If ChkInvoiceCancellation.Checked = True Then
                    .InvoiceCancellation = 1
                Else
                    .InvoiceCancellation = 0
                End If

                'Update NPWP 16 Digit.
                .NPWPNo = Me.txtNoNPWP.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim TanggalInvoice As Date

        Try
            'jika retur nya adalah nota retur pajak, maka nomor nota returnya kosong
            If Me.cboTipeSR.SelectedIndex = 0 Then
                Me.TxtNotaRetur.Clear()
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("sales return  number")
            KoleksiTextBox.Add(Me.TxtSalesReturnNo)
            KoleksiNamaData.Add("stock in number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            If Me.cboTipeSR.SelectedIndex = 1 Or Me.cboTipeSR.SelectedIndex = 2 Then
                KoleksiNamaData.Add("return note")
                KoleksiTextBox.Add(Me.TxtNotaRetur)
            End If
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'periode sales return dan stock in sales return harus sama
            If Format(Me.ObjStockInOut.GetTanggalTransaksi(Me.txtPrimaryStockIn.Text), "MMM yyyy") <> Format(Me.dtpDate.Value, "MMM yyyy") Then
                MsgBox("The stock in sales return and sales return must be dated on the same month. Please fix the data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan transaksi closing
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.Baru = False Then
                If Me.ObjSalesReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjSalesReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjSalesReturn.TanggalTransaksi
                        Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.NoTransaksi
                        Return False
                    End If
                End If
            End If

            If Me.cboTipeSR.SelectedIndex = 1 Or Me.cboTipeSR.SelectedIndex = 2 Then
                'cek apakah do yang dimasukkan sudah ada invoicenya --> untuk yang nota return pajak harus sudah ada
                'invoicenya
                If Me.ObjInv.ApakahSudahAdaInvoice(Me.txtPrimaryDO.Text) <> True Then
                    MsgBox("For a sales return to generate a tax return invoice the delivery order must be invoiced.", _
                    MsgBoxStyle.Information, "Information")
                    Me.cboTipeSR.Focus()
                    Return False
                End If

                'jika sudah ada nomor nota return yang sama maka berikan peringatan.
                Dim DummySR As String = Me.ObjSalesReturn.GetNoSalesReturn(Me.TxtNotaRetur.Text, Me.TxtPrimarySalesReturn.Text)
                If DummySR <> "" Then
                    If MsgBox("The following sales return has a tax return invoice number '" + Me.TxtNotaRetur.Text + "': " + vbCrLf + DummySR + vbCrLf + "Continue data saving process?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If

                'jika Primary account berbeda dengan primary account punya AOP maka di blok 
                Dim PrimaryAOP As Double = -1
                PrimaryAOP = ObjInv.GetPrimaryInvoice(Me.txtInvoiceNo.Text)
                If PrimaryAOP > -1 Then
                    If Me.txtPrimaryAccount.Text <> Me.ObjInv.GetPrimaryAccountAR(PrimaryAOP) Then
                        Dim PrimaryAccountInv As Double = -1
                        Dim AccountInv As String = ""
                        PrimaryAccountInv = Me.ObjInv.GetPrimaryAccountAR(PrimaryAOP)
                        AccountInv = ObjAkun.GetAccountCode(PrimaryAccountInv) & " - " & ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAccountInv)
                        MsgBox("Account receivable of invoice no. " & Me.txtInvoiceNo.Text & " is " & AccountInv & _
                               ", you cannot use different account payable with sales invoice.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'jika PrimarycurrTax berbeda dengan PrimarycurrTax punya AOP maka di blok (akan berpengaruh pada journal)
                    If Me.txtPrimaryCurrTax.Text <> Me.ObjInv.GetPrimaryCurrTax(PrimaryAOP) Then
                        Dim PrimaryCurrTaxInv As Double = -1
                        Dim CurrTaxInv As String = ""
                        PrimaryCurrTaxInv = Me.ObjInv.GetPrimaryCurrTax(PrimaryAOP)
                        CurrTaxInv = Me.ObjCur.GetCurrencyCode(PrimaryCurrTaxInv)
                        MsgBox("Currency tax of invoice no. " & Me.txtInvoiceNo.Text & " is " & CurrTaxInv & _
                               ", you cannot use different Currency tax with sales invoice.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                If Me.chkTax.Checked = True Then
                    If CDec(Me.txtTax.Text) <= 0 Then
                        MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                        'Me.TabControl1.SelectedIndex = 0
                        Me.txtTax.Focus()
                        Return False
                    End If
                    If CDec(Me.txtTax.Text) <> 10 Then
                        If MsgBox("Tax rate <> 10%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            Me.txtTax.Focus()
                            Return False
                        End If
                    End If
                End If

            ElseIf Me.cboTipeSR.SelectedIndex = 0 Then
                'untuk yang revisi invoice, invoice ybs tidak boleh sudah berada di dalam periode closing
                TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod)
                TanggalInvoice = Me.ObjInv.GetTanggalTransaksi_ByPrimaryDO(Me.txtPrimaryDO.Text)
                If TanggalInvoice <> #1/1/1950# Then
                    If TanggalInvoice <= TanggalClosingMax.Date Then
                        MsgBox("The invoice is dated in closed period, you cannot set this sales return as revision invoice.", _
                        MsgBoxStyle.Information, "Information")
                        Me.cboTipeSR.Focus()
                        Return False
                    End If
                End If
            End If

            'untuk quick invoice, tidak bisa revisi invoice
            If Me.ObjInv.ApakahQuickInvoice(Me.ObjInv.GetPrimaryInvoice(Me.txtInvoiceNo.Text)) = True Then
                If Me.cboTipeSR.SelectedIndex = 0 Then
                    MsgBox("This data refers to a quick invoice, must be a return tax invoice.", MsgBoxStyle.Information, "Information")
                    Me.cboTipeSR.SelectedIndex = 1
                End If
            End If

            'Update NPWP 16 Digit
            If cboTipeSR.SelectedIndex = 1 And chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true dan tipe nya tax retur note
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If Me.txtNoNPWP.Text <> "" And Me.dtpDate.Value.Date < ObjSetNPWP.Tglberlaku Then
                    If ObjSetNPWP.ApakahSudah15Digit(Me.txtNoNPWP.Text, True) = False Then
                        Me.txtNoNPWP.Text = ObjCustomer.GetNPWPOld(Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text))
                    End If
                End If

                If ObjSetNPWP.SetNPWP(Me.dtpDate.Value.Date, Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text), Me.txtNoNPWP.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Customer) = False Then
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
        End Try
    End Function

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiSalesReturn
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjSalesReturn.CariDetail(hj.PrimaryTransaksi, Me.dgvItem, Me) = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    '====================================================================================================================================================================

    Private Sub IsiDataStockIn()
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.txtPrimaryStockIn.Text <> "" And Me.txtPrimaryStockIn.Text <> "-1" Then
                Me.TxtStockInNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.txtPrimaryStockIn.Text)
                Dim PrimaryDO As Double = Me.ObjStockInOut.GetPrimaryDO(Me.txtPrimaryStockIn.Text)
                Me.txtPrimaryDO.Text = PrimaryDO
                Me.ObjDO.Isi_DONumber_CustName_CustCode_InvoiceNo_InvoiceDate(Me.txtDONo.Text, _
                Me.txtCustomerCode.Text, Me.TxtCustomerName.Text, Me.txtInvoiceNo.Text, Me.dtpInvoiceDate.Value, _
                PrimaryDO, Me.txtNoNPWP.Text)
                Dim PrimaryCustomer As Double = Me.ObjStockInOut.GetPrimaryCustomer(Me.txtPrimaryStockIn.Text)
                If Me.txtCustomerCode.Text = "" Then
                    'isikan data customer code dan name dari stock in sales return
                    Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(PrimaryCustomer)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(PrimaryCustomer)
                End If
                'Update NPWP 16 Digit
                If Me.txtNoNPWP.Text = "" And PrimaryCustomer > 0 Then
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(PrimaryCustomer)
                End If
                '2025 Core Tax Tahap 2.1
                If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                    Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                End If
            Else
                Me.TxtStockInNo.Text = ""
                Me.txtDONo.Text = ""
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.txtInvoiceNo.Clear()
                Me.dtpInvoiceDate.Value = #1/1/1950#
            End If

            If Me.txtInvoiceNo.Text = "" Then
                Me.txtInvoiceNo.Visible = False
                Me.lblInvoiceNo.Visible = True
                Me.dtpInvoiceDate.Visible = False
                Me.txtInvoiceNo.TabStop = False
                Me.dtpInvoiceDate.TabStop = False
                Me.chkTaxReturn.Enabled = False
                Me.chkTaxReturn.Checked = False
            Else
                Me.txtInvoiceNo.Visible = True
                Me.lblInvoiceNo.Visible = False
                Me.dtpInvoiceDate.Visible = True
                Me.txtInvoiceNo.TabStop = True
                Me.dtpInvoiceDate.TabStop = True
                If Me.TxtRemark.Enabled = True Then
                    Me.chkTaxReturn.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware

        Try
            Me.ColPrimarySalesPerson.Visible = False
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboTipeSR.SelectedIndex = 0
            Me.ObjSalesReturn.QuickSR = False
            Me.ShowIcon = False

            'Me.ColPriceByMultiUnit.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvSOC.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSalesReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()

                            If Me.ObjSalesReturn.ApakahHapusMengakibatkanMinus2(Me.TxtPrimarySalesReturn.Text, False) = False Then
                                'konfirm ke user apakah akan buat data sales order cancellation
                                If MsgBox("Create new sales order cancellation regarding this return?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Dim FrmDia As New FrmDialogCancellationReason
                                    If FrmDia.ShowDialog = Windows.Forms.DialogResult.OK Then
                                        Me.ObjSalesReturn.SimpanSalesOrderCancellation(Me.dgvItem, Me.dgvSOC, FrmDia.TextBox1.Text, _
                                        Me.UserCode, Me.ObjSalesReturn.NoTransaksi, Me)
                                    Else
                                        MsgBox("Sales order generation cancelled by user.", MsgBoxStyle.Information, "Information")
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSalesReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
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
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjSalesReturn.CariDetail(Me.TxtPrimarySalesReturn.Text, Me.dgvItem, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtSalesReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesReturnNo.Text <> "" Then
                'hanya boleh edit yg bukan dari quick sr
                If Me.ObjSalesReturn.ApakahQuickSR(Me.TxtPrimarySalesReturn.Text) = True Then
                    MsgBox("Please use quick sales return form to edit this data", MsgBoxStyle.Information, "Information")
                Else
                    If Me.ObjSalesReturn.ApakahBolehEdit(Me.TxtPrimarySalesReturn.Text, Me.UserCode, True) = True Then
                        Me.Baru = False
                        Me.AktifkanDetail()
                        Me.dtpDate.Focus()
                        'nilai min juga harus diatas tanggal pengkuncian data terakhir
                        'msgbox("belum")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtSalesReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesReturnNo.Text <> "" Then
                'hanya boleh hapus yg bukan quick sr
                If Me.ObjSalesReturn.ApakahQuickSR(Me.TxtPrimarySalesReturn.Text) = True Then
                    MsgBox("Please use quick sales return form to delete this data", MsgBoxStyle.Information, "Information")
                Else
                    If Me.ObjSalesReturn.ApakahBolehHapus(Me.TxtPrimarySalesReturn.Text, Me.UserCode, True) = False Then
                        Me.btnFind.Focus()
                    Else
                        hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSalesReturn.Hapus(Me.TxtPrimarySalesReturn.Text, Me.UserCode) = True Then
                                Me.KosongkanField()
                                Me.btnNew.Focus()
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSalesReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()

                                'konfirm ke user apakah akan buat data sales order cancellation
                                If MsgBox("Create new sales order cancellation regarding this return?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Dim FrmDia As New FrmDialogCancellationReason
                                    If FrmDia.ShowDialog = Windows.Forms.DialogResult.OK Then
                                        Me.ObjSalesReturn.SimpanSalesOrderCancellation(Me.dgvItem, Me.dgvSOC, FrmDia.TextBox1.Text, _
                                        Me.UserCode, Me.ObjSalesReturn.NoTransaksi, Me)
                                    Else
                                        MsgBox("Sales order generation cancelled by user.", MsgBoxStyle.Information, "Information")
                                    End If
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSalesReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim IsMultiUnit As Boolean = False

        Try
            Dim AdaSec As Boolean = False
            Dim AdaSo As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaBatchNo As Boolean = False

            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                row.Cells(Me.ColQtyMaster.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyMaster.Index).Value)
                row.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty.Index).Value)
                row.Cells(Me.ColQtyPerMaster.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyPerMaster.Index).Value)
                row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
                If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                    AdaSec = True
                End If
                If row.Cells(Me.ColPrimarySO.Index).Value <> -1 Then
                    AdaSo = True
                End If
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    If row.Cells(Me.ColExpiryDate.Index).Value <> #1/1/1950# Then
                        AdaExpiry = True
                        row.Cells(Me.ColExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy")
                    Else
                        row.Cells(Me.ColExpiryDate.Index).Value = "-"
                    End If
                End If
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    AdaBatchNo = True
                End If
            Next

            If AdaSec = True Then
                Me.ColSecQty.Visible = True
                Me.ColSecUnit.Visible = True
            Else
                Me.ColSecQty.Visible = False
                Me.ColSecUnit.Visible = False
            End If
            If AdaSo = True Then
                Me.ColSONO.Visible = True
            Else
                Me.ColSONO.Visible = False
            End If
            If AdaExpiry = True Then
                Me.ColExpiryDate.Visible = True
            Else
                Me.ColExpiryDate.Visible = False
            End If
            If AdaBatchNo = True Then
                Me.ColBatchNo.Visible = True
            Else
                Me.ColBatchNo.Visible = False
            End If

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '2. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Me.dgvItem.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        End If
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        If Me.dgvItem.Rows.Count = 0 Then
            Me.dgvItem.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            If Me.TxtRemark.Enabled = True Then
                'If Me.BolehEdit = True Then
                Me.TxtStockInNo.Enabled = True
                Me.btnFindStockIn.Enabled = True
                'End If
            End If
        End If
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            'format tampilan
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'tidak perlu cek akibatnya pada sales order balance
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
            'tandai apakah pada baris yang sama no SO-nya sama
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub TxtStockOutNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.GotFocus
        TxtStockInNo.BackColor = BackColorAll
    End Sub

    Private Sub TxtStockOutNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.LostFocus
        TxtStockInNo.BackColor = Me.TxtPrimarySalesReturn.BackColor
    End Sub

    Private Sub TxtReasonOfReturn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.GotFocus
        TxtReasonOfReturn.BackColor = BackColorAll
    End Sub

    Private Sub TxtReasonOfReturn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.LostFocus
        TxtReasonOfReturn.BackColor = Me.TxtPrimarySalesReturn.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.TxtPrimarySalesReturn.BackColor
    End Sub

    'Khusus untuk CMSI, bila sudah memiliki invoice maka mata uang (yang digunakan untuk perhitungan total
    'dan sub total), exchange rate to IDR, dan tax rate digunakan data dari invoice dari mana SR ybs berasal.
    'Lebih jauh lagi bila buat nota retur pajak, maka 
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmCetakTransaksiSalesReturn

        Try
            If Me.TxtSalesReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtSalesReturnNo.Text <> "" Then
                If Me.cboTipeSR.SelectedIndex = 1 Or Me.cboTipeSR.SelectedIndex = 2 Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    Dim obj As New ClsTransaksiSalesReturn
                    Dim ObjSO As New ClsTransaksiSalesOrder

                    obj.CetakTaxReturn(Me.TxtPrimarySalesReturn.Text, ij.crv1, Me.DS, 0, Me.ObjCD)

                    obj = Nothing
                    ObjSO = Nothing

                    Windows.Forms.Cursor.Current = Cursors.Arrow
                    ij.ShowDialog()
                Else
                    MsgBox("This feature only available on sales return with tax return", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Function GetExchangeRate(ByVal vTBL As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateDataTable, _
    ByVal vTipePerhitungan As ClsTransaksiSalesInvoice.EnumCalculatesBy, ByVal vPrimarySO As Double, _
    ByVal vPrimaryDO As Double, ByVal vPrimaryInv As Double) As Double
        Dim Filter As String = ""
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateRow

        Try
            If vTipePerhitungan = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                If vTBL.Rows.Count > 0 Then
                    row = vTBL.Rows(0)
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimaryDO.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimarySO.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanDGV()
        Dim row As DataGridViewRow
        Dim Filter As String = ""
        Dim TipeExcRate As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice
        Dim DTA_Inv As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter
        Dim TBL_Inv As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateDataTable
        Dim VPrimaryInvoice As Double = -1
        Dim Rounding As Integer = 0
        Dim RoundingInvoice As Integer = 0
        Dim RowDgv As DataGridViewRow
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Try
            If Me.txtPrimaryStockIn.Text <> "" And Me.txtPrimaryStockIn.Text <> "-1" Then
                Me.ObjStockInOut.CariDetail(Me.txtPrimaryStockIn.Text, Me.dgvItemStockIn, TblItemCustom)
                Me.dgvItem.Rows.Clear()
                For Each row In Me.dgvItemStockIn.Rows
                    'tambahkan ke dalam dgv
                    Dim Index As Integer = Me.GetIndex(row.Cells(Me.SiR_ColPrimaryCodeNo.Index).Value, _
                    row.Cells(Me.SiR_ColPrimarySO.Index).Value, row.Cells(Me.SiR_ColBatch.Index).Value, _
                    row.Cells(Me.SiR_ColSecondaryQty.Index).Value)
                    If Index = -1 Then
                        With row
                            'tambahkan baris baru
                            Me.dgvItem.Rows.Add()
                            RowDgv = dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                            RowDgv.Cells(Me.ColNo.Index).Value = ""
                            RowDgv.Cells(Me.ColCodeNo.Index).Value = .Cells(SiR_ColCodeNo.Index).Value
                            RowDgv.Cells(Me.ColDescription.Index).Value = .Cells(SiR_ColDescription.Index).Value
                            RowDgv.Cells(Me.ColSONO.Index).Value = .Cells(SiR_ColPONoDONo.Index).Value
                            RowDgv.Cells(Me.ColBatchNo.Index).Value = .Cells(SiR_ColBatch.Index).Value
                            RowDgv.Cells(Me.ColExpiryDate.Index).Value = .Cells(SiR_ColExpiredDate.Index).Value
                            RowDgv.Cells(Me.ColQtyMaster.Index).Value = 0
                            RowDgv.Cells(Me.ColQtyPerMaster.Index).Value = 0
                            RowDgv.Cells(Me.ColMasterUnits.Index).Value = ""
                            RowDgv.Cells(Me.ColQty.Index).Value = 0
                            RowDgv.Cells(Me.ColTotalQty.Index).Value = .Cells(SiR_ColQty.Index).Value
                            RowDgv.Cells(Me.ColUnit.Index).Value = .Cells(SiR_ColUnit.Index).Value
                            RowDgv.Cells(Me.ColSecQty.Index).Value = .Cells(SiR_ColSecondaryQty.Index).Value
                            RowDgv.Cells(Me.ColSecUnit.Index).Value = .Cells(SiR_ColSecondaryUnit.Index).Value
                            RowDgv.Cells(Me.ColRemarks.Index).Value = .Cells(SiR_ColRemark.Index).Value
                            RowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = .Cells(SiR_ColPrimaryCodeNo.Index).Value
                            RowDgv.Cells(Me.ColPrimarySO.Index).Value = .Cells(SiR_ColPrimarySO.Index).Value
                            RowDgv.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                            RowDgv.Cells(Me.ColPrimaryUnit.Index).Value = .Cells(SiR_ColPrimaryUnits.Index).Value
                            RowDgv.Cells(Me.ColPrimarySecUnit.Index).Value = .Cells(SiR_ColPrimarySecondaryUnits.Index).Value
                            RowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = .Cells(SiR_ColQtyMulti.Index).Value
                            RowDgv.Cells(Me.ColUnitsMulti.Index).Value = .Cells(SiR_ColUnitsCodeMulti.Index).Value
                            RowDgv.Cells(Me.ColRatioDesc.Index).Value = .Cells(SiR_ColRatioDesc.Index).Value
                            RowDgv.Cells(Me.ColRatioDefault.Index).Value = .Cells(SiR_ColRatioDefault.Index).Value
                            RowDgv.Cells(Me.ColRatioReference.Index).Value = .Cells(SiR_ColRatioReference.Index).Value
                            RowDgv.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Cells(SiR_ColPrimaryUnitsMulti.Index).Value

                            RowDgv.Cells(Me.ColPrimarySalesPerson.Index).Value = -1
                            Dim PrimarySalesInvoice As Double = ObjInv.GePrimarySalesPerson(RowDgv.Cells(Me.ColPrimarySO.Index).Value, Me.txtPrimaryDO.Text)
                            If PrimarySalesInvoice <> -1 Then
                                RowDgv.Cells(Me.ColPrimarySalesPerson.Index).Value = PrimarySalesInvoice
                            Else
                                RowDgv.Cells(Me.ColPrimarySalesPerson.Index).Value = ObjSO.GetPrimarySalesPerson(RowDgv.Cells(Me.ColPrimarySO.Index).Value)
                            End If

                            'Me.dgvItem.Rows.Add("", .Cells(Me.ColCodeNoStockIn.Index).Value, _
                            '.Cells(Me.ColDescStockIn.Index).Value, .Cells(Me.ColPONoStockIn.Index).Value, _
                            '.Cells(Me.ColBatchStockIn.Index).Value, .Cells(Me.ColExpiredDateStockIn.Index).Value, _
                            '0, 0, "", 0, .Cells(Me.ColQtyStockIn.Index).Value, .Cells(Me.ColUnitsStockIn.Index).Value, _
                            '.Cells(Me.ColSecondaryQtyStockIn.Index).Value, .Cells(Me.ColSecondaryUnitStockIn.Index).Value, _
                            '.Cells(Me.ColRemarkStockIn.Index).Value, .Cells(Me.ColPrimaryCodeStockIn.Index).Value, _
                            '.Cells(Me.ColPrimaryPOnoStockIn.Index).Value, -1, .Cells(Me.ColPrimaryUnitsStockIn.Index).Value, _
                            '.Cells(Me.ColPrimarySecondaryUnitsStockIn.Index).Value)
                            'set ulang nomor
                            Me.ObjCD.SetNumber(Me.dgvItem)
                            Me.FormatDGV()
                        End With
                    Else
                        'ubah data qty saja
                        Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, Index).Value = ObjCD.DigitGrouping(CDbl(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, Index).Value) _
                            + CDbl(row.Cells(Me.SiR_ColQtyMulti.Index).Value), ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, Index).Value, _
                            Me.dgvItem(Me.ColPrimaryUnitsMulti.Index, Index).Value))
                        Me.dgvItem.Item(Me.ColTotalQty.Index, Index).Value = ObjCD.DigitGrouping(CDbl(Me.dgvItem.Item(Me.ColTotalQty.Index, Index).Value) _
                            + CDbl(row.Cells(Me.SiR_ColQty.Index).Value), ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, Index).Value))
                    End If
                Next

                'isikan data harga pada item
                Me.IsikanDataHarga()

                'hitung total
                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Inv.Dispose()
            TBL_Inv.Dispose()
        End Try
    End Sub

    Private Sub IsikanDataHarga()
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.DataTable7TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.DataTable7DataTable
        Dim rowTBL As dseTransaksiSalesReturn.DataTable7Row
        Dim Filter As String = ""
        Dim TipeExcRate As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice
        Dim DTA_Inv As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter
        Dim TBL_Inv As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateDataTable
        Dim VPrimaryInvoice As Double = -1
        Dim Rounding As Integer = 0
        Dim RoundingInvoice As Integer = 0
        Dim vPrimaryCurrency, vTaxExchangeRate, vTaxPercent, vPrimaryCurrTax, vPrimaryAccount As Double
        Dim vUseTax As Boolean

        Try
            'isikan data harga pada item
            DTA.Fill(TBL, Me.txtPrimaryDO.Text)
            VPrimaryInvoice = Me.ObjInv.GetPrimaryInvoice_ByPrimaryDO(Me.txtPrimaryDO.Text)
            DTA_Inv.FillBy_PrimarySalesInvoice(TBL_Inv, VPrimaryInvoice)

            'isikan data currency, tipe pajak, tax percent, tax exchannge rate
            If VPrimaryInvoice = -1 Then
                Me.TxtPrimaryCurrency.Text = -1
                Me.TxtCurrency.Text = ""
                Me.chkTax.Checked = False
                Me.txtTax.Text = 0
                Me.txtTaxAmount.Text = 0
                Me.txtTaxExchangeRate.Text = 1
                Me.txtSubTotalAfterDP.Text = 0
                Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = 0
                RoundingInvoice = 0
                Me.txtPrimaryCurrTax.Text = -1
                Me.txtPrimaryAccount.Text = -1
                Me.cboCurrTax.DataSource = Nothing
                Me.cboAccount.DataSource = Nothing
                Me.cboPrimaryCurrTax.DataSource = Nothing
                Me.cboPrimaryCurrTax.DataSource = Nothing
            Else
                Me.ObjInv.Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(VPrimaryInvoice, vUseTax, vTaxPercent, _
                vTaxExchangeRate, vPrimaryCurrency)
                Me.TxtPrimaryCurrency.Text = vPrimaryCurrency
                Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(vPrimaryCurrency)
                RoundingInvoice = Me.ObjCur.GetDefaultPrecission(vPrimaryCurrency)
                Me.chkTax.Checked = vUseTax
                Me.txtTax.Text = vTaxPercent
                Me.txtTaxAmount.Text = 0
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(vTaxExchangeRate, 2)
                Else
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.txtSubTotalAfterDP.Text = 0
                If Me.cboTipeSR.SelectedIndex = 0 Then
                    Me.TxtDutystamp.Text = 0
                End If
                Me.TxtTotal.Text = 0
                vPrimaryCurrTax = Me.ObjInv.GetPrimaryCurrTax(VPrimaryInvoice)
                vPrimaryAccount = Me.ObjInv.GetPrimaryAccountAR(VPrimaryInvoice)

                Me.IsiCBOArAkunByTipePerkiraan1()
                If vPrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = vPrimaryAccount
                    Me.cboPrimaryAccount.Text = vPrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOCurrencyTax()
                If vPrimaryCurrTax > -1 Then
                    Me.txtPrimaryCurrTax.Text = vPrimaryCurrTax
                    Me.cboPrimaryCurrTax.Text = vPrimaryCurrTax
                Else
                    Me.cboCurrTax.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                End If
            End If

            For Each row In Me.dgvItem.Rows
                With row
                    If TBL.Rows.Count <> 0 Then
                        Filter = "[Primary Product] =" + row.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString _
                        + " AND [Primary SO] = " + row.Cells(Me.ColPrimarySO.Index).Value.ToString + " AND [Batch Number] = '" + row.Cells(Me.ColBatchNo.Index).Value + "'"
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count > 0 Then
                            rowTBL = TBL.DefaultView(0).Row
                            .Cells(Me.ColPrimaryCurrency.Index).Value = rowTBL.Primary_Currency
                            Rounding = Me.ObjCur.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                            .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Unit_Price, Rounding))
                            .Cells(Me.ColDisc.Index).Value = ObjCD.DigitGrouping(rowTBL.Discount, 2)
                            .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Nett_Price, Rounding))
                            .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(TBL_Inv, Me.ObjInv.GetTipePerhitungan(VPrimaryInvoice), _
                            row.Cells(Me.ColPrimarySO.Index).Value, Me.txtPrimaryDO.Text, VPrimaryInvoice))
                            .Cells(Me.ColCurrency.Index).Value = rowTBL.Currency_Code
                            .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowTBL.Disc_Amount, Rounding)
                            .Cells(Me.ColDiscFromPerct.Index).Value = rowTBL.DiscFromPerc
                            .Cells(Me.ColPriceByMultiUnit.Index).Value = rowTBL.price_by_multi_unit
                            If rowTBL.price_by_multi_unit = 1 Then
                                .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Cells(Me.ColUnitsMulti.Index).Value
                                .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColNettPrice.Index).Value * .Cells(Me.ColTotalQtyMulti.Index).Value, Rounding)
                            Else
                                .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Cells(Me.ColUnit.Index).Value
                                .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColNettPrice.Index).Value * .Cells(Me.ColTotalQty.Index).Value, Rounding)
                            End If
                        Else
                            .Cells(Me.ColUnitPrice.Index).Value = 0
                            .Cells(Me.ColDisc.Index).Value = 0
                            .Cells(Me.ColNettPrice.Index).Value = 0
                            .Cells(Me.ColExchangeRate.Index).Value = 1
                            .Cells(Me.ColCurrency.Index).Value = ""
                            .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                            .Cells(Me.ColAmount.Index).Value = 0
                            .Cells(Me.ColDiscAmount.Index).Value = 0
                            .Cells(Me.ColDiscFromPerct.Index).Value = 1
                            .Cells(Me.ColPriceByMultiUnit.Index).Value = "1"
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = ""
                        End If
                    Else
                        .Cells(Me.ColUnitPrice.Index).Value = 0
                        .Cells(Me.ColDisc.Index).Value = 0
                        .Cells(Me.ColNettPrice.Index).Value = 0
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColCurrency.Index).Value = ""
                        .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                        .Cells(Me.ColAmount.Index).Value = 0
                        .Cells(Me.ColDiscAmount.Index).Value = 0
                        .Cells(Me.ColDiscFromPerct.Index).Value = 1
                        .Cells(Me.ColPriceByMultiUnit.Index).Value = "1"
                        .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = ""
                    End If
                End With
            Next

            'hitung total
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_Inv.Dispose()
            TBL_Inv.Dispose()
        End Try
    End Sub


    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal LotNo As String, ByVal SecQty As Double) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If CDbl(row.Cells(Me.ColSecQty.Index).Value) = SecQty And row.Cells(Me.ColPrimaryCodeNo.Index).Value _
                = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                LotNo = row.Cells(Me.ColBatchNo.Index).Value Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockInNo.KeyDown
        Dim i As New FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            If e.KeyValue = 13 Then
                If Me.ObjStockInOut.ApakahSudahAda(Me.TxtStockInNo.Text) = True Then
                    Me.txtPrimaryStockIn.Text = Me.ObjStockInOut.GetPrimary(Me.TxtStockInNo.Text)
                    'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                    Dim NoTransaksiLain As String = Me.ObjSalesReturn.GetNoTransaksiByStockIn(Me.txtPrimaryStockIn.Text, Me.TxtPrimarySalesReturn.Text)
                    If NoTransaksiLain <> "" Then
                        MsgBox("This stock in sales return referred by sales return " _
                        + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Me.IsikanDataHarga()
                        Exit Sub
                    End If

                    Me.IsiDataStockIn()
                    Me.IsikanDGV()
                    Me.TxtNotaRetur.Focus()
                Else
                    'ada 2 stock in dengan kode yang sama atau tidak ada stock in yang dimaksud
                    If Me.TxtStockInNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockInNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiStockInOut > 0 Then
                        'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                        Dim NoTransaksiLain As String = Me.ObjSalesReturn.GetNoTransaksiByStockIn(i.PrimaryTransaksiStockInOut, Me.TxtPrimarySalesReturn.Text)
                        If NoTransaksiLain <> "" Then
                            MsgBox("This stock in sales return referred by sales return " _
                            + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                            Me.TxtStockInNo.Clear()
                            Me.txtPrimaryStockIn.Text = "-1"
                            Me.IsiDataStockIn()
                            Me.IsikanDataHarga()
                            Exit Sub
                        End If

                        Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                        Me.IsiDataStockIn()
                        Me.IsikanDGV()
                        Me.TxtNotaRetur.Focus()
                    Else
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Me.dgvItem.Rows.Clear()
                        Me.IsikanDataHarga()
                        Me.TxtStockInNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.Leave
        Dim i As New FrmCariTransaksiStockInOut
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            If Me.TxtStockInNo.Text <> "" Then
                If Me.ObjStockInOut.ApakahSudahAda(Me.TxtStockInNo.Text) = True Then
                    Me.txtPrimaryStockIn.Text = Me.ObjStockInOut.GetPrimary(Me.TxtStockInNo.Text)
                    Me.IsiDataStockIn()
                    Me.IsikanDGV()
                Else
                    If Me.TxtStockInNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockInNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiStockInOut > 0 Then
                        'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                        Dim NoTransaksiLain As String = Me.ObjSalesReturn.GetNoTransaksiByStockIn(i.PrimaryTransaksiStockInOut, Me.TxtPrimarySalesReturn.Text)
                        If NoTransaksiLain <> "" And NoTransaksiLain <> Me.TxtSalesReturnNo.Text Then
                            MsgBox("This stock in sales return referred by sales return " _
                            + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                            Me.TxtStockInNo.Clear()
                            Me.txtPrimaryStockIn.Text = "-1"
                            Me.IsiDataStockIn()
                            Me.IsikanDataHarga()
                            Exit Sub
                        End If

                        Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                        Me.IsiDataStockIn()
                        Me.IsikanDGV()
                        Me.TxtNotaRetur.Focus()
                    Else
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Me.TxtStockInNo.Focus()
                        Me.dgvItem.Rows.Clear()
                        Me.IsikanDataHarga()
                    End If
                End If
            Else
                Me.txtPrimaryStockIn.Text = "-1"
                Me.IsiDataStockIn()
                Me.dgvItem.Rows.Clear()
                Me.IsikanDataHarga()
            End If
            Me.SetKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindStockIn.Click
        Dim i As New FrmCariTransaksiStockInOut

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
            If Me.TxtStockInNo.Text <> "" Then
                i.txtCari.Text = Me.TxtStockInNo.Text
            End If
            i.ShowDialog()

            If i.PrimaryTransaksiStockInOut > 0 Then
                'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                Dim NoTransaksiLain As String = Me.ObjSalesReturn.GetNoTransaksiByStockIn(i.PrimaryTransaksiStockInOut, Me.TxtPrimarySalesReturn.Text)
                If NoTransaksiLain <> "" Then
                    MsgBox("This stock in sales return referred by sales return " _
                    + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                    Me.TxtStockInNo.Clear()
                    Me.txtPrimaryStockIn.Text = "-1"
                    Me.IsiDataStockIn()
                    Me.IsikanDataHarga()
                    Exit Sub
                End If

                Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                Me.IsiDataStockIn()
                Me.IsikanDGV()
                Me.TxtNotaRetur.Focus()
            Else
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockIn.Text = "-1"
                Me.IsiDataStockIn()
                Me.TxtStockInNo.Focus()
                Me.IsikanDataHarga()
                Me.dgvItem.Rows.Clear()
            End If
            Me.SetKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.Baru = False Then
                    If Me.ObjSalesReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjSalesReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolom()

        Try
            If Me.cboTipeSR.SelectedIndex = 0 Then 
                Me.ColNettPrice.Visible = False
                Me.ColDisc.Visible = False
                Me.ColDiscAmount.Visible = False
                Me.ColUnitPrice.Visible = False
                Me.ColCurrency.Visible = False
                Me.ColAmount.Visible = False
                Me.ColPriceByMultiUnit_Str.Visible = False
                Me.ColExchangeRate.Visible = False
            Else
                Me.ColNettPrice.Visible = True
                Me.ColDisc.Visible = True
                Me.ColDiscAmount.Visible = True
                Me.ColUnitPrice.Visible = True
                Me.ColCurrency.Visible = True
                Me.ColAmount.Visible = True
                Me.ColPriceByMultiUnit_Str.Visible = True
                Me.ColExchangeRate.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTipeSR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipeSR.SelectedIndexChanged
        Try
            Me.ChkInvoiceCancellation.Checked = False
            If Me.cboTipeSR.SelectedIndex = 0 Then
                Me.chkTaxReturn.Enabled = False
                Me.TxtNotaRetur.ReadOnly = True
                Me.TxtNotaRetur.BackColor = Me.TxtReadOnly.BackColor
                Me.chkTaxReturn.Checked = False
                Me.TxtNotaRetur.Text = ""
                Me.ColNettPrice.Visible = False
                Me.ColDisc.Visible = False
                Me.ColDiscAmount.Visible = False
                Me.ColUnitPrice.Visible = False
                Me.ColCurrency.Visible = False
                Me.ColAmount.Visible = False
                Me.ColPriceByMultiUnit_Str.Visible = False
                Me.TxtDutystamp.ReadOnly = True
                Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                Me.ColExchangeRate.Visible = False

                Me.cboAccount.DataSource = Nothing
                Me.cboPrimaryAccount.DataSource = Nothing
                Me.txtPrimaryAccount.Text = -1
                Me.cboPrimaryAccount.Enabled = False
                Me.cboAccount.Enabled = False

                Me.cboCurrTax.Enabled = False
                Me.cboPrimaryCurrTax.Enabled = False
                Me.cboCurrTax.DataSource = Nothing
                Me.cboPrimaryCurrTax.DataSource = Nothing
                Me.txtPrimaryCurrTax.Text = -1
            Else
                Me.chkTaxReturn.Enabled = True
                Me.TxtNotaRetur.ReadOnly = False
                Me.TxtNotaRetur.BackColor = Me.txtPrimaryStockIn.BackColor
                Me.ColNettPrice.Visible = True
                Me.ColDisc.Visible = True
                Me.ColDiscAmount.Visible = True
                Me.ColUnitPrice.Visible = True
                Me.ColCurrency.Visible = True
                Me.ColAmount.Visible = True
                Me.ColPriceByMultiUnit_Str.Visible = True
                Me.TxtDutystamp.ReadOnly = False
                Me.TxtDutystamp.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.ColExchangeRate.Visible = True

                Me.cboCurrTax.Enabled = True
                Me.cboPrimaryCurrTax.Enabled = True
                Me.cboPrimaryAccount.Enabled = True
                Me.cboAccount.Enabled = True
                'Me.txtPrimaryAccount.Visible = True
                '  Me.TxtPrimaryCurrency.Text = 1
                Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                Me.IsiCBOArAkunByTipePerkiraan1()
                If IsNumeric(txtPrimaryAccount.Text) AndAlso Me.txtPrimaryAccount.Text > -1 Then
                    Me.txtPrimaryAccount.Text = Me.txtPrimaryAccount.Text
                    Me.cboPrimaryAccount.Text = Me.txtPrimaryAccount.Text
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOCurrencyTax()
                If IsNumeric(txtPrimaryCurrTax.Text) AndAlso txtPrimaryCurrTax.Text > -1 Then
                    Me.txtPrimaryCurrTax.Text = txtPrimaryCurrTax.Text
                    Me.cboPrimaryCurrTax.Text = txtPrimaryCurrTax.Text
                Else
                    Me.cboCurrTax.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                End If

                If Me.cboTipeSR.SelectedIndex = 2 Then
                    Me.ChkInvoiceCancellation.Checked = True
                End If

            End If
            If Me.TxtRemark.Enabled = True Then
                Me.IsikanDataHarga()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutystamp.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutystamp.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

  
    Private Sub TxtDutystamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = False Then
                    Me.TxtDutystamp.BackColor = Me.BackColorAll
                Else
                    Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = False Then
                    Me.TxtDutystamp.BackColor = Me.TxtPrimaryCurrency.BackColor
                Else
                    Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentTerm.ReadOnly = True Then
                    Me.TxtPaymentTerm.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPaymentTerm.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPaymentTerm, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.Leave
        Try
            Me.TxtPaymentTerm.Text = Me.ObjCD.DigitGrouping(Me.TxtPaymentTerm.Text, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentTerm.ReadOnly = True Then
                    Me.TxtPaymentTerm.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPaymentTerm.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNotaRetur_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotaRetur.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNotaRetur.Enabled = True Then
                    If Me.TxtNotaRetur.ReadOnly = True Then
                        Me.TxtNotaRetur.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.TxtNotaRetur.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNotaRetur_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotaRetur.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNotaRetur.Enabled = True Then
                    If Me.TxtNotaRetur.ReadOnly = True Then
                        Me.TxtNotaRetur.BackColor = Me.TxtPrimaryCurrency.BackColor
                    Else
                        Me.TxtNotaRetur.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjSalesReturn.CariDetail(vPrimary, Me.dgvItem, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SalesReturn_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimarySISR As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryDO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimarySO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)
        Dim TblTemp As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.Baru = False
            Me.ObjSalesReturn.ResetObject(Me.dgvItem)
            Me.ObjSalesReturn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me)

            Me.ObjSalesReturn.PrimaryDO = ObjCutOff_Core.GetPrimaryUpdate(Me.ObjSalesReturn.PrimaryDO, vTblPrimaryDO)
            Me.ObjSalesReturn.PrimaryStockInSalesReturn = ObjCutOff_Core.GetPrimaryUpdate(Me.ObjSalesReturn.PrimaryStockInSalesReturn, vTblPrimarySISR)
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimarySO.Index).Value) = False And rowD.Cells(Me.ColPrimarySO.Index).Value <> "-1" Then
                    TblTemp.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value)
                    If TblTemp.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), vTblPrimarySO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), PrimaryUpdate, TblTemp)
                        rowD.Cells(Me.ColPrimarySO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimarySO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), TblTemp)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SalesReturn_Save() As Double
        Try
            Me.Baru = True
            Me.ObjSalesReturn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me, False, True)
            Return Me.ObjSalesReturn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class