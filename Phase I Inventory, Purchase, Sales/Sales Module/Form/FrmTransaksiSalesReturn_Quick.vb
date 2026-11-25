Option Explicit On
Option Strict Off

Imports Core_Module
Imports Inventory_Module

Public Class FrmTransaksiSalesReturn_Quick
    Dim mTblMultiUnit As New DseTransaksiSO.tblmaster_product_item_multiunitDataTable
    Dim tblCboUnit As New DseTransaksiSO.CboUnitDataTable
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo

    '-balikin lagi
    Dim objCustom As New CustomSMESS.ClsTransaksiSalesReturn
    Public TblDummyCustom As New CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable
    Public TblCustomInOutDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable  'new

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
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    'Dim ObjStockInOutItem As New Inventory_Module.clsTransaksiStockInOutItem
    Dim ObjItemBundling As New Inventory_Module.clsTransaksiStockBundlingItem
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson

    Dim Baru As Boolean
    Dim buffer As String = ""
    Dim DS As New dseTransaksiSalesReturn

    'berkaitan dengan import quick sales return
    Dim vStartDate, vEndDate As Date
    Dim tblOldSalesReturn As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturnDataTable
    Dim tblOldSalesReturnItem As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item1DataTable
    Dim tblOldSalesReturnItem2 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item1DataTable
    Dim tblCheckSR As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn1DataTable
    Dim ErrMsg As String = ""
    Dim LanjutTransfer As Boolean = False
    Dim SedangKonfirmasi As Boolean = False
    Dim FrmErr As New FrmTransferBeginingBalance_Error
    Dim listOfTransfered As New List(Of String)
    Dim listOfSkiped As New List(Of String)
    Dim listOfFailed As New List(Of String)
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

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccount, Me.txtPrimaryCurrency.Text)

            Me.cboAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboPrimaryAccount.DisplayMember = "Primary"
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text


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


            objMulti.IsiTBLCurrTax(Me.TxtPrimaryCurrency.Text, Me.tblCurrTax)



            If Me.tblCurrTax.Rows.Count > 0 Then
                Me.cboCurrTax.DataSource = Me.tblCurrTax.DefaultView
                Me.cboCurrTax.DisplayMember = "Currency Code"
                Me.cboPrimaryCurrTax.DataSource = Me.tblCurrTax.DefaultView
                Me.cboPrimaryCurrTax.DisplayMember = "Primary"
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
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
        Dim SubTotal, TaxAmount, Total As Decimal
        Dim Rounding As Integer = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecCur As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku And Me.TxtPrimaryCurrency.Text = "1" Then
                Me.HitungTotal_MiddlePoint()
                Exit Sub
            End If
            DefPrecCur = Me.ObjCur.GetDecimal_DefaultCurrency
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrecCur
            End If


            If Me.cboTipeSR.SelectedIndex = 1 Then
                'hitung subtotal
                SubTotal = 0
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQty.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    End If
                Next
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))



                'hitung pajak
                If Me.TxtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                Else
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                              CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    Else

                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                               CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    End If
                End If

                'dutystamp
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DefPrecTax)
            Else
                Me.txtSubTotalAfterDP.Text = 0
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.txtTaxExchangeRate.Text = 1
                Me.txtTax.Text = 0
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
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
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
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub HitungTotal_MiddlePoint()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal, TaxAmount, Total As Decimal
        Dim Rounding As Integer = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecCur As Integer = 0

        Try
            DefPrecCur = Me.ObjCur.GetDecimal_DefaultCurrency
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrecCur
            End If


            If Me.cboTipeSR.SelectedIndex = 1 Then
                'hitung subtotal
                SubTotal = 0
                For Each rowdgv In Me.dgvItem.Rows
                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQty.Index).Value), Rounding, MidpointRounding.AwayFromZero)
                    End If
                Next
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))



                'hitung pajak
                If Me.TxtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                Else
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                              CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    Else

                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                               CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    End If
                End If

                'dutystamp
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DefPrecTax)
            Else
                Me.txtSubTotalAfterDP.Text = 0
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.txtTaxExchangeRate.Text = 1
                Me.txtTax.Text = 0
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
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
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
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
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
        Dim UseTaxSetting As String = ""
        Dim TaxPercentSetting As String = ""

        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.ObjSalesReturn.ResetObject(Me.dgvItem)
            Me.TxtSalesReturnNo.Text = Me.ObjSalesReturn.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            'default tipe sales return adalah nota retur pajak
            Me.cboTipeSR.SelectedIndex = 1

            Me.TxtPrimaryCurrency.Text = 1
            Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.chkTax.Checked = True
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If

            Me.SetTaxRate(True)
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
            Me.TxtPrimarySalesP.Text = "-1"
            Me.txtPrimaryDO.Text = "-1"
            Me.TxtPrimaryCustomer.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.TxtPaymentTerm.Text = "0"
            Me.txtTax.Text = "0.00"
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping("0", CInt(Me.ObjCur.GetDecimal_DefaultCurrency))
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", CInt(Me.ObjCur.GetDecimal_DefaultCurrency))

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
            Me.dgvItem.Rows.Clear()
            Me.KosongkanItem()

            Me.TblDummyCustom.Rows.Clear()
            Me.TblCustomInOutDummy.Rows.Clear()
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
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.chkTaxReturn.Enabled = True

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True

            'Me.dtpInvoiceDate.Enabled = False
            Me.chkTax.Enabled = True

            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            Else
                Me.TxtCurrency.ReadOnly = False
                Me.TxtCurrency.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.TxtCurrency.TabStop = True
                Me.btnFindCurrency.Enabled = True
            End If

            Me.cboAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If


            Me.llbImport.Enabled = False
            Me.Atur1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.chkTaxReturn.Enabled = False
            chkTax.Enabled = False
            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimarySalesReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimarySalesReturn.BackColor, False)

            Me.llbImport.Enabled = True
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
                Me.TxtPrimarySalesP.Text = .PrimarySalesPerson
                Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(Me.TxtPrimarySalesP.Text)
                Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)

                Me.TxtReasonOfReturn.Text = .ReturnReason
                Me.txtPrimaryDO.Text = .PrimaryDO
                Me.chkTaxReturn.Checked = .CetakNotaReturn
                Me.TxtNotaRetur.Text = .NoNotaRetur
                If .TipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak Then
                    Me.cboTipeSR.SelectedIndex = 1
                Else
                    Me.cboTipeSR.SelectedIndex = 0
                End If
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency
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
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                    Me.lblCurrTax.Text = Me.ObjCur.GetCurrencyCode(1)
                End If

                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Rounding))
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                Me.chkTax.Checked = .UseTax
                Me.txtTax.Text = ObjCD.DigitGrouping(.TaxPercent, 2)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(.Dutystamp, Me.ObjCur.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Rounding))
                Me.TxtPaymentTerm.Text = Me.ObjCD.DigitGrouping(.PaymentTerm, 0)

                'Update 16 Digit NPWP
                Me.txtNoNPWP.Text = .NPWPNo

                Me.TxtTaxInvoiceNumber.Text = .TaxInvoiceNoQuickSR
                Me.dtpInvoiceDate.Value = .InvoiceDateQuickSR
                Me.txtInvoiceNo.Text = .InvoiceNoQuickSR
                Me.TxtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomer.Text)
            End With
            Me.FormatDGV()
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
                .PrimarySalesPerson = Me.TxtPrimarySalesP.Text
                .PrimaryStockInSalesReturn = Me.txtPrimaryStockIn.Text
                .Remarks = Me.TxtRemark.Text
                .PrimaryDO = Me.txtPrimaryDO.Text
                .CetakNotaReturn = Me.chkTaxReturn.Checked
                .NoNotaRetur = Me.TxtNotaRetur.Text
                If Me.cboTipeSR.Text = "Revisi Invoice" Then
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
                .PaymentTerm = Math.Round(CDec(Me.TxtPaymentTerm.Text), 0, MidpointRounding.AwayFromZero)
                .TaxPercent = Me.txtTax.Text

                'Update 16 Digit NPWP
                .NPWPNo = Me.txtNoNPWP.Text

                .InvoiceNoQuickSR = Me.txtInvoiceNo.Text
                .TaxInvoiceNoQuickSR = Me.TxtTaxInvoiceNumber.Text
                .InvoiceDateQuickSR = Me.dtpInvoiceDate.Value
                .PrimaryCustomer = Me.TxtPrimaryCustomer.Text
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
            End With

            'salin exchange rate item untuk setiam item
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency.Text
                rowDGV.Cells(Me.ColExchangeRate.Index).Value = Me.txtTaxExchangeRate.Text
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim TanggalInvoice As Date
        Dim Dummy As String = ""
        Try
            'jika retur nya adalah nota retur pajak, maka nomor nota returnya kosong
            If Me.cboTipeSR.SelectedIndex = 0 Then
                Me.TxtNotaRetur.Clear()
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("sales return  number")
            KoleksiTextBox.Add(Me.TxtSalesReturnNo)
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("invoice number")
            KoleksiTextBox.Add(Me.txtInvoiceNo)
            KoleksiNamaData.Add("sales person")
            KoleksiTextBox.Add(Me.TxtSalesPCode)
            If CDec(Me.txtTax.Text) > 0 Then
                KoleksiNamaData.Add("tax invoice number")
                KoleksiTextBox.Add(Me.TxtTaxInvoiceNumber)
            End If
            'KoleksiNamaData.Add("stock in number")
            'KoleksiTextBox.Add(Me.TxtStockInNo)
            If Me.cboTipeSR.SelectedIndex = 1 Then
                KoleksiNamaData.Add("return note")
                KoleksiTextBox.Add(Me.TxtNotaRetur)
            End If
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'periode sales return dan stock in sales return harus sama --> untuk quick sales return pasti sama
            'If Format(Me.ObjStockInOut.GetTanggalTransaksi(Me.txtPrimaryStockIn.Text), "MMM yyyy") <> Format(Me.dtpDate.Value, "MMM yyyy") Then
            '    MsgBox("The stock in sales return and sales return must be dated on the same month. Please fix the data.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

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

            If Me.cboTipeSR.SelectedIndex = 1 Then
                'jika sudah ada nomor nota return yang sama maka berikan peringatan.
                Dim DummySR As String = Me.ObjSalesReturn.GetNoSalesReturn(Me.TxtNotaRetur.Text, Me.TxtPrimarySalesReturn.Text)
                If DummySR <> "" Then
                    If MsgBox("The following sales return has a tax return invoice number '" + Me.TxtNotaRetur.Text + "': " + vbCrLf + DummySR + vbCrLf + "Continue data saving process?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
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

                'jika Primary account berbeda dengan primary account punya AOP maka di blok (akan berpengaruh pada journal)
                Dim PrimaryAOP As Double = -1
                PrimaryAOP = ObjInv.GetPrimaryInvoice(Me.txtInvoiceNo.Text)
                If Me.txtPrimaryAccount.Text <> Me.ObjInv.GetPrimaryAccountAR(PrimaryAOP) Then
                    Dim PrimaryAccountInv As Double = -1
                    Dim AccountInv As String = ""
                    PrimaryAccountInv = Me.ObjInv.GetPrimaryAccountAR(PrimaryAOP)
                    AccountInv = ObjAkun.GetAccountCode(PrimaryAccountInv) & " - " & ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAccountInv)
                    MsgBox("account receivable of invoice no." & Me.txtInvoiceNo.Text & " is " & AccountInv & _
                           ", you cannot use different account payable with sales invoice.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'jika PrimarycurrTax berbeda dengan PrimarycurrTax punya AOP maka di blok (akan berpengaruh pada journal)
                If Me.txtPrimaryCurrTax.Text <> Me.ObjInv.GetPrimaryCurrTax(PrimaryAOP) Then
                    Dim PrimaryCurrTaxInv As Double = -1
                    Dim CurrTaxInv As String = ""
                    PrimaryCurrTaxInv = Me.ObjInv.GetPrimaryCurrTax(PrimaryAOP)
                    CurrTaxInv = Me.ObjCur.GetCurrencyCode(PrimaryCurrTaxInv)
                    MsgBox("Currency for tax of invoice no." & Me.txtInvoiceNo.Text & " is " & CurrTaxInv & _
                           ", you cannot use different Currency for tax with sales invoice.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Me.txtTax.Focus()
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Me.txtTax.Focus()
                        Return False
                    End If
                End If
            End If
            'untuk quick invoice, tidak bisa revisi invoice --> untuk quick sales return tidak ada batasan ini
            'If Me.ObjInv.ApakahQuickInvoice(Me.ObjInv.GetPrimaryInvoice(Me.txtInvoiceNo.Text)) = True Then
            '    If Me.cboTipeSR.SelectedIndex = 0 Then
            '        MsgBox("This data refers to a quick invoice, must be a return tax invoice.", MsgBoxStyle.Information, "Information")
            '        Me.cboTipeSR.SelectedIndex = 1
            '    End If
            'End If

            'cek tanggal invoice, harus lebih kecil dari tanggal SR
            If Me.dtpDate.Value.Date < Me.dtpInvoiceDate.Value.Date Then
                MsgBox("Invoice date must be smaller than sales return date.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If CDec(.Cells(Me.ColQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next

            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty or set ratio.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dim ObjCust As New CustomSMESS.ClsTransaksiStockInOut
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            TblDummy = Me.TblCustomInOutDummy.Copy
            If TblDummy.Rows.Count > 0 Then
                For ijk As Integer = 0 To Me.dgvItem.RowCount - 1
                    rd = Me.dgvItem.Rows(ijk)
                    With rd
                        Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                        Dim PrimaryPOSO As Double = -1 '.Cells(Me.ColPrimarySO.Index).Value
                        Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                        Dim Batch As String = .Cells(Me.ColBatchNo.Index).Value
                        Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                        Dim QtyMulti As Decimal = .Cells(Me.ColQtyMulti.Index).Value
                        Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                        If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryPOSO, PrimaryLocation, Batch, QtyDef, QtyMulti, TblDummy) = False Then
                            If String.IsNullOrEmpty(Batch) = False Then
                                ProdCode &= " batch: " & Batch
                            End If
                            MsgBox("Inconsisten data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            rd.Selected = True
                            Return False
                        End If

                    End With
                Next
            End If

            Dim ObjCustSR As New CustomSMESS.ClsTransaksiSalesReturn_Quick_AddByCustomProduct
            If ObjCustSR.ApakahDataBenar(Me.dgvItem, Me.ColPrimaryCodeNo.Index, Me.ColPriceByMultiUnit.Index, Me.ColPrimaryUnitsMulti.Index, Me.ColPrimaryUnits.Index) = False Then
                Return False
            End If
            ObjCustSR = Nothing

            'Update NWPW 16 Digit
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If ObjSetNPWP.SetNPWP(Me.dtpDate.Value.Date, Me.TxtPrimaryCustomer.Text, Me.txtNoNPWP.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Customer) = False Then
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
            hj.IsQuick = 1
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.ObjSalesReturn.CariDetail(hj.PrimaryTransaksi, Me.dgvItem, Me, Me.TblDummyCustom) = True Then
                    Me.NonAktifkanDetail()
                    Me.SalinKeField()
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
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FrmTransaksiSalesReturn_Quick.CheckForIllegalCrossThreadCalls = False
        Dim ObjLS As New clsMasterListedSoftware

        Try
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboTipeSR.SelectedIndex = 0
            Me.ColSecondaryQty.Visible = False
            Me.ColSecondaryUnit.Visible = False
            Me.ColBatchNo.Visible = False
            Me.ColLocation.Visible = False
            Me.ColSONO.Visible = False
            Me.ColCurrency.Visible = False
            Me.ColExchangeRate.Visible = False
            Me.ObjSalesReturn.QuickSR = True

            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsDefault.Left = Me.TxtDiscount.Right - Me.lblUnitsDefault.Width
            'Me.TxtQtyDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyDefault.Left - 6
            Me.lblSecQty.Visible = False

            Me.llbImport.Visible = False
            Me.ProgressBar1.Visible = False
            '-balikin lagi
            If Me.objCustom.ApakahBolehImportSR = True Then
                Me.llbImport.Visible = True
                Me.ProgressBar1.Visible = True
            End If
            Me.cboPrimaryAccount.Left = Me.cboAccount.Left
            Me.cboPrimaryAccount.Visible = False
            Me.cboPrimaryAccount.TabStop = False

            Me.ColRatioDefault.Visible = False
            Me.ColRatioReference.Visible = False
            Me.ColRemark.MaxInputLength = 250
            Me.ShowIcon = False

            Me.Atur1()
            Me.LinkLabel1.Parent = Me.gboItem
            Me.LinkLabel1.BringToFront()
            Me.LinkLabel1.Enabled = True
            Me.TxtPrimarySalesReturn.Text = "-1"
            Me.TxtPrimaryCustomer.Text = "-1"
            Me.ShowIcon = False
            Me.LinkLabel1.Visible = True
            Me.LinkLabel1.BringToFront()

            Me.btnUP.BackgroundImage = My.Resources.arrow_up
            Me.btnDown.BackgroundImage = My.Resources.arrow_down

            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.cboTipeSR.SelectedIndex = 1

            'Me.ColPriceByMultiUnit.Visible = True
            'Me.txtPriceByMultiUnit.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSalesReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjSalesReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
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
                    Me.KosongkanItem()
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjSalesReturn.CariDetail(Me.TxtPrimarySalesReturn.Text, Me.dgvItem, Me, Me.TblDummyCustom)
                        Me.SalinKeField()
                        Me.KosongkanItem()
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
                'hanya boleh edit yg berasal dari quick sr
                If Me.ObjSalesReturn.ApakahQuickSR(Me.TxtPrimarySalesReturn.Text) = False Then
                    MsgBox("Please use sales return form to edit this data", MsgBoxStyle.Information, "Information")
                Else
                    If Me.ObjSalesReturn.ApakahBolehEdit(Me.TxtPrimarySalesReturn.Text, Me.UserCode, True) = True Then
                        Me.Baru = False
                        Me.AktifkanDetail()
                        Me.dtpDate.Focus()
                        Me.SetTaxRate(False)
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
                'hanya boleh hapus dr yg berasal dr quick sr
                If Me.ObjSalesReturn.ApakahQuickSR(Me.TxtPrimarySalesReturn.Text) = False Then
                    MsgBox("Please use sales return form to delete this data", MsgBoxStyle.Information, "Information")
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
                            If Me.ObjSalesReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()

                                'konfirm ke user apakah akan buat data sales order cancellation
                                'If MsgBox("Create new sales order cancellation regarding this return?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                '    Dim FrmDia As New FrmDialogCancellationReason
                                '    If FrmDia.ShowDialog = Windows.Forms.DialogResult.OK Then
                                '        Me.ObjSalesReturn.SimpanSalesOrderCancellation(Me.dgvItem, Me.dgvSOC, FrmDia.TextBox1.Text, _
                                '        Me.UserCode, Me.ObjSalesReturn.NoTransaksi, Me)
                                '    Else
                                '        MsgBox("Sales order generation cancelled by user.", MsgBoxStyle.Information, "Information")
                                '    End If
                                'End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSalesReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom) = True Then
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
                If Me.cboTipeSR.SelectedIndex = 1 Then
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


    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal LotNo As String, ByVal SecQty As Double) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If CDec(row.Cells(Me.ColSecondaryQty.Index).Value) = SecQty And row.Cells(Me.ColPrimaryCodeNo.Index).Value _
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

    Private Sub cboTipeSR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipeSR.SelectedIndexChanged
        Try
            If Me.cboTipeSR.SelectedIndex = 1 Then
                Me.chkTaxReturn.Enabled = True
                Me.TxtNotaRetur.ReadOnly = False
                Me.TxtNotaRetur.BackColor = Me.txtPrimaryStockIn.BackColor
                Me.ColNettPrice.Visible = True
                Me.ColDisc.Visible = True
                Me.ColUnitPrice.Visible = True
                Me.ColCurrency.Visible = False
                Me.ColAmount.Visible = True
                Me.TxtDutystamp.ReadOnly = False
                Me.TxtDutystamp.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.ColExchangeRate.Visible = False
                Me.ColPriceByMultiUnit_Str.Visible = True
            Else
                Me.chkTaxReturn.Enabled = False
                Me.TxtNotaRetur.ReadOnly = True
                Me.TxtNotaRetur.BackColor = Me.TxtReadOnly.BackColor
                Me.chkTaxReturn.Checked = False
                Me.TxtNotaRetur.Text = ""
                Me.ColNettPrice.Visible = False
                Me.ColDisc.Visible = False
                Me.ColUnitPrice.Visible = False
                Me.ColCurrency.Visible = False
                Me.ColAmount.Visible = False
                Me.TxtDutystamp.ReadOnly = True
                Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                Me.ColExchangeRate.Visible = False
                Me.ColPriceByMultiUnit_Str.Visible = False
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
            Me.ObjCD.Number_KeyUp(Me.TxtDutystamp, sender, e, False)
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

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjSalesReturn.CariDetail(vPrimary, Me.dgvItem, Me, Me.TblDummyCustom) = True Then
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

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.txtCustomerCode.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = True Then
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        Try
            If Me.txtCustomerCode.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = True Then
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.Enter
        Try
            Me.buffer = Me.txtInvoiceNo.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.GotFocus
        Try
            If Me.txtInvoiceNo.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = True Then
                    Me.txtInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtInvoiceNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.Leave
        Dim dummy As Date
        Dim Update1 As Boolean = False
        Dim PrimaryInvoice As Double = -1
        Dim DummyStr As String = ""

        Try
            If Me.TxtRemark.Enabled = True Then
                Me.txtInvoiceNo.Text = Me.txtInvoiceNo.Text.Trim
                If Me.txtInvoiceNo.Text <> "" Then
                    Me.txtInvoiceNo.Text = Me.txtInvoiceNo.Text.Trim
                    If Me.txtInvoiceNo.Text <> Me.buffer Then
                        PrimaryInvoice = Me.ObjInv.GetPrimaryInvoice(Me.txtInvoiceNo.Text)
                        If PrimaryInvoice <> -1 Then
                            dummy = Me.ObjInv.GetTanggalTransaksi(PrimaryInvoice)
                            If dummy <> #1/1/1950# Then
                                If Me.dtpInvoiceDate.Value.Date <> dummy Then
                                    Me.dtpInvoiceDate.Value = dummy
                                    Update1 = True
                                End If
                            End If
                            DummyStr = Me.ObjInv.GetTaxInvoiceNumber(PrimaryInvoice)
                            If DummyStr <> "" Then
                                Me.TxtTaxInvoiceNumber.Text = DummyStr
                                Update1 = True
                            End If

                            If Me.dgvItem.Rows.Count = 0 Then
                                Me.KosongkanItem()
                                DummyStr = Me.ObjInv.GetPrimaryCurrency(PrimaryInvoice)
                                If DummyStr <> Me.TxtPrimaryCurrency.Text Then
                                    Me.TxtPrimaryCurrency.Text = DummyStr
                                    Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                                    Update1 = True
                                End If
                            End If

                            If Update1 = True Then
                                MsgBox("Invoice data updated", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.LostFocus
        Try
            If Me.txtInvoiceNo.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = True Then
                    Me.txtInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtInvoiceNo.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.GotFocus
        Try
            If Me.TxtTaxInvoiceNumber.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                    Me.TxtTaxInvoiceNumber.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxInvoiceNumber.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.LostFocus
        Try
            If Me.TxtTaxInvoiceNumber.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                    Me.TxtTaxInvoiceNumber.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtTaxInvoiceNumber.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNotaRetur_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotaRetur.GotFocus
        Try
            If Me.TxtNotaRetur.Enabled = True Then
                If Me.TxtNotaRetur.ReadOnly = True Then
                    Me.TxtNotaRetur.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtNotaRetur.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNotaRetur_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotaRetur.LostFocus
        Try
            If Me.TxtNotaRetur.Enabled = True Then
                If Me.TxtNotaRetur.ReadOnly = True Then
                    Me.TxtNotaRetur.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtNotaRetur.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.Enter
        Try
            Me.buffer = Me.TxtCurrency.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.GotFocus
        Try
            If Me.TxtCurrency.Enabled = True Then
                If Me.TxtCurrency.ReadOnly = True Then
                    Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCurrency.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.LostFocus
        Try
            If Me.TxtCurrency.Enabled = True Then
                If Me.TxtCurrency.ReadOnly = True Then
                    Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCurrency.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.GotFocus
        Try
            If Me.TxtPaymentTerm.Enabled = True Then
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

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        Try
            If Me.TxtPaymentTerm.Enabled = True Then
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

    Private Sub TxtDutystamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.GotFocus
        Try
            If Me.TxtDutystamp.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = True Then
                    Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDutystamp.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.LostFocus
        Try
            If Me.TxtDutystamp.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = True Then
                    Me.TxtDutystamp.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDutystamp.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = True Then
                    Me.TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemark.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = True Then
                    Me.TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemark.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomer.Text)
                'Update NPWP 16 Digit
                Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.TxtPrimaryCustomer.Text)
                '2025 Core Tax Tahap 2.1
                If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                    Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                End If
                Me.TxtSalesPCode.Focus()
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
                'Update NPWP 16 Digit
                Me.txtNoNPWP.Clear()
                Me.txtCustomerCode.Focus()
            End If
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.TxtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomer.Text)
                    'Update NPWP 16 Digit
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.TxtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    End If
                    Me.TxtSalesPCode.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    Me.BtnFindCustomer.PerformClick()
                End If
                Me.UpdateSalesPerson()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.TxtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomer.Text)
                    'Update NPWP 16 Digit
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.TxtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    End If
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomer.Text)
                        'Update NPWP 16 Digit
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.TxtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        End If
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.TxtPrimaryCustomer.Text = "-1"
                        'Update NPWP 16 Digit
                        Me.txtNoNPWP.Clear()
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.TxtPrimaryCustomer.Text = "-1"
                'Update NPWP 16 Digit
                Me.txtNoNPWP.Clear()
            End If
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.txtCustomerCode.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = False Then
                    Me.txtCustomerCode.BackColor = Me.BackColorAll
                Else
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        Try
            If Me.txtCustomerCode.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = False Then
                    Me.txtCustomerCode.BackColor = Me.TxtPrimaryCurrency.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Dim Rounding As Integer = 0

        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtSecondaryQty.Text = "0"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsDefault.Left = Me.TxtDiscount.Right - Me.lblUnitsDefault.Width
            'Me.TxtQtyDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyDefault.Left - 6
            Me.lblSecQty.Visible = False
            Me.lblExpiryDate.Visible = False
            Me.DtpExpired.Value = Today
            Me.DtpExpired.Visible = True
            Me.lblLocation.Visible = False
            Me.TxtLocation.Visible = True
            Me.btnFindLocation.Visible = True
            Me.lblBatchNo.Visible = False
            Me.txtBatchNo.Visible = True
            Me.txtBatchNo.TabStop = True
            Me.txtBatchNo.ReadOnly = False

            Rounding = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping("0", Rounding)
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            Me.txtAmount.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            Me.TxtDiscFromPerct.Text = "1"

            Me.SetGboItem()

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If
            '2.b. nilai qty default <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If
            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True Then
                If Me.txtBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtBatchNo.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryLocation.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtLocation.Focus()
                    Return False
                End If
            End If

            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnit.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQty.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQty.Focus()
                    Return False
                End If
            End If

            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlled.Checked = True And Me.chkExpiryDateControlled.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.txtBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value = Me.txtBatchNo.Text And Me.txtBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiryDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '8.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjStockInOut.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text, _
                    JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
            End If

            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
            End If

            '7. unit price/nett price =0
            If CDec(Me.TxtUnitPrice.Text) <= 0 Then
                MsgBox("Please specify positive data for unit price.", MsgBoxStyle.Information, "Information")
                Me.TxtUnitPrice.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.txtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

            Me.txtPrimaryPDO.Text = "-1"
            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtLocation.Text = ""
                Me.TxtLocation.ReadOnly = True
                Me.lblLocation.Visible = True
                Me.TxtLocation.Visible = False
                Me.btnFindLocation.Visible = False
            Else
                Me.TxtLocation.ReadOnly = False
                Me.lblLocation.Visible = False
                Me.TxtLocation.Visible = True
                Me.btnFindLocation.Visible = True
            End If
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
                Me.DtpExpired.Value = Today
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If
            If Me.chkBatchNoControlled.Checked = True Then
                Me.txtBatchNo.ReadOnly = False
                Me.txtBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.txtBatchNo.Visible = True
            Else
                Me.txtBatchNo.ReadOnly = True
                Me.txtBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.txtBatchNo.Visible = False
            End If
            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.txtSecondaryQty.Visible = False
                Me.lblSecondaryUnits.Visible = False
                'Me.lblUnitsDefault.Left = Me.TxtDiscount.Right - Me.lblUnitsDefault.Width
                'Me.TxtQtyDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyDefault.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                'Me.lblUnitsDefault.Left = Me.lblSecQty.Left - 37
                'Me.TxtQtyDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyDefault.Left - 6
                Me.txtSecondaryQty.Text = "0"
            End If
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnitsDefault.Text = "Units"
                Me.lblSecondaryUnits.Text = "Units"
            End If

            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboPriceByMultiUnit()
        Dim UnitPrice, Discount As Decimal
        Dim UnitCodeDummy As String = ""

        Try
            Me.cboPriceByMultiUnit.Items.Clear()

            If Me.CboUnit.Items.Count = 1 Then
                Me.LblUnitPrice.Text = "Unit Price / " + Me.CboUnit.Text
                Me.cboPriceByMultiUnit.Visible = False
                Me.cboPriceByMultiUnit.TabStop = False
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                Me.cboPriceByMultiUnit.SelectedIndex = 0
                Me.txtPriceByMultiUnit.Text = "1"
            Else
                Me.LblUnitPrice.Text = "Unit Price"
                Me.cboPriceByMultiUnit.Visible = True
                Me.cboPriceByMultiUnit.TabStop = True
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                UnitCodeDummy = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.CboUnit.Text <> UnitCodeDummy Then
                    Me.cboPriceByMultiUnit.Items.Add(Me.CboUnit.Text)
                    Me.cboPriceByMultiUnit.SelectedIndex = Me.ObjProduct.GetPriceByMultiUnit(Me.TxtPrimaryCodeNo.Text)
                Else
                    Me.cboPriceByMultiUnit.SelectedIndex = 0
                End If
                If Me.cboPriceByMultiUnit.Text <> UnitCodeDummy Then
                    Me.txtPriceByMultiUnit.Text = "1"
                Else
                    Me.txtPriceByMultiUnit.Text = "0"
                End If
                Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
            End If

            If Me.TxtPrimaryCurrency.Text.Trim = "" Then
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
            Me.HitungItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown, TxtStockInNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave, TxtStockInNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
            End If

            Me.SetGboItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click, btnFindStockIn.Click
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtPrimaryCustomer.Text.Trim = "" Or Me.TxtPrimaryCustomer.Text = "-1" Then
                Me.txtCustomerCode.Focus()
                MsgBox("Please specify customer.", MsgBoxStyle.Information, "Information")
            End If

            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.SetGboItem()
                If Me.txtBatchNo.Visible = True Then
                    Me.txtBatchNo.Focus()
                ElseIf Me.TxtLocation.Visible = True Then
                    Me.TxtLocation.Focus()
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vPrimaryLocation"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="PrimaryPdO"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vUnitPrice As Decimal, ByVal vDisc As Decimal, _
    ByVal vPrimaryUnitsMulti As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        And row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And row.Cells(Me.ColDisc.Index).Value = vDisc And vPrimaryUnitsMulti = row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# And row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And row.Cells(Me.ColDisc.Index).Value = vDisc And vPrimaryUnitsMulti = row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) And row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And row.Cells(Me.ColDisc.Index).Value = vDisc And vPrimaryUnitsMulti = row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) And row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And row.Cells(Me.ColDisc.Index).Value = vDisc And vPrimaryUnitsMulti = row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim rowDGV As DataGridViewRow
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0

        Try
            If Me.ApakahDataItemBenar = True Then
                DefPrecDefault = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value, Me.TxtUnitPrice.Text, Me.TxtDiscount.Text, _
                 Me.txtPrimaryUnitsMulti.Text)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text), DefPrecMulti)
                    SetQtyDefaultOnDgv(row, DefPrecDefault)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    'Me.dgvItem.Rows.Add("", "", "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                    'Me.lblUnits.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.txtBatchNo.Text, _
                    'ExpiredDateStr, Me.TxtLocation.Text, Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryLocation.Text), Me.TxtRemarkItem.Text, _
                    '-1, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnits.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)

                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColAmount.Index).Value = Me.txtAmount.Text
                        .Cells(Me.ColBatchNo.Index).Value = Me.txtBatchNo.Text
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColCurrency.Index).Value = Me.TxtCurrency.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColDisc.Index).Value = Me.TxtDiscount.Text
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = Me.TxtLocation.Text
                        .Cells(Me.ColNettPrice.Index).Value = Me.txtNettPrice.Text
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryPdO.Index).Value = -1
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                        .Cells(Me.ColPrimarySO.Index).Value = -1
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyDefault.Text
                        .Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                        .Cells(Me.ColSONO.Index).Value = ""
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.TxtUnitPrice.Text
                        .Cells(Me.ColWareHouse.Index).Value = Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                        .Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount.Text
                        .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text

                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.CboUnit.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(Me.ObjSO.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecMulti)
                        .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(Me.ObjSO.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecDefault)
                        Me.SetRatioOnDgv(rowDGV)

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = 1
                        .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = ""

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsMulti.Index).Value
                        Else
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                    End With
                End If

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'hitung total
                Me.HitungTotal()

                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtCurrency.ReadOnly = True
                    Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtCurrency.TabStop = False
                    Me.btnFindCurrency.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim DummyAutosize As DataGridViewAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.None
        Dim IsMultiUnit As Boolean = False
        Dim PriceByUnitDefault As Boolean = False

        Try
            If Me.dgvItem.AutoSizeColumnsMode <> DataGridViewAutoSizeColumnsMode.None Then
                DummyAutosize = Me.dgvItem.AutoSizeColumnsMode
            End If
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    row.Cells(Me.ColExpiryDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiryDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiryDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecondaryQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecondaryQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = "-"
                    row.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocation.Index).Value = "-"
                    row.Cells(Me.ColWareHouse.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    PriceByUnitDefault = True
                    Exit For
                End If
            Next
            Me.ColPriceByMultiUnit_Str.Visible = PriceByUnitDefault

            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = True
            End If

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If

            Me.dgvItem.AutoSizeColumnsMode = DummyAutosize
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub txtSecondaryQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            '7. unit price
            '8. discount
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatchNo.Index Or e.ColumnIndex = Me.ColExpiryDate.Index _
                Or e.ColumnIndex = Me.ColLocation.Index Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDisc.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = Me.ColSecondaryQty.Index Then
                        'cek secondary qty
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColExpiryDate.Index Then
                        'cek expiry date
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    'buang nol di belakang decimal jika edit kolom qty
                    If e.ColumnIndex = Me.ColQtyMulti.Index Or _
                    e.ColumnIndex = Me.ColDiscAmount.Index Or _
                    e.ColumnIndex = Me.ColDisc.Index Or _
                    e.ColumnIndex = Me.ColUnitPrice.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
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

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ObjUser As New Core_Module.ClsUser
        Try
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesReturn_Quick_AddByCustomProduct") = True Then
                Exit Sub
            End If

            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                Dim FrmRatio As New FrmTransaksiQuotation_SetQtyRatio
                Dim row As DataGridViewRow

                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                    Me.TxtRemark.Enabled = False Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value.ToString
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value.ToString
                FrmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                FrmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                FrmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                FrmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                FrmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                FrmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If FrmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = FrmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = FrmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioOnDgv(row)
                    Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)

                    ''batal jika qty default jadi 0
                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    End If

                    'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                    If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                        Dim ExpiryDate As Date = Now
                        Dim SecondaryQty As Decimal = 0
                        Dim JmlDGV As Decimal = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        If Me.ApakahEditSah(row, Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColPrimaryPdO.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, ExpiryDate, SecondaryQty, JmlDGV, False) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)
                        Else
                            Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                            If AdaSama = True Then
                                If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                                "Confirm") = MsgBoxResult.No Then
                                    'kembalikan data asal
                                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                                    Me.SetRatioOnDgv(row)
                                    Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)
                                    Exit Sub
                                End If

                                Me.GabungkanBaris()
                            Else
                                Me.HitungItem(e.RowIndex, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                                Me.HitungTotal()
                            End If
                        End If
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index And Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.HitungItem(e.RowIndex, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        'hitung total
                        Me.HitungTotal()
                    End If

                    FrmRef = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Double
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> Me.ColExpiryDate.Index Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                            ExpiryDate = CDate(.Cells(Me.ColExpiryDate.Index).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            SecondaryQty = .Cells(Me.ColSecondaryQty.Index).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = Me.ColQtyMulti.Index Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColExpiryDate.Index Then
                            Me.EndEditExpiryDate(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColDisc.Index Then
                            Me.EndEditDiscount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            Me.EndEditDiskonAmount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColUnitPrice.Index Then
                            Me.EndEditUnitPrice(sender, e, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemark.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditUnitPrice(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
   ByVal row As DataGridViewRow)
        Dim DefPrec As Integer = 0

        Try
            With row
                'untuk nilai unit price, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for unit price.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                'jika unit price < nett price maka kembalikan ke nilai buffer
                'hanya berlaku klo menentukan diskon berdasarkan amount
                If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0" Then
                    If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                If AdaSama = True Then
                    If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                    "Confirm") = MsgBoxResult.No Then
                        'kembalikan data asal
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    Me.GabungkanBaris()
                Else
                    'hitung ulang row
                    Me.HitungItem(row.Index, DefPrec)
                    'hitung total
                    Me.HitungTotal()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiscount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
   ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai diskon, pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter non-negative number for discount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek nili maksimum untuk diskon adalah 100%
                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    MsgBox("Maximum value is 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, 2)
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDisc.Index).Value))) / 100, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))

                'hanya hitung ulang kalo ada perubahan data
                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data asal
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                            ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDisc.Index).Value))) / 100, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
                            Exit Sub
                        End If
                        'tandai di ubah dari diskon
                        row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                        Me.GabungkanBaris()
                    Else
                        'tandai di ubah dari diskon
                        row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                        'hitung ulang row
                        Me.HitungItem(row.Index, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        'hitung total
                        Me.HitungTotal()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditExpiryDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan format yang dituliskan adalah format tanggal
                If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), _
                "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'pastikan tanggal expiry date tidak konflik dengan data yang sudah ada
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'bandingkan dengan yang sudah ada didalam dgv
                Dim row1 As DataGridViewRow
                For Each row1 In Me.dgvItem.Rows
                    If row1.Cells(Me.ColBatchNo.Index).Value = row.Cells(Me.ColBatchNo.Index).Value And CStr(row.Cells(Me.ColBatchNo.Index).Value) <> "" _
                    And row1.Index <> row.Index And row1.Cells(Me.ColPrimaryCodeNo.Index).Value = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        If CDate(row1.Cells(Me.ColExpiryDate.Index).Value).Date <> CDate(row.Cells(Me.ColExpiryDate.Index).Value).Date Then
                            MsgBox("You already entered expiry date for batch no: " + CStr(row1.Cells(Me.ColBatchNo.Index).Value) + " on " + _
                            Format(CDate(row1.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy"), MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                'bandingkan dengan yang sudah ada didalam transaksi lain. Detail lihat point no 8.c pada 
                'fungsi apakah data item benar
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjStockInOut.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItem(ColExpiryDate.Index, e.RowIndex).Value) Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(ColExpiryDate.Index, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")

                'expiry date, tidak perlu cek kemungkinan merger, karena tidak mungkin
                'hanya expiry date saja yang berbeda sedangkan batch no-nya sama
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatchNo.Index).Value, CDate(Me.buffer), SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Try
            With row
                DefPrecDefault = ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                DefPrecMulti = ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefPrecMulti)
                Me.SetQtyDefaultOnDgv(row, DefPrecDefault)
                If CDec(row.Cells(Me.ColQty.Index).Value) <= 0 Then
                    MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    SetQtyDefaultOnDgv(row, DefPrecDefault)
                    Exit Sub
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatchNo.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(row, DefPrecDefault)
                    Else
                        Me.HitungItem(e.RowIndex, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        'Me.GabungkanBaris() -> tidak mungkin ada penggabungan baris klo hanya edit qty saja
                        Me.HitungTotal()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatchNo.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNo(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper.Trim
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                Me.buffer, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) = True Then
                        '0. basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                        Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItem(ColExpiryDate.Index, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If

                        'cek juga ke transaksi:
                        '1. stock in (semua stock in)
                        Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                        For loopInOut As Integer = 0 To 4 Step 1
                            Select Case loopInOut
                                Case 0
                                    '   a. stock in purchase
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                                    Exit Select
                                Case 1
                                    '   b. stock in sales return
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                                    Exit Select
                                Case 2
                                    '   c. stock in adjust
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                                    Exit Select
                                Case 3
                                    '   d. stock in back
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                                    Exit Select
                                Case 4
                                    '   e. stock in proses
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                                    Exit Select
                            End Select
                            ExpiryDateLama = Me.ObjStockInOut.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, JenisTransaksi)
                            If ExpiryDateLama <> #1/1/1950# Then
                                If ExpiryDateLama.Date <> CDate(Me.dgvItem(ColExpiryDate.Index, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                                    Format(ExpiryDateLama, "dd MMM yyyy") + _
                                    ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    Else
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value)
                                        Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                    End If
                                End If
                            End If
                        Next
                        '2. transaksi stock bundling untuk item hasil bundling
                        ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                        Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItem(ColExpiryDate.Index, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        '3. proses stock taking
                        'msgbox("belum")

                        '4. data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItem.Rows
                            If row1.Cells(Me.ColBatchNo.Index).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value Then
                                If CDate(row1.Cells(Me.ColExpiryDate.Index).Value) <> CDate(Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value)
                                        Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value = row1.Cells(Me.ColExpiryDate.Index).Value
                                        Exit For
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(ColExpiryDate.Index, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocation(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                Dim Obj As New clsMasterStorageArea
                If Obj.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                Obj = Nothing
                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatchNo.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'update data primary location
                    Dim Obj1 As New clsMasterStorageArea
                    .Cells(Me.ColPrimaryLocation.Index).Value = Obj1.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWareHouse.Index).Value = Obj1.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)

                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'kembalikan data gudang
                            .Cells(Me.ColPrimaryLocation.Index).Value = Obj1.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                            'update data gudang
                            .Cells(Me.ColWareHouse.Index).Value = Obj1.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Obj1 = Nothing
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If

                    Obj1 = Nothing
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaBarisYangSama(Optional ByVal vCekDisc As Boolean = True) As Boolean
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= Me.dgvItem.RowCount - 2
                row = dgvItem.Rows(i)
                j = i + 1
                While j <= Me.dgvItem.RowCount - 1
                    row1 = dgvItem.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse

                    If row.Cells(Me.ColSONO.Index).Value = row1.Cells(Me.ColSONO.Index).Value And row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And _
                    row.Cells(Me.ColSuppDONo.Index).Value = row1.Cells(Me.ColSuppDONo.Index).Value And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) And row.Cells(Me.ColLocation.Index).Value = _
                    row1.Cells(Me.ColLocation.Index).Value And CDec(row1.Cells(Me.ColUnitPrice.Index).Value) = CDec(row.Cells(Me.ColUnitPrice.Index).Value) And _
                    CDec(row1.Cells(Me.ColPrimaryUnitsMulti.Index).Value) = CDec(row.Cells(Me.ColPrimaryUnitsMulti.Index).Value) Then
                        If vCekDisc = True Then
                            'cek kesamaan berdasarkan % disc
                            If CDec(row1.Cells(Me.ColDisc.Index).Value) = CDec(row.Cells(Me.ColDisc.Index).Value) Then
                                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                    If row.Cells(Me.ColSecondaryQty.Index).Value = row1.Cells(Me.ColSecondaryQty.Index).Value Then
                                        Return True
                                    Else
                                        Return False
                                    End If
                                ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                                    Return True
                                End If
                            End If
                        Else
                            'cek kesamaan berdasrkan disc amount
                            If CDec(row1.Cells(Me.ColDiscAmount.Index).Value) = CDec(row.Cells(Me.ColDiscAmount.Index).Value) Then
                                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                    If row.Cells(Me.ColSecondaryQty.Index).Value = row1.Cells(Me.ColSecondaryQty.Index).Value Then
                                        Return True
                                    Else
                                        Return False
                                    End If
                                ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                                    Return True
                                End If
                            End If
                        End If
                    End If
                    j += 1
                End While
                i += 1
            End While

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' menggabungkan 2 baris yang sama dalam dgv menjadi 1 dengan menambahkan qty-nya
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GabungkanBaris(Optional ByVal vCekDisc As Boolean = True)
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= Me.dgvItem.RowCount - 2
                row = dgvItem.Rows(i)
                j = i + 1
                While j <= Me.dgvItem.RowCount - 1
                    row1 = dgvItem.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse
                    If CStr(row.Cells(Me.ColSONO.Index).Value) = CStr(row1.Cells(Me.ColSONO.Index).Value) And _
                    CStr(row.Cells(Me.ColSuppDONo.Index).Value) = CStr(row1.Cells(Me.ColSuppDONo.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) And _
                    CStr(row.Cells(Me.ColLocation.Index).Value) = CStr(row1.Cells(Me.ColLocation.Index).Value) _
                    And CDec(row1.Cells(Me.ColUnitPrice.Index).Value) = CDec(row.Cells(Me.ColUnitPrice.Index).Value) And _
                    CDec(row1.Cells(Me.ColPrimaryUnitsMulti.Index).Value) = CDec(row.Cells(Me.ColPrimaryUnitsMulti.Index).Value) Then
                        If vCekDisc = True Then
                            'pengecekan berdasarkan kesamaan disc %
                            If CDec(row1.Cells(Me.ColDisc.Index).Value) = CDec(row.Cells(Me.ColDisc.Index).Value) Then
                                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                    If CDec(row.Cells(Me.ColSecondaryQty.Index).Value) = CDec(row1.Cells(Me.ColSecondaryQty.Index).Value) Then
                                        Me.GabungkanBaris(row, row1)
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                                    Me.GabungkanBaris(row, row1)
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            'pengecekan berdasarkan kesamaan disc amount
                            If CDec(row1.Cells(Me.ColDiscAmount.Index).Value) = CDec(row.Cells(Me.ColDiscAmount.Index).Value) Then
                                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                    If CDec(row.Cells(Me.ColSecondaryQty.Index).Value) = CDec(row1.Cells(Me.ColSecondaryQty.Index).Value) Then
                                        Me.GabungkanBaris(row, row1)
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                                    Me.GabungkanBaris(row, row1)
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        j += 1
                    End If
                End While
                i += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GabungkanBaris(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Try
            'gabungkan data qty.
            row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(row1.Cells(Me.ColQtyMulti.Index).Value), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
            Me.SetQtyDefaultOnDgv(row, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
            'gabungkan data remarks
            If CStr(row.Cells(Me.ColRemark.Index).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(Me.ColRemark.Index).Value = CStr(row.Cells(Me.ColRemark.Index).Value) + ". "
            End If
            row.Cells(Me.ColRemark.Index).Value = CStr(row.Cells(Me.ColRemark.Index).Value) + CStr(row1.Cells(Me.ColRemark.Index).Value)
            'hapus baris ybs
            Me.dgvItem.Rows.RemoveAt(row1.Index)
            'hitung ulang
            Me.HitungItem(row.Index, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
   Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
   ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit As Boolean) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(Me.ColSecondaryQty.Index).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(Me.ColBatchNo.Index).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Double
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjSRItem.GetJml(Me.TxtPrimarySalesReturn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)

            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    Dim Max_ As String = ObjCD.DigitGrouping((JmlObjLama - JmlGudang) * CDec(row.Cells(Me.ColRatioReference.Index).Value) / CDec(row.Cells(Me.ColRatioDefault.Index).Value), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                    MsgBox("Minimum value you must can enter is " & Max_ & " " & row.Cells(Me.ColUnitsMulti.Index).Value & ".", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
        End If
    End Sub

    Private Sub TxtLocation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = Obj.GetPrimary(Me.TxtLocation.Text)
                    Me.TxtQtyMulti.Focus()
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtQtyMulti.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtLocation.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    'Me.TxtQtyMulti.Focus()
                    Me.txtPrimaryLocation.Text = obj.GetPrimary(Me.TxtLocation.Text)
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        'Me.TxtQtyMulti.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        'Me.TxtLocation.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocation.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocation.Text <> "" Then
                i.txtCari.Text = Me.TxtLocation.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                Me.TxtQtyMulti.Focus()
            Else
                Me.TxtLocation.Clear()
                Me.TxtLocation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
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
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            Else
                Me.TxtCurrency.ReadOnly = False
                Me.TxtCurrency.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.TxtCurrency.TabStop = True
                Me.btnFindCurrency.Enabled = True
            End If
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
        Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
        JmlDGV, JmlSetelahHapus As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim row As DataGridViewRow = e.Row

        Try
            If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                With row
                    PrimaryLocation = .Cells(Me.ColPrimaryLocation.Index).Value
                    PrimaryPdO = .Cells(Me.ColPrimaryPdO.Index).Value
                    PrimaryProduct = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    BatchNo = .Cells(Me.ColBatchNo.Index).Value
                    If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                        ExpiryDate = CDate(.Cells(Me.ColExpiryDate.Index).Value)
                    Else
                        ExpiryDate = #1/1/1950#
                    End If
                    If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                        SecondaryQty = .Cells(Me.ColSecondaryQty.Index).Value
                    Else
                        SecondaryQty = -1
                    End If
                    JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                    'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv =0
                    JmlDGV = 0
                    JmlObjLama = Me.ObjSRItem.GetJml(Me.TxtPrimarySalesReturn.Text, PrimaryProduct, _
                    SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                    If JmlSetelahHapus < 0 Then
                        MsgBox("You cannot delete this data, because this will make you stock quantity minus.", _
                        MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = BackColorAll
            If Me.TxtPrimaryCustomer.Text.Trim = "" Or Me.TxtPrimaryCustomer.Text = "-1" Then
                Me.txtCustomerCode.Focus()
                MsgBox("Please specify customer.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.GotFocus
        Try
            Me.txtBatchNo.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.LostFocus
        Try
            Me.txtBatchNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.GotFocus
        Try
            Me.TxtLocation.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.LostFocus
        Try
            Me.TxtLocation.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.GotFocus
        Try
            Me.TxtQtyDefault.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.LostFocus
        Try
            Me.TxtQtyDefault.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.GotFocus
        Try
            Me.txtSecondaryQty.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.LostFocus
        Try
            Me.txtSecondaryQty.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Try
            Me.TxtRemarkItem.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim
            If Me.TxtRemark.Enabled = True Then
                If Me.txtBatchNo.Text <> "" And Me.chkExpiryDateControlled.Checked = True Then
                    Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                    If Me.DtpExpired.Value = #1/1/1950# Then
                        Me.DtpExpired.Value = Today
                    End If
                Else
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.chkTax.Checked = True Then
                    Me.txtTax.Enabled = True
                    Me.txtTax.Text = "10.00"
                    Me.txtTax.TabStop = True

                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy = "" Then
                        Dummy = "0"
                    End If
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                Else
                    Me.txtTax.Enabled = False
                    Me.txtTax.Text = "0.00"
                    Me.txtTax.TabStop = False
                End If
                Me.HitungTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrency.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCur.ApakahSudahAda(Me.TxtCurrency.Text) = True Then
                    Me.TxtPrimaryCurrency.Text = Me.ObjCur.GetPrimary(Me.TxtCurrency.Text)

                    DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                    For Each rowDGV In Me.dgvItem.Rows
                        Me.HitungItem(rowDGV.Index, DefPrec)
                    Next
                    Me.HitungTotal()

                    Me.TxtPaymentTerm.Focus()
                Else
                    If Me.TxtCurrency.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrency.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)

                        DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                        For Each rowDGV In Me.dgvItem.Rows
                            Me.HitungItem(rowDGV.Index, DefPrec)
                        Next
                        Me.HitungTotal()

                        Me.TxtPaymentTerm.Focus()
                    Else
                        Me.TxtCurrency.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtCurrency.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetTaxRate(ByVal vSetNewRate As Boolean)
        Dim Rate As Decimal = 0

        Try
            If Me.TxtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(1, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.TabStop = False
                Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.TabStop = True
                If vSetNewRate = True Then
                    Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.TxtPrimaryCurrency.Text)
                    If Rate = 0 Then
                        Rate = 1
                    End If
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, _
                    Me.ObjTaxRate.DefaultPrec)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.Leave
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim DTAQuickSR As New DseLaporanSalesReturnListTableAdapters.DataTable7TableAdapter
        Dim TBLQuickSR As New DseLaporanSalesReturnList.DataTable7DataTable
        'Dim rowQuickSR As DseLaporanSalesReturnList.DataTable7Row

        Try
            DTAQuickSR.Connection.ConnectionString = Core_Module.GetConstring
            DTAQuickSR.Fill(TBLQuickSR, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpDate.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpDate.Value))
            If Me.TxtCurrency.Text <> "" Then
                If Me.ObjCur.ApakahSudahAda(Me.TxtCurrency.Text) = True Then
                    Me.TxtPrimaryCurrency.Text = Me.ObjCur.GetPrimary(Me.TxtCurrency.Text)

                    DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                    For Each rowDGV In Me.dgvItem.Rows
                        Me.HitungItem(rowDGV.Index, DefPrec)
                    Next
                    If Me.buffer <> Me.TxtCurrency.Text Then
                        Me.SetTaxRate(True)
                    End If
                    Me.HitungTotal()
                Else
                    If Me.TxtCurrency.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrency.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)

                        DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                        For Each rowDGV In Me.dgvItem.Rows
                            Me.HitungItem(rowDGV.Index, DefPrec)
                        Next
                        If Me.buffer <> Me.TxtCurrency.Text Then
                            Me.SetTaxRate(True)
                        End If
                        Me.HitungTotal()

                    Else
                        Me.TxtCurrency.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtCurrency.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency.Text = "-1"
                MsgBox("Please speficy currency.", MsgBoxStyle.Information, "Information")
                Me.TxtCurrency.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            DTAQuickSR.Dispose()
            TBLQuickSR.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            i.ShowDialog()
            If i.PrimaryCurrency > 0 Then
                Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                Me.TxtCurrency.Text = Me.ObjCur.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                Me.TxtPaymentTerm.Focus()

                DefPrec = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                For Each rowDGV In Me.dgvItem.Rows
                    Me.HitungItem(rowDGV.Index, DefPrec)
                Next
                Me.SetTaxRate(True)
                Me.HitungTotal()
            Else
                Me.TxtCurrency.Clear()
                Me.TxtCurrency.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.Enabled = True Then
                    If Me.txtTax.ReadOnly = True Then
                        Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtTax.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtTax, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtTax, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            Me.ObjCD.Number_Leave(Me.txtTax, sender, e, 2, Me.ObjCD)
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.Enabled = True Then
                    If Me.txtTaxExchangeRate.ReadOnly = True Then
                        Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtTaxExchangeRate.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtTaxExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtTaxExchangeRate, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Try
            Me.ObjCD.Number_Leave(Me.txtTaxExchangeRate, sender, e, 2, Me.ObjCD)
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'txtExhcangeRate
    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyDefault, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Leave
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            Me.ObjCD.Number_Leave(Me.TxtQtyDefault, sender, e, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text), Me.ObjCD)
            If Me.TxtPrimaryCurrency.Text.Trim = "" Then
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
            Me.HitungItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            If CDec(Me.TxtUnitPrice.Text) = 0 Then
                Me.TxtDiscount.Text = 0
                Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            End If
            If CDec(Me.TxtDiscount.Text) <> CDec(Me.buffer) Then
                Me.TxtDiscFromPerct.Text = "1"
                Me.HitungItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            Me.buffer = Me.TxtDiscount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.buffer = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPrice, sender, e, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), Me.ObjCD)
            'unit price harus selalu lebih besar dari disc amount
            'hanya berlaku kalo ngitung disc berdasarkan disc amount, klo berdasarkan perc, langsung hitung ulang disc amountnya
            If Me.TxtDiscFromPerct.Text = "0" Then
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.buffer
                    Exit Sub
                End If
            End If
            Me.HitungItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungItem(ByVal vrowIndex As Integer, ByVal DefaultPrecision As Integer)
        Dim rowDgv As DataGridViewRow
        Dim DefaultPrecisionRow As Integer = 0
        Dim DiscPerct As Integer = 0

        Try
            rowDgv = Me.dgvItem.Rows(vrowIndex)
            DefaultPrecisionRow = Me.ObjCur.GetDefaultPrecission(rowDgv.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = rowDgv.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                rowDgv.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) * CDec(rowDgv.Cells(Me.ColDisc.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                rowDgv.Cells(Me.ColDisc.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            rowDgv.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) - CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            'hitung amount
            If Me.ObjProduct.ApakahUseSecondary(rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                If rowDgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    rowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(rowDgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowDgv.Cells(Me.ColQtyMulti.Index).Value), DefaultPrecision))
                Else
                    rowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(rowDgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowDgv.Cells(Me.ColQty.Index).Value), DefaultPrecision))
                End If
            Else
                If rowDgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    rowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(rowDgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowDgv.Cells(Me.ColQtyMulti.Index).Value) * CDec(rowDgv.Cells(Me.ColSecondaryQty.Index).Value), DefaultPrecision))
                Else
                    rowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(rowDgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowDgv.Cells(Me.ColQty.Index).Value) * CDec(rowDgv.Cells(Me.ColSecondaryQty.Index).Value), DefaultPrecision))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungItem()
        Dim DefaultPrecision As Integer = 0

        Try
            DefaultPrecision = Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
            'hitung nett price
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, DefaultPrecision)
            Else
                'hitung disc %
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
            End If
            'hitung nett price
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount.Text), _
            DefaultPrecision)

            'hitung amount
            If Me.chkUseSecondaryUnit.Checked = False Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtNettPrice.Text) * _
                    CDec(Me.TxtQtyMulti.Text), DefaultPrecision))
                Else
                    Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtNettPrice.Text) * _
                    CDec(Me.TxtQtyDefault.Text), DefaultPrecision))
                End If
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(0, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtDiscount.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtUnitPrice.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtUnitPrice.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.GotFocus
        Try
            If Me.TxtSalesPCode.Enabled = True Then
                If Me.TxtSalesPCode.ReadOnly = False Then
                    Me.TxtSalesPCode.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesPCode.BackColor = Me.TxtSalesPName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.LostFocus
        Try
            If Me.TxtSalesPCode.Enabled = True Then
                If Me.TxtSalesPCode.ReadOnly = False Then
                    Me.TxtSalesPCode.BackColor = Me.TxtPrimarySalesP.BackColor
                Else
                    Me.TxtSalesPCode.BackColor = Me.TxtSalesPName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalesP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSalesP.Click
        Dim FrmCariGroup As New FrmMasterSalesPerson_Cari

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimarySalesP.Text = -1 And Me.TxtSalesPCode.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtSalesPCode.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDec(FrmCariGroup.PrimarySalesPerson) > 0 Then
                Me.TxtPrimarySalesP.Text = FrmCariGroup.PrimarySalesPerson
                Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(FrmCariGroup.PrimarySalesPerson)
                Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(FrmCariGroup.PrimarySalesPerson)
                Me.txtInvoiceNo.Focus()
            Else
                Me.TxtPrimarySalesP.Text = -1
                Me.TxtSalesPCode.Text = ""
                Me.TxtSalesPName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesPCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPCode.Text) = True Then
                    Me.TxtPrimarySalesP.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPCode.Text)
                    Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                    Me.txtInvoiceNo.Focus()
                Else
                    Me.btnFindSalesP.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.Leave
        Try
            If Me.TxtSalesPCode.Text <> "" Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPCode.Text) = True Then
                    Me.TxtPrimarySalesP.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPCode.Text)
                    Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                Else
                    Me.btnFindSalesP.PerformClick()
                End If
            Else
                Me.TxtPrimarySalesP.Text = -1
                Me.TxtSalesPCode.Text = ""
                Me.TxtSalesPName.Text = ""
            End If
            Me.TxtSalesPCode.BackColor = Me.TxtPrimarySalesP.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateSalesPerson()
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCustomer.Text <> "-1" Then
                    If Me.TxtPrimarySalesP.Text = "-1" Or Me.TxtPrimarySalesP.Text = "" Then
                        Me.TxtPrimarySalesP.Text = Me.ObjCustomer.GetPrimary_SalesPerson(Me.TxtPrimaryCustomer.Text)
                        Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(Me.TxtPrimarySalesP.Text)
                        Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Enter
        Try
            Me.buffer = Me.TxtDiscAmount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.BackColorAll
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount, sender, e, CDec(Me.TxtUnitPrice.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDiscAmount, sender, e, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), Me.ObjCD)
            'disc amount tidak boleh lebih besar dari unit price
            If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                Me.TxtDiscAmount.Text = Me.buffer
                Exit Sub
            End If
            If CDec(Me.TxtDiscAmount.Text) <> CDec(Me.buffer) Then
                Me.TxtDiscFromPerct.Text = "0"
                Me.HitungItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.TxtPrimarySalesReturn.BackColor
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiskonAmount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim DefaultPrecision As Integer = 0

        Try
            With row
                DefaultPrecision = Me.ObjCur.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                'untuk nilai disc amount, pastikan nilainya positif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)
                row.Cells(Me.ColDisc.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(False)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            row.Cells(Me.ColDisc.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                            ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))

                            Exit Sub
                        End If
                        'tandai diskon dihitung berdasarkan amount
                        row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                        Me.HitungItem(row.Index, DefaultPrecision)
                        Me.GabungkanBaris(False)
                    Else
                        'tandai diskon dihitung berdasarkan amount
                        row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                        'hitung ulang row
                        Me.HitungItem(row.Index, DefaultPrecision)
                        'hitung total
                        Me.HitungTotal()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.Enabled = True Then
                    If Me.txtTax.ReadOnly = True Then
                        Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtTax.BackColor = Me.TxtPrimarySalesReturn.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.Enabled = True Then
                    If Me.txtTaxExchangeRate.ReadOnly = True Then
                        Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtTaxExchangeRate.BackColor = Me.TxtPrimarySalesReturn.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbImport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImport.LinkClicked
        Dim frmTransferDate As New FrmTransferBeginingBalance_Date

        Try
            If MsgBox("Start import data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If frmTransferDate.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.vStartDate = frmTransferDate.dtpFrom.Value
                    Me.vEndDate = frmTransferDate.dtpTo.Value

                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmTransferDate = Nothing
        End Try
    End Sub

    '-balikin lagi
    Private Sub TransferSalesReturn()
        Dim frmTransfer As New CustomSMESS.FrmTransferBeginingBalance
        Dim dtaSalesReturnCopy As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesreturn_copyTableAdapter
        Dim tblSalesReturnCopy As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_copyDataTable
        Dim rowSalesReturnCopy As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_copyRow

        Dim dtaSalesReturnItemCopy As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesreturn_item_copyTableAdapter
        Dim tblSalesReturnItemCopy As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item_copyDataTable
        'Dim rowSalesReturnItemCopy As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item_copyRow

        Dim tblSalesReturn As New dseTransaksiSalesReturn.tbltrans_salesreturn_NewDataTable
        Dim rowSalesReturn As dseTransaksiSalesReturn.tbltrans_salesreturn_NewRow

        Try
            Me.Cursor = Cursors.WaitCursor
            Me.ControlBox = False
            Me.btnFind.Enabled = False
            Me.btnNew.Enabled = False
            Me.btnEdit.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnPrint.Enabled = False
            Me.btnExit.Enabled = False
            Me.llbImport.Enabled = False

            Me.listOfFailed.Clear()
            Me.listOfSkiped.Clear()
            Me.listOfTransfered.Clear()
            ErrMsg = ""

            'cek terkait dengan data-data
            Me.LanjutTransfer = True
            If Me.objCustom.TransferSalesReturn_CekData(Me.ProgressBar1, Me.vStartDate.Date, Me.vEndDate.Date, Me.ErrMsg, Me.LblProgress, Me.tblOldSalesReturn, _
            Me.tblOldSalesReturnItem, Me.tblOldSalesReturnItem2, Me.tblCheckSR) = False Then
                Me.SedangKonfirmasi = True
                Me.BackgroundWorker1.ReportProgress(0)
                While Me.SedangKonfirmasi = True
                    Threading.Thread.Sleep(1000)
                End While
                If Me.LanjutTransfer = False Then
                    Exit Sub
                End If
            End If

            Me.objCustom.TransferSalesReturn(Me.ProgressBar1, Me.vStartDate, Me.vEndDate, Me.ErrMsg, Me.LblProgress, Me.tblOldSalesReturn, _
            Me.tblOldSalesReturnItem, Me.tblOldSalesReturnItem2, Me.tblCheckSR)

            dtaSalesReturnCopy.Fill(tblSalesReturnCopy)
            dtaSalesReturnItemCopy.Fill(tblSalesReturnItemCopy)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblSalesReturnCopy.Rows.Count
            For Each rowSalesReturnCopy In tblSalesReturnCopy
                With rowSalesReturnCopy
                    Me.ProgressBar1.Value = tblSalesReturnCopy.Rows.IndexOf(rowSalesReturnCopy) + 1
                    Me.LblProgress.Text = "Saving to database, data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + .No_Transaksi

                    If Me.ObjSalesReturn.ApakahSudahAda(.No_Transaksi) = False Then
                        tblSalesReturn.Clear()
                        rowSalesReturn = tblSalesReturn.Newtbltrans_salesreturn_NewRow
                        rowSalesReturn.Primary = .Primary
                        rowSalesReturn.No_Transaksi = .No_Transaksi
                        rowSalesReturn.Urutan = .Urutan
                        rowSalesReturn.Primary_Stock_In_Sales_Return = .Primary_Stock_In_Sales_Return
                        rowSalesReturn.Primary_DO = .Primary_DO
                        rowSalesReturn.Tanggal_Transaksi = .Tanggal_Transaksi
                        rowSalesReturn.Remarks = .Remarks
                        rowSalesReturn.Reason_Of_Return = .Reason_Of_Return
                        rowSalesReturn.Use_Nota_Retur_Pajak = .Use_Nota_Retur_Pajak
                        rowSalesReturn.Created_By = .Created_By
                        rowSalesReturn.Created_Date = .Created_Date
                        rowSalesReturn.Modified_By = .Modified_By
                        rowSalesReturn.Modified_Date = .Modified_Date
                        rowSalesReturn.Timestamp = .Timestamp
                        rowSalesReturn.No_Nota_Retur_Pajak = .No_Nota_Retur_Pajak
                        rowSalesReturn.Tipe_SR = .Tipe_SR
                        rowSalesReturn.Primary_Currency = .Primary_Currency
                        rowSalesReturn.Use_Tax = .Use_Tax
                        rowSalesReturn.Payment_Term = .Payment_Term
                        rowSalesReturn.Sub_Total = .Sub_Total
                        rowSalesReturn.Tax_Percent = .Tax_Percent
                        rowSalesReturn.Tax_Amount = .Tax_Amount
                        rowSalesReturn.Tax_Exchange_Rate = .Tax_Exchange_Rate
                        rowSalesReturn.Duty_stamp = .Duty_stamp
                        rowSalesReturn.Total = .Total
                        rowSalesReturn.primary_invoice = .primary_invoice
                        rowSalesReturn.Quick_Sales_Return = .Quick_Sales_Return
                        rowSalesReturn.Invoice_No_Quick = .Invoice_No_Quick
                        rowSalesReturn.Invoice_Date_Quick = .Invoice_Date_Quick
                        rowSalesReturn.Tax_Invoice_No_Quick = .Tax_Invoice_No_Quick
                        rowSalesReturn.Primary_Customer_Quick = .Primary_Customer_Quick
                        rowSalesReturn.Primary_Sales_Person = .Primary_Sales_Person
                        rowSalesReturn.primary_ar_account = .primary_ar_account
                        tblSalesReturn.Addtbltrans_salesreturn_NewRow(rowSalesReturn)

                        Me.ObjSalesReturn.CariDetail(tblSalesReturn, tblSalesReturnItemCopy, rowSalesReturn.Primary, Me.dgvItem, Me, Me.TblDummyCustom)
                        If Me.ObjSalesReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom) = True Then
                            listOfTransfered.Add(rowSalesReturnCopy.No_Transaksi)
                        Else
                            listOfFailed.Add(rowSalesReturnCopy.No_Transaksi)
                            'If MsgBox("Stop process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            '    Exit For
                            'End If
                        End If
                    Else
                        listOfSkiped.Add(rowSalesReturnCopy.No_Transaksi)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ProgressBar1.Maximum = 0
            Me.ProgressBar1.Value = 0
            Me.ControlBox = True
            Me.btnFind.Enabled = True
            Me.btnNew.Enabled = True
            Me.btnEdit.Enabled = True
            Me.btnDelete.Enabled = True
            Me.btnPrint.Enabled = True
            Me.btnExit.Enabled = True
            Me.llbImport.Enabled = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            '-balikin lg
            Me.ObjCD.UpdateThreading()
            Me.TransferSalesReturn()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            FrmErr.txtError.Text = Me.ErrMsg
            Me.LblProgress.Text = "Waiting form confirmation..."
            If FrmErr.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Me.LanjutTransfer = False
            Else
                Me.LanjutTransfer = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SedangKonfirmasi = False
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If listOfTransfered.Count > 0 Then
                Me.ErrMsg = "Sales return transfered: " & listOfTransfered.Count
            End If
            If listOfSkiped.Count > 0 Then
                If String.IsNullOrEmpty(Me.ErrMsg) = False Then
                    Me.ErrMsg += vbNewLine
                End If
                Me.ErrMsg += "Sales return skiped: " & listOfSkiped.Count
            End If
            If listOfFailed.Count > 0 Then
                If String.IsNullOrEmpty(Me.ErrMsg) = False Then
                    Me.ErrMsg += vbNewLine
                End If
                Me.ErrMsg += "Sales return failed: " & listOfFailed.Count
            End If
            If String.IsNullOrEmpty(Me.ErrMsg) = False Then
                MsgBox(Me.ErrMsg, MsgBoxStyle.Information, "Information")
            Else
                MsgBox("No sales return transfered.", MsgBoxStyle.Information, "Information")
            End If

            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.TxtQtyMulti.BackColor = BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyMulti, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Leave
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            Me.ObjCD.Number_Leave(Me.TxtQtyMulti, sender, e, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text), Me.ObjCD)
            If Me.TxtPrimaryCurrency.Text.Trim = "" Then
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
            Me.HitungItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.PrepareMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrepareMultiUnit()
        Dim Dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnit.DataSource = Nothing
            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                Dta.FillBy_PrimaryProduct(Me.mTblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.IsiCboUnit(Me.TxtPrimaryCodeNo.Text, Me.CboUnit, Me.tblCboUnit, Me.mTblMultiUnit)
                Me.CboUnit.SelectedIndex = -1
                Me.CboUnit.SelectedIndex = 0
            Else
                txtPrimaryUnitsDefault.Text = "-1"
                txtPrimaryUnitsMulti.Text = "-1"
                lblUnitsDefault.Text = "Units"
                'lblUnitsMulti.Text = "Units"
                TxtQtyDefault.Text = "0"
                TxtQtyMulti.Text = "0"
                LblUnitPrice.Text = "Unit Price"
                CboUnit.DataSource = Nothing
                CboUnit.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsiCboUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
            ByVal vTbl As DseTransaksiSO.CboUnitDataTable, _
            ByVal vTblMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitDataTable) As Boolean
        Dim rowTbl As DseTransaksiSO.CboUnitRow
        Dim rowMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitRow
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vTbl.Rows.Clear()
            vTblMultiUnit.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            rowTbl = vTbl.NewCboUnitRow
            rowTbl.Primary = objProduct.GetPrimaryUnit(vPrimaryProduct)
            rowTbl.Code = objUnit.GetUnitCode(rowTbl.Primary)
            vTbl.AddCboUnitRow(rowTbl)
            If vTblMultiUnit.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblMultiUnit.DefaultView.Count - 1
                    rowMultiUnit = vTblMultiUnit.DefaultView(i).Row
                    rowTbl = vTbl.NewCboUnitRow
                    rowTbl.Primary = rowMultiUnit.Primary_Unit
                    rowTbl.Code = objUnit.GetUnitCode(rowMultiUnit.Primary_Unit)
                    vTbl.AddCboUnitRow(rowTbl)
                Next
            End If
            vCbo.DataSource = vTbl.DefaultView
            vTbl.DefaultView.Sort = "[Code] ASC"
            vCbo.DisplayMember = "Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub CboUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.Leave
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            If Me.TxtPrimaryCurrency.Text.Trim = "" Then
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                  Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)

            Me.HitungItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal
        Dim row As DseTransaksiSO.CboUnitRow
        Try
            If Me.CboUnit.SelectedIndex <> -1 Then
                row = Me.tblCboUnit.DefaultView(Me.CboUnit.SelectedIndex).Row
                Me.txtPrimaryUnitsMulti.Text = row.Primary
                'Me.lblUnitsMulti.Text = Me.CboUnit.Text
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnit.Text
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                If UnitPrice = 0 And Discount = 0 Then
                    If Me.txtPriceByMultiUnit.Text = "1" Then
                        Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                        UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                    Else
                        Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                        UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                    End If
                End If
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.SetQtyDefault()
                Me.SetCboPriceByMultiUnit()
                'Me.HitungItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Dim RatioDefault As Decimal = 0
        Dim RatioReference As Decimal = 0
        Try
            RatioDefault = Me.ObjSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            RatioReference = Me.ObjSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * RatioDefault / RatioReference, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.LostFocus
        If Me.TxtQtyMulti.ReadOnly = False And Me.TxtQtyMulti.Enabled = True Then
            TxtQtyMulti.BackColor = Me.TxtPrimarySalesReturn.BackColor
        End If
    End Sub

    Private Sub TxtQtyMulti_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.TextChanged
        Try
            If IsNumeric(Me.TxtQtyMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtPrimaryCodeNo.Text) = True And _
            IsNumeric(Me.txtPrimaryUnitsMulti.Text) = True Then
                SetQtyDefault()
                Me.HitungItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.TextChanged
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtQtyMulti) = True Then
                    Me.HitungItem()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRow As DataGridViewRow, ByVal vDefPrecDefault As Integer)
        Try
            vRow.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(CDec(vRow.Cells(Me.ColQtyMulti.Index).Value) * CDec(vRow.Cells(Me.ColRatioDefault.Index).Value) / CDec(vRow.Cells(Me.ColRatioReference.Index).Value), vDefPrecDefault)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetRatioOnDgv(ByVal vDgvRow As DataGridViewRow)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitsMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SReturn_Quick_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.Baru = False
            Me.ObjSalesReturn.ResetObject(Me.dgvItem)

            Me.ObjSalesReturn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me, Me.TblDummyCustom)
            Me.SalinKeField()
            Me.UserCode = Me.ObjSalesReturn.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SReturn_Quick_Save(ByVal NoStockInSR As String, ByRef vPrimaryStockInSR As Double) As Double
        Try
            Me.Baru = True
            Me.ObjSalesReturn.QuickSR = True
            Me.ObjSalesReturn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me, Me.TblDummyCustom, NoStockInSR, False, False, True)
            vPrimaryStockInSR = Me.ObjSalesReturn.PrimaryStockInSalesReturn
            Return Me.ObjSalesReturn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(LinkLabel1, New Point(0, 15))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddByCustomBinaya()
        Dim Frm As New CustomSMESS.FrmTransaksiSalesReturn_Quick_AddItembyCustom
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            Frm.PrimarySalesReturn = Me.TxtPrimarySalesReturn.Text
            Frm.PrimaryCustomer = Me.TxtPrimaryCustomer.Text
            Frm.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
            Frm.TblDummyCustom = Me.TblDummyCustom
            Frm.TanggalTransaksi = Me.dtpDate.Value
            If Me.btnOK.Enabled = True Then
                Frm.AktifForm = True
            Else
                Frm.AktifForm = False
            End If
            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblDummyCustom = Frm.TblDummyCustom
                Me.IsikanDgvFromTblDummy()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim Frm As New CustomSMESS.FrmTransaksiSalesReturn_Quick_AddByCustomProduct
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ObjCustIdentifier As New CustomSMESS.ClsIdentifier
        Try

            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCustomer.Text = -1 And Me.TxtPrimaryCustomer.Text <> "" Then
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
            End If

            If ObjCustIdentifier.GetIdentifier = "smess_binaya" Then
                Me.AddByCustomBinaya()
                Exit Sub
            End If

            Frm.PrimarySalesReturn = Me.TxtPrimarySalesReturn.Text
            Frm.PrimaryCustomer = Me.TxtPrimaryCustomer.Text
            Frm.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
            Frm.TblItemCustomSR = Me.TblDummyCustom.Copy
            Frm.TblItemCustom = Me.TblCustomInOutDummy.Copy
            Frm.dtpDate = Me.dtpDate.Value
            If Me.btnOK.Enabled = True Then
                Frm.AktifForm = True
            Else
                Frm.AktifForm = False
            End If
            Frm.Baru = Me.Baru


            Dim ApakahCustomAdmiral As Boolean = False
            If ObjCustIdentifier.GetIdentifier = "smess_admiral2" Then
                ApakahCustomAdmiral = True
            End If

            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblDummyCustom = Frm.TblItemCustomSR
                Me.TblCustomInOutDummy = Frm.TblItemCustom
                If ApakahCustomAdmiral = False Then
                    Me.IsikanDgvFromTblDummy() 'dipanggil custom binaya
                Else
                    Me.IsiDgvByTbl()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try

        'sebelum penyesuaian dengan admiral 'binaya only
        'Dim Frm As New CustomSMESS.FrmTransaksiSalesReturn_Quick_AddByCustomProduct
        'Windows.Forms.Cursor.Current = Cursors.WaitCursor
        'Try
        '    Frm.PrimarySalesReturn = Me.TxtPrimarySalesReturn.Text
        '    Frm.PrimaryCustomer = Me.TxtPrimaryCustomer.Text
        '    Frm.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
        '    Frm.TblItemCustomSR = Me.TblDummyCustom
        '    Frm.dtpDate = Me.dtpDate.Value
        '    If Me.btnOK.Enabled = True Then
        '        Frm.AktifForm = True
        '    Else
        '        Frm.AktifForm = False
        '    End If
        '    If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
        '        Me.TblDummyCustom = Frm.TblItemCustomSR
        '        Me.IsikanDgvFromTblDummy()
        '    End If

        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    Windows.Forms.Cursor.Current = Cursors.Default
        'End Try
    End Sub

    Private Sub IsiDgvByTbl()
        Dim rowItemCustomSR As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1Row
        Dim rd As DataGridViewRow
        Dim PrimaryCurrenct As Double = -1
        Dim DecCurr As Integer = 0

        Try
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.TblDummyCustom.DefaultView.RowFilter = Nothing
            Me.TblDummyCustom.DefaultView.Sort = "[Product Code], [Expiry Date], [Batch No], [Qty At]"
            For i As Integer = 0 To Me.TblDummyCustom.DefaultView.Count - 1
                rowItemCustomSR = Me.TblDummyCustom.DefaultView(i).Row
                With rowItemCustomSR
                    Dim Index As Integer = Me.GetIndex(.Primary_Product, .Primary_Location, .Batch_No, .Secondary_Qty, -1, _
                    .Expiry_Date, .Unit_Price, .Discount, .Primary_Units_Multi)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        rd = Me.dgvItem.Rows(Index)
                        'tambahkan qty multi, 
                        rd.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyMulti.Index).Value) + .Total_Qty_Multi, .dec_multi)
                        rd.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQty.Index).Value) + .Total_Qty, ._Decimal)
                        If .is_kain = True Then
                            rd.Cells(Me.ColRatioDefault.Index).Value = rd.Cells(Me.ColQty.Index).Value
                            rd.Cells(Me.ColRatioReference.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value
                            rd.Cells(Me.ColRatioDesc.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value & " " & rd.Cells(Me.ColUnitsMulti.Index).Value & " = " & rd.Cells(Me.ColQty.Index).Value & " " & rd.Cells(Me.ColUnit.Index).Value
                        End If

                        If String.IsNullOrEmpty(.Remarks) = False Then
                            If String.IsNullOrEmpty(rd.Cells(Me.ColRemark.Index).Value) = False Then
                                rd.Cells(Me.ColRemark.Index).Value = rd.Cells(Me.ColRemark.Index).Value & ", " & .Remarks
                            Else
                                rd.Cells(Me.ColRemark.Index).Value = .Remarks
                            End If
                        End If

                        Me.HitungItem(rd.Index, DecCurr)
                    Else
                        'tambahkan baris baru
                        If i = 0 Then
                            DecCurr = Me.ObjCur.GetDefaultPrecission(.Primary_Currency)
                        End If
                        Dim ExpiredDateStr As String
                        If .Expiry_Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                        rd.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount, DecCurr)
                        rd.Cells(Me.ColBatchNo.Index).Value = .Batch_No
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColCurrency.Index).Value = Me.TxtCurrency.Text
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColDisc.Index).Value = Me.ObjCD.DigitGrouping(.Discount, 2)
                        rd.Cells(Me.ColExchangeRate.Index).Value = 1
                        rd.Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                        rd.Cells(Me.ColLocation.Index).Value = .Storage_Area_Code
                        rd.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(.Nett_Price, DecCurr)
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location
                        rd.Cells(Me.ColPrimaryPdO.Index).Value = -1
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimarySO.Index).Value = -1
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, ._Decimal)
                        rd.Cells(Me.ColSecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(.Secondary_Qty, ._Decimal)
                        rd.Cells(Me.ColSONO.Index).Value = ""
                        rd.Cells(Me.ColSuppDONo.Index).Value = ""
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price, DecCurr)
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, DecCurr)
                        rd.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct

                        rd.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty_Multi, .dec_multi)
                        rd.Cells(Me.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rd.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        If .is_kain = True Then
                            rd.Cells(Me.ColRatioDefault.Index).Value = rd.Cells(Me.ColQty.Index).Value
                            rd.Cells(Me.ColRatioReference.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value
                        Else
                            rd.Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Default, ._Decimal)
                            rd.Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, .dec_multi)
                        End If

                        If rd.Cells(Me.ColPrimaryUnitsMulti.Index).Value <> rd.Cells(Me.ColPrimaryUnits.Index).Value Then
                            If .is_kain = True Then
                                rd.Cells(Me.ColRatioDesc.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value & " " & rd.Cells(Me.ColUnitsMulti.Index).Value & " = " & rd.Cells(Me.ColQty.Index).Value & " " & rd.Cells(Me.ColUnit.Index).Value
                            Else
                                rd.Cells(Me.ColRatioDesc.Index).Value = rd.Cells(Me.ColRatioReference.Index).Value & " " & rd.Cells(Me.ColUnitsMulti.Index).Value & " = " & rd.Cells(Me.ColRatioDefault.Index).Value & " " & rd.Cells(Me.ColUnit.Index).Value
                            End If
                        Else
                            rd.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If

                        rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        'Me.SetRatioOnDgv(rd)

                        rd.Cells(Me.ColPriceByMultiUnit.Index).Value = 1
                        rd.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = ""

                        rd.Cells(Me.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                            rd.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Units_Code_Multi
                        Else
                            rd.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
                    End If
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.FormatDGV()

            'hitung total
            Me.HitungTotal()

            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub Atur1()
        Dim ObjUser As New Core_Module.ClsUser
        Try
            Me.AddItemByCustom.Visible = False
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesReturn_Quick_AddByCustomProduct") = True Then
                Me.SplitContainer2.Panel1.Enabled = False
                Me.SplitContainer2.Panel2.Enabled = False
                Me.ToolStripMenuItem1.Visible = True
                Me.dgvItem.ReadOnly = True
                Me.dgvItem.AllowUserToDeleteRows = False
            Else
                Me.gboItem.Enabled = True
                Me.ToolStripMenuItem1.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
            End If
            Me.LinkLabel1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanDgvFromTblDummy() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowDGV As DataGridViewRow
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Dim rowDummy As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1Row
        Dim ObjP As New Core_Module.clsMasterProduct
        Try
            Me.dgvItem.Rows.Clear()
            For Each rowDummy In Me.TblDummyCustom.Rows
                DefPrecDefault = ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product)
                DefPrecMulti = ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Units_Multi)
                Dim Index As Integer = Me.GetIndex(rowDummy.Primary_Product, rowDummy.Primary_Location, rowDummy.Batch_No, _
                 rowDummy.Secondary_Qty, -1, rowDummy.Expiry_Date, rowDummy.Unit_Price, rowDummy.Discount, _
                 rowDummy.Primary_Units_Multi)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(rowDummy.Total_Qty_Multi), DefPrecMulti)
                    SetQtyDefaultOnDgv(row, DefPrecDefault)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If rowDummy.Expiry_Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(rowDummy.Expiry_Date, "dd MMM yyyy")
                    End If

                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColRemark.Index).Value = rowDummy.Remarks
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Amount, rowDummy.Default_Prec_Currency)
                        .Cells(Me.ColBatchNo.Index).Value = rowDummy.Batch_No
                        .Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                        .Cells(Me.ColCurrency.Index).Value = rowDummy.Currency
                        .Cells(Me.ColDescription.Index).Value = rowDummy.Description
                        .Cells(Me.ColDisc.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Discount, 2)
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = rowDummy.Storage_Area_Code
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Nett_Price, rowDummy.Default_Prec_Currency)
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDummy.Primary_Currency
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDummy.Primary_Location
                        .Cells(Me.ColPrimaryPdO.Index).Value = -1
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = rowDummy.Primary_Secondary_Units
                        .Cells(Me.ColPrimarySO.Index).Value = -1
                        .Cells(Me.ColPrimaryUnits.Index).Value = rowDummy.Primary_Units
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Total_Qty, rowDummy._Decimal)
                        .Cells(Me.ColSecondaryQty.Index).Value = rowDummy.Secondary_Qty
                        .Cells(Me.ColSONO.Index).Value = rowDummy.SO_No
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColUnit.Index).Value = rowDummy.Unit_Code
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Unit_Price, rowDummy.Default_Prec_Currency)
                        .Cells(Me.ColWareHouse.Index).Value = Me.ObjLocation.GetWarehouseCode(rowDummy.Primary_Location)
                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Disc_Amount, rowDummy.Default_Prec_Currency)
                        .Cells(Me.ColDiscFromPerct.Index).Value = rowDummy.Disc_From_Perct

                        DefPrecMulti = ObjP.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Units_Multi)
                        .Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Total_Qty_Multi, DefPrecMulti)
                        .Cells(Me.ColUnitsMulti.Index).Value = rowDummy.Units_Code_Multi
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowDummy.Primary_Units_Multi
                        .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Reference, DefPrecMulti)
                        .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Default, rowDummy._Decimal)
                        Me.SetRatioOnDgv(rowDGV)
                    End With
                End If

                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtCurrency.ReadOnly = True
                    Me.TxtCurrency.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtCurrency.TabStop = False
                    Me.btnFindCurrency.Enabled = False
                End If
            Next

            Me.KosongkanItem()
            Me.FormatDGV()
            Me.HitungTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
        End Try
    End Function

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiSalesReturn_Quick_AddByCustomProduct
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPriceByMultiUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriceByMultiUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            If Me.cboPriceByMultiUnit.SelectedIndex <> -1 Then
                If Me.TxtRemarkItem.Enabled = True Then
                    If Me.cboPriceByMultiUnit.Text <> Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) Then
                        Me.txtPriceByMultiUnit.Text = "1"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                        Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                        If UnitPrice = 0 And Discount = 0 Then
                            Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                                UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.TxtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                        Me.TxtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
                        If UnitPrice = 0 And Discount = 0 Then
                            Me.ObjSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text, _
                            UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.TxtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    End If

                    If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        If CDec(Me.TxtUnitPrice.Text) > 0 Then
                            Me.HitungItem()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class