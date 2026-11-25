Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPurchaseReturn
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjCurrency As New Core_Module.clsMasterCurrency

    Dim baru As Boolean
    Dim objPurchaseReturn As New ClsTransaksiPurchaseReturn
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjPurchaseReturnItem As New ClsTransaksiPurchaseReturnItem
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjLS As New clsMasterListedSoftware
    Dim objGRN As New ClsTransaksiGRN
    Dim objGRNItem As New ClsTransaksiGRNItem
    Dim objInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjCL As New ClsTransaksiMonthlyClosing
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjLicense As New ClsLicense
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim buffer As String = ""

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================

    Dim DseAccount As New AccModule.DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New AccModule.DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New AccModule.DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New AccModule.DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As AccModule.DseMultiAccount.tblCurrencyTaxRow

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim objMulti As New AccModule.ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryAccount.Items.Clear()
            Me.tblAccount.Rows.Clear()
            Me.txtPrimaryAccount.Text = -1

            If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                Me.cboAccount.Enabled = False
                Me.cboPrimaryAccount.Enabled = False
                Me.txtPrimaryAccount.Text = -1
            Else
                If Me.TxtRemark.Enabled = True Then
                    Me.cboAccount.Enabled = True
                    Me.cboPrimaryAccount.Enabled = True
                End If
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
        Dim objMulti As New AccModule.ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboCurrTax.DataSource = Nothing
            Me.cboCurrTax.Items.Clear()
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.txtPrimaryCurrTax.Text = -1

            If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                Me.cboCurrTax.Enabled = False
                Me.cboPrimaryCurrTax.Enabled = False
                Me.txtPrimaryCurrTax.Text = -1
            Else
                If Me.TxtRemark.Enabled = True Then
                    Me.cboCurrTax.Enabled = True
                    Me.cboPrimaryCurrTax.Enabled = True
                End If
               objMulti.IsiTBLCurrTax(Me.TxtPrimaryCurrency.Text, Me.tblCurrTax)
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

    Private Sub cboCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutyStamp.Text) Then
                    Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.HitungTotal_Multi()
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
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutyStamp.Text) Then
                    Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.HitungTotal_Multi()
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

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCurrency.TextChanged

        Try
            If IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso CDec(Me.txtPrimaryCurrency.Text) > -1 Then
                If Me.TxtRemark.Enabled = True Then
                    If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                        Me.IsiCBOArAkunByTipePerkiraan1()
                        Me.IsiCBOCurrencyTax()
                    End If
                End If
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

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.FormatTxtStockIN()
            Me.cboPPNDipungut.Enabled = False
            'cboTipePurchaseReturn_SelectedIndexChanged(Nothing, Nothing)

         
            'If Me.txtPrimaryCurrTax.Text = -1 Then
            '    If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
            '        Me.cboAccount.Enabled = True
            '        Me.cboCurrTax.Enabled = True
            '        Me.IsiCBOArAkunByTipePerkiraan1()
            '        Me.IsiCBOCurrencyTax()
            '    End If
            'End If
            ''@rosadinnn
            ''@MUaT
            FormatDGV()

            Me.chkTax.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)


            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            Me.chkTax.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Dim DefPrec As Integer = 0

        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

            Me.TxtPrimaryPurchaseReturn.Text = "-1"
            Me.txtPrimaryGRN.Text = "-1"
            Me.TxtPrimaryPurchaseReturn.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryAccount.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.DGVGRN.Rows.Clear()

            Me.TxtPrimaryCurrency.Text = "-1"
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            Me.TxtPaymentTerm.Text = "0"
            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Me.chkTax.Checked = False
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            Me.cboTipePurchaseReturn.SelectedIndex = 0
            Me.ChkInvoiceCancellation.Checked = False

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim NoAOP As String = ""
        Dim NoStockOut As String = ""
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
        Dim Jurnal As String = ""
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim ObjAkun As New ClsMasterAkun1
        Dim ConvertionDate As Date = #1/1/1950#
        Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing

        Try
            ConvertionDate = ObjAcc.GetSMESSConvertionDate
            If ConvertionDate.Date > Me.dtpDate.Value.Date Then
                MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(ConvertionDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("stock in number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek supplier jangan sampai tidak diisikan
            If Me.TxtPurchaseReturnNo.Text = "-1" Then
                MsgBox("Please specify stock in purchase number", MsgBoxStyle.Information, "Information")
                Me.TxtPurchaseReturnNo.Focus()
                Return False
            End If

            'berkaitan dengan tanggal closing
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.objPurchaseReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objPurchaseReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objPurchaseReturn.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.objPurchaseReturn.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'untuk purchase return yang nota retur, maka harus sudah ada AOP yang bisa dirujuk oleh purchase return tersebut
            If Me.ObjAOP.GetPrimaryAOP_ByPrimaryGRN(Me.txtPrimaryGRN.Text) = -1 And (Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2) Then
                MsgBox("No AOP detected. Please specify AOP for tax return notes purchase return.", _
                     MsgBoxStyle.Information, "Information")
                Me.cboTipePurchaseReturn.Focus()
                Return False
            ElseIf Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                'sedangkan untuk transaksi purchase return revisi AOP, tanggal stock in dan tanggal purchase return harus berada dalam 
                'bulan yang sama
                If Format(Me.objGRN.GetTanggalTransaksi(Me.txtPrimaryGRN.Text), "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    MsgBox("For a purchase return to generate a aop revision the AOP must be dated on the same month as the purchase return.", _
                         MsgBoxStyle.Information, "Information")
                    Me.cboTipePurchaseReturn.Focus()
                    Return False
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objPurchaseReturn.TimeStamp <> Me.objPurchaseReturn.GetTimestamp(Me.objPurchaseReturn.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objPurchaseReturn.CariDetail(Me.objPurchaseReturn.Primary, Me.dgvItem, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If

                '1. Jika sudah ada AOP yang merujuk ke purchase return ini
                'untuk AOP, berikan pesan no AOP dan no stock out purchase return
                NoAOP = ObjAOP.GetNoAOPBy_PrimaryPurchaseReturn(Me.objPurchaseReturn.Primary)
                If NoAOP <> "" Then
                    NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.objPurchaseReturn.Primary)
                    MsgBox("This purchase return is refer by AOP #" + NoAOP + _
                    " and stock out #" + NoStockOut + ". Please delete the transaction to edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '3. sudah ada stock out purchase return yang merujuk ke purchase return ini
                If DTA.ScalarQuery_GetJmlStockOut_PR(TipeTrans, Me.objPurchaseReturn.Primary).Value > 0 Then
                    NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.objPurchaseReturn.Primary)
                    MsgBox("This purchase return is refer by stock out #" + NoStockOut + _
                    ". Please delete the transaction to edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '4. jika sudah ada purchase return receivement journal
                Jurnal = ObjJurnal.Get_PurchaseReturn_Jurnal(Me.objPurchaseReturn.Primary)
                If Jurnal <> "" Then
                    MsgBox("This data is refered by purchase return receivement jounal: " + Jurnal + _
                        ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'transaksi aop yang dirujuk merupakan quick aop dan tipenya revisi aop
            If Me.txtAopNo.Text <> "" And Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                If Me.ObjAOP.ApakahQuickPurchase(Me.ObjAOP.GetPrimary(Me.txtAopNo.Text)) = True Then
                    MsgBox("This transaction refers to a quick aop, you cannot save purchase return with aop revision type.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


            If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                'jika Primary account berbeda dengan primary account punya AOP maka di blok (akan berpengaruh pada journal)
                Dim PrimaryAOP As Double = -1
                PrimaryAOP = ObjAOP.GetPrimary(Me.txtAopNo.Text)
                If PrimaryAOP > -1 Then
                    If Me.txtPrimaryAccount.Text <> Me.ObjAOP.GetPrimaryAccountAP(PrimaryAOP) Then
                        Dim PrimaryAccountAOP As Double = -1
                        Dim AccountAOP As String = ""
                        PrimaryAccountAOP = Me.ObjAOP.GetPrimaryAccountAP(PrimaryAOP)
                        AccountAOP = ObjAkun.GetAccountCode(PrimaryAccountAOP) & " - " & ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAccountAOP)
                        MsgBox("account payable of AOP No." & Me.txtAopNo.Text & " is " & AccountAOP & _
                               ", you cannot use different account payable with transaction aop.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'jika PrimarycurrTax berbeda dengan PrimarycurrTax punya AOP maka di blok (akan berpengaruh pada journal)
                    If Me.txtPrimaryCurrTax.Text <> Me.ObjAOP.GetPrimaryCurrTax(PrimaryAOP) Then
                        Dim PrimaryCurrTaxAOP As Double = -1
                        Dim CurrTaxAOP As String = ""
                        PrimaryCurrTaxAOP = Me.ObjAOP.GetPrimaryCurrTax(PrimaryAOP)
                        CurrTaxAOP = Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTaxAOP)
                        MsgBox("Currency for tax of AOP No." & Me.txtAopNo.Text & " is " & CurrTaxAOP & _
                               ", you cannot use different Currency for tax with transaction aop.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'jika tipenya nota retur pajak dan tax checked, maka harus isikan no nota retur pajak
            If Me.chkTax.Checked = True And (Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2) Then
                Me.TxtTaxReturnNotes.Text = Me.TxtTaxReturnNotes.Text.Trim
                'perubahan hanya tipe  purchase yang tax retur notes saja yang tax retur nya tidak boleh kosong
                If Me.TxtTaxReturnNotes.Text = "" And cboTipePurchaseReturn.Text = "Tax Return Notes" Then
                    MsgBox("Please specify tax return notes.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkTax.Checked = True And (Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2) Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
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

            If Me.objInOut.ApakahStockInPurchaseQuick(Me.txtPrimaryStockInPurchase.Text) = True Then
                MsgBox("You can't enter stock in purchase from quick aop as reference.", MsgBoxStyle.Information, "Information")
                Return False

                Return False
            End If

            'Update NPWP 16 Digit
            If cboTipePurchaseReturn.SelectedIndex = 1 And chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true dan tipe nya tax retur note
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimarySupplier.Text, Me.TxtSupplierNPWPNo.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier, _
                 frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionPurchaseReturn, Me.UserCode, Me.TxtPrimaryPurchaseReturn.Text) = False Then
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
            DTA.Dispose()
            ObjJurnal = Nothing
            ObjAcc = Nothing
        End Try
    End Function


    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objPurchaseReturn.ResetObject(Me.dgvItem)
            Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)


            Me.TxtPrimaryCurrency.Text = "1"
            Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
            If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If
          
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objPurchaseReturn
                .Primary = Me.TxtPrimaryPurchaseReturn.Text
                .NoTransaksi = Me.TxtPurchaseReturnNo.Text
                .PrimaryGRN = Me.txtPrimaryGRN.Text
                .TanggalTransaksi = Me.dtpDate.Value
                If Me.chkTax.Checked = True Then
                    .TaxExclude = True
                Else
                    .TaxExclude = False
                End If
                .ReasonOfReturn = Me.TxtReasonOfReturn.Text
                .Remarks = Me.TxtRemark.Text
                .SupplierDONumber = Me.txtSupplierDONumber.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .Dutystamp = Me.TxtDutystamp.Text
                .PaymentTerm = Me.TxtPaymentTerm.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                If cboTipePurchaseReturn.SelectedIndex = 0 Then
                    .TipeReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Revisi_AOP
                Else
                    .TipeReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
                End If
                .Total = Me.TxtTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .SubTotal = Me.txtSubTotalAfterDP.Text
                .Tax = Me.txtTax.Text
                .PrimarySupplier = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text

                .NoNPWPSupplier = Me.TxtSupplierNPWPNo.Text
                .NoNotaReturPajak = Me.TxtTaxReturnNotes.Text
                If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    .VATDipungut = True
                Else
                    .VATDipungut = False
                End If

                If ChkInvoiceCancellation.Checked = True Then
                    .InvoiceCancellation = 1
                Else
                    .InvoiceCancellation = 0
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim rounding As Double = 0
        Dim DefPrecIDR As Integer = 0

        Try
            DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            With Me.objPurchaseReturn
                Me.TxtPrimaryPurchaseReturn.Text = .Primary
                Me.TxtPurchaseReturnNo.Text = .NoTransaksi
                Me.txtPrimaryStockInPurchase.Text = Me.objGRN.GetPrimaryStockIn(.PrimaryGRN)
                Me.txtPrimaryGRN.Text = .PrimaryGRN
                Me.dtpDate.Text = .TanggalTransaksi
                Me.TxtReasonOfReturn.Text = .ReasonOfReturn
                Me.TxtRemark.Text = .Remarks
                Me.txtSupplierDONumber.Text = .SupplierDONumber
                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency

                If Me.TxtPrimaryCurrency.Text <> "-1" And Me.TxtPrimaryCurrency.Text <> "" Then
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    rounding = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                End If

                If .TipeReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur Then
                    Me.cboTipePurchaseReturn.SelectedIndex = 1
                Else
                    Me.cboTipePurchaseReturn.SelectedIndex = 0
                End If


                If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
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
                    Else
                        Me.cboAccount.DataSource = Nothing
                        Me.cboPrimaryCurrTax.DataSource = Nothing
                        Me.txtPrimaryCurrTax.Text = -1
                    End If
                End If

                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(.Dutystamp, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.Tax, 2)

                Me.TxtPaymentTerm.Text = Me.ObjCD.DigitGrouping(.PaymentTerm, 0)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                If .TaxExclude = True Then
                    Me.chkTax.Checked = True
                Else
                    Me.chkTax.Checked = False
                End If

                Me.cboTipePurchaseReturn.SelectedIndex = .TipeReturn
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, rounding))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, rounding))

                Me.TxtSupplierNPWPNo.Text = .NoNPWPSupplier
                Me.TxtTaxReturnNotes.Text = .NoNotaReturPajak
                If .VATDipungut = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If

                If .InvoiceCancellation = 1 Then
                    Me.cboTipePurchaseReturn.SelectedIndex = 2
                    Me.ChkInvoiceCancellation.Checked = True
                Else
                    Me.ChkInvoiceCancellation.Checked = False
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

            End With
            Me.FormatDGV()
            'isikan data DO
            Me.objGRNItem.IsiDGV_Multi(Me.DGVGRN, CDbl(Me.txtPrimaryGRN.Text))
            Me.FormatDGVGRN()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows

                ''@rosadinnn
                ''@MUaT
                'row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)

                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    row.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty.Index).Value)
                End If
            Next

            Me.IsiDataStockIn()
            Me.HitungTotal_Multi()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub HitungTotal_Multi()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal, TaxAmount, Total, SubtotalAfterDP, Dutystamp, Tax, TaxExchangeRate As Decimal
        Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecIDR As Integer = 0
        Dim AmountInIDR As Decimal = 0
        Dim SubTotal1 As Decimal = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.HitungTotal_Multi_MiddlePoint()
                Exit Sub
            End If
            DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    Dutystamp = CDec(Me.TxtDutystamp.Text)
                Else
                    Dutystamp = 0
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    TaxAmount = CDec(Me.txtTaxAmount.Text)
                Else
                    TaxAmount = 0
                End If

                If IsNumeric(Me.txtTax.Text) Then
                    Tax = CDec(Me.txtTax.Text)
                Else
                    Tax = 0
                End If

                If IsNumeric(Me.txtTaxExchangeRate.Text) Then
                    TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
                Else
                    TaxExchangeRate = 0
                End If

                'hitung subtotal
                For Each rowdgv In Me.dgvItem.Rows

                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        Total = Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                        Rounding, MidpointRounding.AwayFromZero)
                    Else
                        Total = Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                        Rounding, MidpointRounding.AwayFromZero)
                    End If
                    rowdgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Total, Rounding, ClsCekData.EnumTipeRounding.Middlepoint)
                    SubTotal += Total
                Next
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                SubtotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                'hitung pajak

                Dim DefPrecTax As Integer = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.TxtPrimaryCurrency.Text = "1" Then
                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
                            SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(TxtPrimaryCurrency.Text))
                        Else
                            If Me.txtPrimaryCurrTax.Text = "1" Then
                                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * Tax * _
                                                      SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                            Else

                                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
                                                       SubtotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                            End If
                        End If
                    Else
                        TaxAmount = 0
                    End If
                Else
                    TaxAmount = 0
                End If
            Else
                Me.TxtDutystamp.Text = 0
                Me.txtTaxAmount.Text = 0
                Me.txtTax.Text = 0
            End If

            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Dutystamp, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(TaxExchangeRate, 2)
            Me.txtTax.Text = Me.ObjCD.DigitGrouping(Tax, 2)

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                decTotal = 0
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
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            Else
                decTotal = 0
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
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
                Else
                    decTotal = 0
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
                End If
            End If




            'hitung purchase price yang akan digunakan untuk jurnal closing
            AmountInIDR = 0
            SubTotal1 = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                Else
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                End If
            Next
            SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
                    rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = SubTotal - AmountInIDR
                Else
                    If SubTotal1 <> 0 Then
                        If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                            rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        Else
                            rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        End If
                        AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungTotal_Multi_MiddlePoint()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal, TaxAmount, Total, SubtotalAfterDP, Dutystamp, Tax, TaxExchangeRate As Decimal
        Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecIDR As Integer = 0
        Dim AmountInIDR As Decimal = 0
        Dim SubTotal1 As Decimal = 0

        Try
            DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    Dutystamp = CDec(Me.TxtDutystamp.Text)
                Else
                    Dutystamp = 0
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    TaxAmount = CDec(Me.txtTaxAmount.Text)
                Else
                    TaxAmount = 0
                End If

                If IsNumeric(Me.txtTax.Text) Then
                    Tax = CDec(Me.txtTax.Text)
                Else
                    Tax = 0
                End If

                If IsNumeric(Me.txtTaxExchangeRate.Text) Then
                    TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
                Else
                    TaxExchangeRate = 0
                End If

                'hitung subtotal
                For Each rowdgv In Me.dgvItem.Rows

                    If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        Total = Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                        Rounding, MidpointRounding.AwayFromZero)
                    Else
                        Total = Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                        CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                        Rounding, MidpointRounding.AwayFromZero)
                    End If
                    rowdgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Total, Rounding, ClsCekData.EnumTipeRounding.Middlepoint)
                    SubTotal += Total
                Next
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                SubtotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
                'hitung pajak

                Dim DefPrecTax As Integer = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.TxtPrimaryCurrency.Text = "1" Then
                            TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
                            SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), Me.ObjCurrency.GetDefaultPrecission(TxtPrimaryCurrency.Text))
                        Else
                            If Me.txtPrimaryCurrTax.Text = "1" Then
                                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * Tax * _
                                SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            Else

                                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
                                SubtotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            End If
                        End If
                    Else
                        TaxAmount = 0
                    End If
                Else
                    TaxAmount = 0
                End If
            Else
                Me.TxtDutystamp.Text = 0
                Me.txtTaxAmount.Text = 0
                Me.txtTax.Text = 0
            End If

            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Dutystamp, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(TaxExchangeRate, 2)
            Me.txtTax.Text = Me.ObjCD.DigitGrouping(Tax, 2)

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                decTotal = 0
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
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            Else
                decTotal = 0
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
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
                Else
                    decTotal = 0
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
                End If
            End If




            'hitung purchase price yang akan digunakan untuk jurnal closing
            AmountInIDR = 0
            SubTotal1 = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                Else
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                End If
            Next
            SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
                    rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = SubTotal - AmountInIDR
                Else
                    If SubTotal1 <> 0 Then
                        If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                            rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        Else
                            rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        End If
                        AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDataStockIn()
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.txtPrimaryStockInPurchase.Text = "" Or Me.txtPrimaryStockInPurchase.Text = "-1" Then
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "1"
                Me.txtGRNNo.Clear()
                Me.txtPrimaryGRN.Text = "-1"
                Me.txtAopNo.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "-1"
                Me.txtPrimaryCurrTax.Text = "-1"
                Me.txtPrimaryAccount.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtSupplierDONumber.Clear()
                Me.TxtTaxReturnNotes.Clear()
                Me.TxtSupplierNPWPNo.Clear()
                Me.TxtSupplierInvoiceNumber.Clear()
            Else
                Me.TxtStockInNo.Text = Me.objInOut.GetNoTransaksi(Me.txtPrimaryStockInPurchase.Text)
                '  Dim PrimaryDO As Double = Me.ObjStockInOut.GetPrimaryDO(Me.txtPrimaryStockIn.Text)
                '    Me.txtPrimaryDO.Text = PrimaryDO
                Me.objGRN.GetInfo(Me.txtPrimaryStockInPurchase.Text, Me.txtGRNNo.Text, Me.txtPrimaryGRN.Text, Me.txtAopNo.Text, _
                            Me.txtPrimarySupplier.Text, Me.txtSupplierCode.Text, Me.TxtSupplierName.Text, Me.txtSupplierDONumber.Text)

                If Me.txtAopNo.Text <> "" Then
                    Me.TxtSupplierInvoiceNumber.Text = Me.ObjAOP.GetSupplier_InvoiceNumber(Me.txtAopNo.Text)
                    'Me.TxtSupplierNPWPNo.Text = Me.ObjAOP.GetSupplier_NPWPNumber(Me.txtAopNo.Text)
                    'update NPWP 16 Digit.
                    If String.IsNullOrEmpty(Me.TxtSupplierNPWPNo.Text) = True Then
                        Me.TxtSupplierNPWPNo.Text = Me.ObjAOP.GetSupplier_NPWPNumber(Me.txtAopNo.Text)
                        '2025 Core Tax Tahap 2.1
                        If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                            Me.TxtSupplierNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtSupplierNPWPNo.Text)
                        End If
                    End If
                    Me.GetInfoPO()
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                Else
                    Me.TxtSupplierInvoiceNumber.Clear()
                    Me.TxtSupplierNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                    Me.GetInfoPO()
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    Me.IsiCBOArAkunByTipePerkiraan1()
                    Me.IsiCBOCurrencyTax()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub GetInfoPO()
        Dim dta As New DseTransaksiGRNTableAdapters.DataTable5TableAdapter
        Dim TBL As New DseTransaksiGRN.DataTable5DataTable
        Dim row As DseTransaksiGRN.DataTable5Row

        Try
            dta.Fill(TBL, Me.txtPrimaryGRN.Text)

            'Dim I As New Form5
            'I.DataGridView1.DataSource = Nothing
            'I.DataGridView1.DataSource = TBL
            'I.ShowDialog()

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.primary_aop <> -1 Then
                    Me.TxtPrimaryCurrency.Text = row.Primary_Currency_AOP
                Else
                    Me.TxtPrimaryCurrency.Text = row.Primary_Currency_PO
                End If
            End If

            'If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
            '    Me.cboAccount.Enabled = False
            '    Me.cboPrimaryAccount.Enabled = False
            '    Me.cboCurrTax.Enabled = False
            '    Me.cboPrimaryCurrTax.Enabled = False
            'Else
            '    Me.cboAccount.Enabled = True
            '    Me.cboPrimaryAccount.Enabled = True
            '    Me.cboCurrTax.Enabled = True
            '    Me.cboPrimaryCurrTax.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TBL.Dispose()
        End Try
    End Sub
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    Private Sub TxtStockOutNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtStockInNo.Enabled = True Then
                If Me.TxtStockInNo.ReadOnly = True Then
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtStockInNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockOutNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.LostFocus
        Try
            If Me.TxtStockInNo.Enabled = True Or Me.TxtRemark.Enabled = True Then
                If Me.TxtStockInNo.ReadOnly = True Then
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReasonOfReturn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.GotFocus
        TxtReasonOfReturn.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtReasonOfReturn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.LostFocus
        TxtReasonOfReturn.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
    End Sub


    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)

            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryPO.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"

            ''@rosadinnn
            ''@MUaT
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.TxtRatioDefault.Text = "1"
            Me.TxtRatioReference.Text = "1"

            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "0"

            ''@rosadinnn
            ''@MUaT
            Me.TxtQtyperBatchDefault.Text = "0"

            Me.txtSecondaryQty.Text = "0"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsMulti.Left = Me.lblSecondaryUnits.Left
            Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
            Me.lblSecQty.Visible = False
            Me.lblExpiryDate.Visible = False
            Me.DtpExpired.Value = Today
            Me.DtpExpired.Visible = True
            Me.TxtBatchNo.Visible = True
            Me.lblBatchNo.Visible = False
            Me.TxtBatchNo.Visible = True
            Me.TxtBatchNo.TabStop = True
            Me.TxtBatchNo.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '0. tidak ada stock in number yang diberikan
            If Me.TxtStockInNo.Text = "" Then
                MsgBox("Please specify stock in purchase number first", MsgBoxStyle.Information, "Information")
                Me.TxtStockInNo.Focus()
                Return False
            End If

            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If

            Dim blnKetemu As Boolean = False
            Dim rowDGV As New DataGridViewRow
            If Me.DGVGRN.Rows.Count > 0 Then
                For Each rowDGV In Me.DGVGRN.Rows
                    With rowDGV
                        If Me.txtPrimaryCodeNo.Text = .Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value Then
                            blnKetemu = True
                            Exit For
                        End If
                    End With
                Next
            End If

            If blnKetemu = False Then
                MsgBox("We can't found code no : " & Me.TxtCodeNo.Text & " in stock in purchase no : " & Me.TxtStockInNo.Text & ".", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If


            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If


            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True Then
                If Me.TxtBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtBatchNo.Focus()
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
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, Me.TxtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.TxtBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If

                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value = Me.TxtBatchNo.Text And Me.TxtBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.txtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiryDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.TxtBatchNo.Text + " on " + _
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

            'harus memperhitungkan stock jangan sampe minus

            ''@rosadinnn
            ''@MUaT
            If Me.TxtRatioDefault.Text = "" Or Me.TxtRatioDefault.Text = "0" Then
                MsgBox("Invalid ratio default.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.TxtRatioReference.Text = "" Or Me.TxtRatioReference.Text = "0" Then
                MsgBox("Invalid ratio reference.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.txtPrimaryUnitsDefault.Text = "" Or Me.txtPrimaryUnitsDefault.Text = "-1" Then
                MsgBox("Invalid primary unit default.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.txtPrimaryUnitsMulti.Text = "" Or Me.txtPrimaryUnitsMulti.Text = "-1" Then
                MsgBox("Invalid primary unit multi.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.TxtQtyperBatchDefault.Text = "" Or Me.TxtQtyperBatchDefault.Text = "0" Then
                MsgBox("Invalid qty default.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim rowdgv As DataGridViewRow
        Dim TotalQty As Double = 0

        Try
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    TotalQty += rowdgv.Cells(Me.ColTotalQty.Index).Value
                End If
            Next

            Return TotalQty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
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
                        If row.Cells(Me.ColSecQty.Index).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(16).Value Then
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
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.cboTipePurchaseReturn.SelectedIndex = 0
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVGRN.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MinimizeBox = True
            'Me.Text = "Quick Purchase Return"

            'Me.ColPriceByMultiUnit.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanStockOutPurchaseReturn()
        Dim StockOutNo As String = ""

        Try
            If Me.objPurchaseReturn.GenerateStockOutPurchaseReturn_Multi(Me.dgvItem, Me) = True Then
                StockOutNo = Me.objPurchaseReturn.GetNoTransaksiStockOut(Me.objPurchaseReturn.Primary)
                MsgBox("Stock Out #" + StockOutNo + " succesfully generated", _
                MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objPurchaseReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()

                            'konfirmasi simpan stock purchase return
                            If Me.objPurchaseReturn.ApakahBisaGenerateStockOutPurchaseReturn(Me.dgvItem, Me) = True Then
                                If MsgBox("Generate stock out purchase return automatically?", _
                                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Me.SimpanStockOutPurchaseReturn()
                                End If
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objPurchaseReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()

                            'konfirmasi simpan stock purchase return
                            If Me.objPurchaseReturn.ApakahBisaGenerateStockOutPurchaseReturn(Me.dgvItem, Me) = True Then
                                If MsgBox("Generate stock out purchase return automatically?", _
                                MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Me.SimpanStockOutPurchaseReturn()
                                End If
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

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.TxtCodeNo.Text = "" And Me.TxtBatchNo.Text = "" And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objPurchaseReturn.CariDetail(Me.TxtPrimaryPurchaseReturn.Text, Me.dgvItem, Me)
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

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiPurchaseReturn
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryPurchaseReturn = -1 Then
                If Me.objPurchaseReturn.CariDetail(ij.PrimaryPurchaseReturn, Me.dgvItem, Me) = True Then
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
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPurchaseReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPurchaseReturnNo.Text <> "" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                    MsgBox("Please edit this data from quick purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objPurchaseReturn.ApakahBolehEdit(Me.UserCode, True) = True Then
                    '1 baris ini ditambahkan, karena setiap kali print pr, ada perubahan timestamp
                    Me.objPurchaseReturn.TimeStamp = Me.objPurchaseReturn.GetTimestamp(Me.objPurchaseReturn.Primary)

                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPurchaseReturnNo.Text <> "" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                    MsgBox("Please delete this data from quick purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objPurchaseReturn.ApakahBolehHapus(Me.UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objPurchaseReturn.Hapus(Me.TxtPrimaryPurchaseReturn.Text, Me.UserCode, False) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objPurchaseReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
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
                            If Me.objPurchaseReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.txtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.txtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

            Me.TxtBatchNo.ReadOnly = False
            Me.TxtBatchNo.Visible = True

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
                Me.TxtBatchNo.ReadOnly = False
                Me.TxtBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.TxtBatchNo.Visible = True
            Else
                Me.TxtBatchNo.ReadOnly = True
                Me.TxtBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.TxtBatchNo.Visible = False
            End If

            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.txtSecondaryQty.Visible = False
                Me.lblSecondaryUnits.Visible = False
                'Me.lblUnitsMulti.Left = Me.lblSecondaryUnits.Left
                Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                'Me.lblUnitsMulti.Left = Me.lblSecQty.Left - 37
                Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                Me.txtSecondaryQty.Text = "0"
            End If
            If Me.txtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''@MUaT
                'Me.lblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)

                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnitsMulti.Text = "Units"
                Me.lblSecondaryUnits.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.txtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    Me.SetPONumber()
                    If Me.TxtBatchNo.Visible = True Then
                        Me.TxtBatchNo.Focus()
                    ElseIf Me.TxtBatchNo.Visible = True Then
                        Me.TxtBatchNo.Focus()
                    Else
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Me.SetPONumber()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.TxtBatchNo.Visible = True Then
                            Me.TxtBatchNo.Focus()
                        ElseIf Me.TxtBatchNo.Visible = True Then
                            Me.TxtBatchNo.Focus()
                        Else
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                        Me.SetPONumber()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
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

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct
        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.txtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.TxtBatchNo.Visible = True Then
                        Me.TxtBatchNo.Focus()
                    ElseIf Me.TxtBatchNo.Visible = True Then
                        Me.TxtBatchNo.Focus()
                    Else
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Me.SetPONumber()
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.TxtBatchNo.Visible = True Then
                            Me.TxtBatchNo.Focus()
                        ElseIf Me.TxtBatchNo.Visible = True Then
                            Me.TxtBatchNo.Focus()
                        Else
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                        Me.SetPONumber()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else

                ''@rosadinnn
                ''@MUaT
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""
                Me.txtPrimaryCategory.Text = "-1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtRatioReference.Text = "1"
                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.txtPrimaryUnitsMulti.Text = "-1"

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.SetGboItem()
                If Me.TxtBatchNo.Visible = True Then
                    Me.TxtBatchNo.Focus()
                ElseIf Me.TxtBatchNo.Visible = True Then
                    Me.TxtBatchNo.Focus()
                Else
                    Me.TxtQtyperBatchMulti.Focus()
                End If
                Me.SetPONumber()
            Else
                Me.TxtCodeNo.Clear()
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If

            Me.KontrolKolom()
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
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
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
        Dim RowDgv As DataGridViewRow
        Try
            If Me.ApakahDataItemBenar = True Then
                If Me.CekDGVItemDgnDGVGRNBtnInput_Multi(Me.txtPONo.Text, Me.TxtCodeNo.Text, Me.TxtBatchNo.Text, Me.DtpExpired.Value, Me.txtSecondaryQty.Text) = True Then
                    Dim Index As Integer = Me.GetIndex(Me.txtPrimaryCodeNo.Text, Me.TxtBatchNo.Text, _
                    Me.txtSecondaryQty.Text, Me.DtpExpired.Value)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        'tambahkan nilai qty
                        Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                        ''@rosadinnn
                        ''@MUaT
                        'row.Cells(Me.ColTotalQty.Index).Value = CDbl(row.Cells(Me.ColTotalQty.Index).Value) + CDbl(Me.TxtQtyperBatchMulti.Text)
                        row.Cells(Me.ColQtyMulti.Index).Value = CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(Me.TxtQtyperBatchMulti.Text)
                        Me.SetQtyDefaultOnDgv(row.Index)

                        Me.HitungTotal_Multi()
                    Else
                        'tambahkan baris baru
                        Dim ExpiredDateStr As String
                        If Me.DtpExpired.Value.Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                        End If

                        ''@rosadinnn
                        ''@MUaT
                        'Me.dgvItem.Rows.Add("0", Me.TxtCodeNo.Text, Me.txtDesc.Text, _
                        '   Me.txtPONo.Text, Me.TxtBatchNo.Text, Me.DtpExpired.Value, 0, 0, "", Me.TxtQtyperBatchMulti.Text, _
                        '   Me.TxtQtyperBatchMulti.Text, Me.lblUnitsMulti.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, _
                        '   Me.TxtRemarkItem.Text, Me.txtPrimaryCodeNo.Text, _
                        '   Me.txtPrimaryPO.Text, -1, Me.txtPrimaryUnitsDefault.Text, Me.txtPrimarySecondaryUnits.Text)
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With RowDgv
                            .Cells(Me.DataGridViewTextBoxColumn4.Index).Value = ""
                            .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                            .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                            .Cells(Me.ColPONO.Index).Value = Me.txtPONo.Text
                            .Cells(Me.ColBatchNo.Index).Value = Me.TxtBatchNo.Text
                            .Cells(Me.ColExpiryDate.Index).Value = Format(Me.DtpExpired.Value, "dd MMM yyyy")
                            .Cells(Me.ColQtyMaster.Index).Value = 0
                            .Cells(Me.ColQtyPerMaster.Index).Value = 0
                            .Cells(Me.ColMasterUnits.Index).Value = ""
                            .Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchMulti.Text
                            .Cells(Me.ColTotalQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                            .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                            .Cells(Me.ColSecQty.Index).Value = Me.txtSecondaryQty.Text
                            .Cells(Me.ColSecUnit.Index).Value = Me.lblSecondaryUnits.Text
                            .Cells(Me.colRemarks.Index).Value = Me.TxtRemarkItem.Text
                            .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                            .Cells(Me.ColPrimaryPO.Index).Value = Me.txtPrimaryPO.Text
                            .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                            .Cells(Me.ColPrimaryUnit.Index).Value = Me.txtPrimaryUnitsDefault.Text
                            .Cells(Me.ColPrimarySecUnit.Index).Value = Me.txtPrimarySecondaryUnits.Text

                            .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                            .Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                            If Me.txtPrimaryUnitsMulti.Text <> Me.txtPrimaryUnitsDefault.Text Then
                                .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                        ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) _
                                        & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                        Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) & " " & Me.lblUnitsDefault.Text
                            Else
                                .Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If
                            .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                            .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                            .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        End With

                        'set ulang nomor
                        Me.ObjCD.SetNumber(Me.dgvItem)
                        'isikan data harga+hitung total
                        Me.IsikanDataHarga()
                    End If

                    'kosongkan gbo item
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()
                    'format dgv
                    Me.FormatDGV()
                    Me.FormatTxtStockIN()
                    'fokuskan ke baris terakhir
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGVGRN()
        Dim row As DataGridViewRow

        Try
            Dim AdaSec As Boolean = False
            Dim AdaPO As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaBatchNo As Boolean = False

            For Each row In Me.DGVGRN.Rows
                row.Cells(Me.colQTYDGVGRN.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.colQTYDGVGRN.Index).Value)
                row.Cells(Me.colQtyMasterDGVGRN.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.colQtyMasterDGVGRN.Index).Value)
                row.Cells(Me.colSecQtyDGVGRN.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.colSecQtyDGVGRN.Index).Value)
                row.Cells(Me.colQtyPerMasterDGVGRN.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.colQtyPerMasterDGVGRN.Index).Value)

                ''@rosadinnn
                ''@MUaT
                'row.Cells(Me.colTotalQtyDGVGRN.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.colTotalQtyDGVGRN.Index).Value)

                If row.Cells(Me.colSecQtyDGVGRN.Index).Value <> -1 Then
                    AdaSec = True
                End If
                If row.Cells(Me.colPrimaryPODGVGRN.Index).Value <> -1 Then
                    AdaPO = True
                End If

                If IsDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value) = True Then
                    If row.Cells(Me.colExpiryDateDGVGRN.Index).Value <> #1/1/1950# Then
                        AdaExpiry = True
                        row.Cells(Me.colExpiryDateDGVGRN.Index).Value = Format(CDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value), "dd MMM yyyy")
                    Else
                        row.Cells(Me.colExpiryDateDGVGRN.Index).Value = "-"
                    End If
                End If
                If row.Cells(Me.colBatchNoDGVGRN.Index).Value <> "" Then
                    AdaBatchNo = True
                End If
                If IsNumeric(row.Cells(Me.colSecQtyDGVGRN.Index).Value) = False Then
                    row.Cells(Me.colSecQtyDGVGRN.Index).Value = "-"
                    row.Cells(Me.colSecUnitDGVGRN.Index).Value = "-"
                Else
                    If row.Cells(Me.colSecQtyDGVGRN.Index).Value = "-1" Then
                        row.Cells(Me.colSecQtyDGVGRN.Index).Value = "-"
                        row.Cells(Me.colSecUnitDGVGRN.Index).Value = "-"
                    End If
                End If
            Next

            If AdaSec = True Then
                Me.colSecQtyDGVGRN.Visible = True
                Me.colSecUnitDGVGRN.Visible = True
            Else
                Me.colSecQtyDGVGRN.Visible = False
                Me.colSecUnitDGVGRN.Visible = False
            End If

            If AdaPO = True Then
                Me.colPONoDGVGRN.Visible = True
            Else
                Me.colPONoDGVGRN.Visible = False
            End If
            If AdaExpiry = True Then
                Me.colExpiryDateDGVGRN.Visible = True
            Else
                Me.colExpiryDateDGVGRN.Visible = False
            End If
            If AdaBatchNo = True Then
                Me.colBatchNoDGVGRN.Visible = True
            Else
                Me.colBatchNoDGVGRN.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim IsMultiUnit As Boolean = False

        Try
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
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecQty.Index).Value = "-"
                    row.Cells(Me.ColSecUnit.Index).Value = "-"
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = True
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

            ''@rosadinnn
            ''@MUaT
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnit.Index).Value <> .Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        IsMultiUnit = True
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

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
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
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColExpiryDate.Index Or e.ColumnIndex = Me.ColSecQty.Index Or _
                e.ColumnIndex = Me.ColBatchNo.Index Or e.ColumnIndex = Me.ColQtyMulti.Index Or _
                e.ColumnIndex = Me.colRemarks.Index Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = Me.ColSecQty.Index Then
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
                        e.Cancel = True
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    If e.ColumnIndex = Me.ColQtyMulti.Index Then
                        Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value)
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
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim row As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If e.RowIndex <> -1 Then
                    If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value Or _
                            Me.TxtRemark.Enabled = False Then
                        Exit Sub
                    End If

                    Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    Dim BufferRatioDesc As String = Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value

                    frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
                    frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value
                    frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                        If Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value Then
                            Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value & " " & Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value & " = " & _
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value & " " & Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                        Else
                            Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = "-"
                        End If
                        row = Me.dgvItem.Rows(e.RowIndex)
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                        'Me.Hitung(row)
                        If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                            MsgBox("Ratio value restored to previous value, because ratio you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                            Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                            Me.SetQtyDefaultOnDgv(e.RowIndex)
                        End If
                        Me.HitungTotal_Multi()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim SecondaryQty As Double
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            Dim rounding As Integer = 0

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If IsDate(Me.dgvItem.Rows(e.RowIndex).Cells(Me.ColExpiryDate.Index).Value) = True Then
                        ExpiryDate = Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
                    Else
                        ExpiryDate = #1/1/1950#
                    End If
                    If IsNumeric(Me.dgvItem.Item(Me.ColSecQty.Index, e.RowIndex).Value) = True Then
                        SecondaryQty = Me.dgvItem.Item(Me.ColSecQty.Index, e.RowIndex).Value
                    Else
                        SecondaryQty = -1
                    End If

                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColQtyMulti.Index Then
                        'pastikan bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Qty data must be a positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If

                        'digit grouoping

                        ''@rosadinnn
                        ''@MUaT
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping( _
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                                Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))

                        'pastikan stock tidak menjadi minus

                        'berhubungan dengan batasan minus terhadap GRN

                        ''@rosadinnn
                        ''@MUaT
                        ''ganti kode unit dg kode unit multi
                        If Me.CekDGVItemDgnDGVGRNCellEndEdit_Multi(Me.dgvItem.Item(Me.ColPONO.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, _
                                ExpiryDate, SecondaryQty, Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Try
                        Else
                            rounding = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)

                            ''@rosadinnn
                            ''@MUaT
                            Me.SetQtyDefaultOnDgv(e.RowIndex)
                            If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = 1 Then
                                Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
                                CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value), rounding))
                            Else
                                Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
                                CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value), rounding))
                            End If
                            Me.HitungTotal_Multi()
                        End If

                        If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                            MsgBox("Qty restored to previous value, because qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.SetQtyDefaultOnDgv(row.Index)
                        End If

                    ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                        'pastikan upper case
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        'pastikan ada batch number pada dgv grn
                        If Me.CekDGVItemDgnDGVGRNCellEndEdit_Multi(Me.dgvItem.Item(Me.ColPONO.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, _
                                ExpiryDate, SecondaryQty, Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Try
                        Else
                            Me.HitungTotal_Multi()
                        End If
                    End If

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
                Else
                    If e.ColumnIndex = Me.colRemarks.Index Then
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

    Private Function CekDGVItemDgnDGVGRNBtnInput_Default(ByVal vNoPO As String, ByVal vCodeNo As String, _
    ByVal BatchNo As String, ByVal ExpiryDate As Date, ByVal SecQty As Double) As Boolean
        Try
            Dim row As DataGridViewRow
            Dim JmlDGVItem, JmlDGVGRN As Double
            Dim ExpiryDate1 As Date
            Dim SecQty1 As Double = 0

            JmlDGVItem = 0
            JmlDGVGRN = 0
            Dim Ketemu As Boolean = False

            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate1 = #1/1/1950#
                Else
                    ExpiryDate1 = row.Cells(Me.ColExpiryDate.Index).Value
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.ColCodeNo.Index).Value = vCodeNo And row.Cells(Me.ColPONO.Index).Value = vNoPO And _
                ExpiryDate1 = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.ColBatchNo.Index).Value = BatchNo Then
                    JmlDGVItem += CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next
            For Each row In Me.DGVGRN.Rows
                If IsDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value) = False Then
                    ExpiryDate1 = #1/1/1950#
                Else
                    ExpiryDate1 = row.Cells(Me.colExpiryDateDGVGRN.Index).Value
                End If
                If IsNumeric(row.Cells(Me.colSecQtyDGVGRN.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.colSecQtyDGVGRN.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.colCodeNoDGVGRN.Index).Value = vCodeNo And row.Cells(Me.colPONoDGVGRN.Index).Value = vNoPO And _
                ExpiryDate1 = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.colBatchNoDGVGRN.Index).Value = BatchNo Then
                    JmlDGVGRN += CDbl(row.Cells(Me.colTotalQtyDGVGRN.Index).Value)
                    Ketemu = True
                End If
            Next
            If Ketemu = False Then
                MsgBox("Can not find the item you want to input at GRN Item", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If JmlDGVItem + CDbl(Me.TxtQtyperBatchMulti.Text) > JmlDGVGRN Then
                MsgBox("Maximum qty for this item is " + Me.ObjCD.DigitGrouping(CStr(JmlDGVGRN - JmlDGVItem)) + " " + Me.lblUnitsMulti.Text, _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function CekDGVItemDgnDGVGRNBtnInput_Multi(ByVal vNoPO As String, ByVal vCodeNo As String, _
    ByVal BatchNo As String, ByVal ExpiryDate As Date, ByVal SecQty As Double) As Boolean
        Try
            Dim row As DataGridViewRow
            Dim JmlDGVItem, JmlDGVGRN As Double
            Dim ExpiryDate1 As Date
            Dim SecQty1 As Double = 0

            JmlDGVItem = 0
            JmlDGVGRN = 0
            Dim Ketemu As Boolean = False

            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate1 = #1/1/1950#
                Else
                    ExpiryDate1 = row.Cells(Me.ColExpiryDate.Index).Value
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.ColCodeNo.Index).Value = vCodeNo And row.Cells(Me.ColPONO.Index).Value = vNoPO And _
                ExpiryDate1 = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.ColBatchNo.Index).Value = BatchNo Then
                    JmlDGVItem += CDbl(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next
            For Each row In Me.DGVGRN.Rows
                If IsDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value) = False Then
                    ExpiryDate1 = #1/1/1950#
                Else
                    ExpiryDate1 = row.Cells(Me.colExpiryDateDGVGRN.Index).Value
                End If
                If IsNumeric(row.Cells(Me.colSecQtyDGVGRN.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.colSecQtyDGVGRN.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.colCodeNoDGVGRN.Index).Value = vCodeNo And row.Cells(Me.colPONoDGVGRN.Index).Value = vNoPO And _
                ExpiryDate1 = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.colBatchNoDGVGRN.Index).Value = BatchNo Then
                    JmlDGVGRN += CDbl(row.Cells(Me.ColQtyMulti_GRN.Index).Value)
                    Ketemu = True
                End If
            Next
            If Ketemu = False Then
                MsgBox("Can not find the item you want to input at GRN Item", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If JmlDGVItem + CDbl(Me.TxtQtyperBatchMulti.Text) > JmlDGVGRN Then
                MsgBox("Maximum qty for this item is " + Me.ObjCD.DigitGrouping(CStr(JmlDGVGRN - JmlDGVItem)) + " " + Me.lblUnitsMulti.Text, _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekDGVItemDgnDGVGRNCellEndEdit_Default(ByVal vNoPO As String, ByVal vCodeNo As String, _
    ByVal BatchNo As String, ByVal ExpiryDate As Date, ByVal SecQty As Double, ByVal UnitCode As String) As Boolean
        Try
            Dim row As DataGridViewRow
            Dim JmlDGVItem, JmlDGVGRN As Double
            JmlDGVItem = 0
            JmlDGVGRN = 0
            Dim ED As Date = #1/1/1950#
            Dim SecQty1 As Double = -1
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    ED = row.Cells(Me.ColExpiryDate.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.ColCodeNo.Index).Value = vCodeNo And row.Cells(Me.ColPONO.Index).Value = vNoPO And _
                ED = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.ColBatchNo.Index).Value = BatchNo Then
                    JmlDGVItem += CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            For Each row In Me.DGVGRN.Rows
                If IsDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value) = True Then
                    ED = row.Cells(Me.colExpiryDateDGVGRN.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                If IsNumeric(row.Cells(Me.colSecQtyDGVGRN.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.colSecQtyDGVGRN.Index).Value
                Else
                    SecQty1 = -1
                End If

                If row.Cells(Me.colCodeNoDGVGRN.Index).Value = vCodeNo And row.Cells(Me.colPONoDGVGRN.Index).Value = vNoPO And _
                ED = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.colBatchNoDGVGRN.Index).Value = BatchNo Then
                    JmlDGVGRN += CDbl(row.Cells(Me.colTotalQtyDGVGRN.Index).Value)
                End If
            Next


            If JmlDGVItem > JmlDGVGRN Then
                MsgBox("Maximum return qty for this item is " + Me.ObjCD.DigitGrouping(CStr(JmlDGVGRN)) + _
                " " + UnitCode, MsgBoxStyle.Information, "Information")

                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function CekDGVItemDgnDGVGRNCellEndEdit_Multi(ByVal vNoPO As String, ByVal vCodeNo As String, _
    ByVal BatchNo As String, ByVal ExpiryDate As Date, ByVal SecQty As Double, ByVal UnitCode As String) As Boolean
        Try
            Dim row As DataGridViewRow
            Dim JmlDGVItem, JmlDGVGRN As Double
            JmlDGVItem = 0
            JmlDGVGRN = 0
            Dim ED As Date = #1/1/1950#
            Dim SecQty1 As Double = -1
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    ED = row.Cells(Me.ColExpiryDate.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty1 = -1
                End If
                If row.Cells(Me.ColCodeNo.Index).Value = vCodeNo And row.Cells(Me.ColPONO.Index).Value = vNoPO And _
                ED = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.ColBatchNo.Index).Value = BatchNo Then
                    JmlDGVItem += CDbl(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            For Each row In Me.DGVGRN.Rows
                If IsDate(row.Cells(Me.colExpiryDateDGVGRN.Index).Value) = True Then
                    ED = row.Cells(Me.colExpiryDateDGVGRN.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                If IsNumeric(row.Cells(Me.colSecQtyDGVGRN.Index).Value) = True Then
                    SecQty1 = row.Cells(Me.colSecQtyDGVGRN.Index).Value
                Else
                    SecQty1 = -1
                End If

                If row.Cells(Me.colCodeNoDGVGRN.Index).Value = vCodeNo And row.Cells(Me.colPONoDGVGRN.Index).Value = vNoPO And _
                ED = ExpiryDate And SecQty1 = SecQty And row.Cells(Me.colBatchNoDGVGRN.Index).Value = BatchNo Then
                    JmlDGVGRN += CDbl(row.Cells(Me.ColQtyMulti_GRN.Index).Value)
                End If
            Next

            If JmlDGVItem > JmlDGVGRN Then
                MsgBox("Maximum return qty for this item is " + Me.ObjCD.DigitGrouping(CStr(JmlDGVGRN)) + _
                " " + UnitCode, MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaBarisYangSama() As Boolean
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

                    If row.Cells(Me.ColPONO.Index).Value = row1.Cells(Me.ColPONO.Index).Value And _
                    row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And _
                    row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = row1.Cells(Me.ColSecQty.Index).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = False Then
                            Return True
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
    Private Sub GabungkanBaris()
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
                    If CStr(row.Cells(Me.ColPONO.Index).Value) = CStr(row1.Cells(Me.ColPONO.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = True Then
                            If CDbl(row.Cells(Me.ColSecQty.Index).Value) = CDbl(row1.Cells(Me.ColSecQty.Index).Value) Then
                                Me.GabungkanBaris(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = False Then
                            Me.GabungkanBaris(row, row1)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                            Exit Sub
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
            row.Cells(Me.ColTotalQty.Index).Value = CDbl(row.Cells(Me.ColTotalQty.Index).Value) + CDbl(row1.Cells(Me.ColTotalQty.Index).Value)
            'gabungkan data remarks
            If CStr(row.Cells(Me.colRemarks.Index).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + ". "
            End If
            row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + CStr(row1.Cells(Me.colRemarks.Index).Value)
            'hapus baris ybs
            Me.dgvItem.Rows.RemoveAt(row1.Index)
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
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.dgvItem.Rows.Count = 0 Then
                Me.FormatTxtStockIN()
            End If
            Me.HitungTotal_Multi()

            ''@rosadinnn
            ''@MUaT
            Me.FormatDGV()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            Me.TxtRemark.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatchNo.GotFocus
        Try
            Me.TxtBatchNo.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBatchNo.LostFocus
        Try
            Me.TxtBatchNo.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
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
            Me.txtSecondaryQty.BackColor = Me.txtPrimaryCodeNo.BackColor
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
            Me.TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.GotFocus
        Try
            If Me.txtPONo.Enabled = True And Me.txtPONo.ReadOnly = False And Me.TxtRemark.Enabled = True Then
                Me.txtPONo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPONo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPONo.KeyDown
        Dim i As New FrmCariTransaksiPO
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONo.Text) = True Then
                    Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONo.Text)
                    Me.TxtCodeNo.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtPONo.Text <> "" Then
                        i.txtCari.Text = Me.txtPONo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryPO > 0 Then
                        Me.txtPrimaryPO.Text = i.PrimaryPO
                        Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.txtPONo.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtPONo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtPONo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.Leave
        Dim i As New FrmCariTransaksiPO

        Try
            If Me.txtPONo.Text <> "" Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONo.Text) = True Then
                    Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONo.Text)
                    Me.TxtCodeNo.Focus()
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtPONo.Text <> "" Then
                        i.txtCari.Text = Me.txtPONo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        Me.txtPrimaryPO.Text = i.PrimaryPO
                        Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.txtPONo.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtPONo.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryPO.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        Try
            i.BackColorAll = Me.BackColorAll

            If Me.txtPONo.Text <> "" Then
                i.txtCari.Text = Me.txtPONo.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.txtPrimaryPO.Text = i.PrimaryPO
                Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                Me.TxtCodeNo.Focus()
            Else
                Me.txtPONo.Clear()
                Me.txtPrimaryPO.Text = "-1"
                Me.txtPONo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.LostFocus
        Try
            If Me.txtPONo.Enabled = True And Me.txtPONo.ReadOnly = False Then
                Me.txtPONo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objPurchaseReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objPurchaseReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindStockInPurchase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindStockInPurchase.Click
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            If Me.TxtStockInNo.Text <> "" Then
                i.txtCari.Text = Me.TxtStockInNo.Text
            End If
            i.ShowDialog()

            If i.PrimaryTransaksiStockInOut > 0 Then
                'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                Me.txtPrimaryStockInPurchase.Text = i.PrimaryTransaksiStockInOut
                Me.IsiDataStockIn()
                Me.IsikanDGV_Multi()
                Me.IsiDGVItem_Multi()
                'kalo ada item maka stock in no tidak bisa diubah
                Me.FormatTxtStockIN()
                Me.TxtSupplierNPWPNo.Focus()
            Else
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.IsiDataStockIn()
                Me.TxtStockInNo.Focus()
                Me.dgvItem.Rows.Clear()
                Me.DGVGRN.Rows.Clear()
                Me.IsikanDataHarga()
            End If


            Me.cboTipePurchaseReturn.SelectedIndex = 0
            If Me.txtPrimaryStockInPurchase.Text <> "-1" And Me.txtAopNo.Text <> "" Then
                If Format(Me.ObjAOP.GetTanggalAOP(Me.ObjAOP.GetPrimary(Me.txtAopNo.Text)), "dd MMM yyyy") <> _
                        Format(Me.dtpDate.Value, "dd MMM yyyy") Then
                    Me.cboTipePurchaseReturn.SelectedIndex = 1
                End If
            Else
                Me.cboTipePurchaseReturn.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FormatTxtStockIN()
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtStockInNo.ReadOnly = True
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtStockInNo.TabStop = False
                    Me.btnFindStockInPurchase.Enabled = False
                Else
                    Me.TxtStockInNo.ReadOnly = False
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.btnFindStockInPurchase.Enabled = True
                    Me.TxtStockInNo.TabStop = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockInNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockInNo.KeyDown
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

        Try
            If e.KeyValue = 13 Then
                Dim PRimary As Double = Me.objInOut.GetPrimary(Me.TxtStockInNo.Text)
                If PRimary > -1 Then
                    Me.txtPrimaryStockInPurchase.Text = PRimary
                    Me.IsiDataStockIn()
                    Me.IsikanDGV_Multi()
                    Me.IsiDGVItem_Multi()
                    'kalo ada item maka stock in no tidak bisa diubah
                    Me.FormatTxtStockIN()
                    'Me.TxtSupplierNPWPNo.Focus()
                Else
                    Me.btnFindStockInPurchase.PerformClick()
                End If
            End If

            Me.cboTipePurchaseReturn.SelectedIndex = 0
            If Me.txtPrimaryStockInPurchase.Text <> "-1" And Me.txtAopNo.Text <> "" Then
                If Format(Me.ObjAOP.GetTanggalAOP(Me.ObjAOP.GetPrimary(Me.txtAopNo.Text)), "dd MMM yyyy") <> _
                Format(Me.dtpDate.Value, "dd MMM yyyy") Then
                    Me.cboTipePurchaseReturn.SelectedIndex = 1
                End If
            Else
                Me.cboTipePurchaseReturn.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtStockInNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.Leave
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
        Try
            If Me.TxtStockInNo.Text <> "" Then
                Dim PRimary As Double = Me.objInOut.GetPrimary(Me.TxtStockInNo.Text)

                If PRimary > -1 Then
                    Me.txtPrimaryStockInPurchase.Text = PRimary
                    Me.IsiDataStockIn()
                    Me.IsikanDGV_Multi()
                    Me.IsiDGVItem_Multi()
                    'kalo ada item maka stock in no tidak bisa diubah
                    Me.FormatTxtStockIN()
                Else
                    Me.btnFindStockInPurchase.PerformClick()
                End If
            Else
                Me.txtPrimaryStockInPurchase.Text = -1
                Me.IsiDataStockIn()
                Me.DGVGRN.Rows.Clear()
            End If

            Me.cboTipePurchaseReturn.SelectedIndex = 0
            If Me.txtPrimaryStockInPurchase.Text <> "-1" And Me.txtAopNo.Text <> "" Then
                If Format(Me.ObjAOP.GetTanggalAOP(Me.ObjAOP.GetPrimary(Me.txtAopNo.Text)), "dd MMM yyyy") <> _
                Format(Me.dtpDate.Value, "dd MMM yyyy") Then
                    Me.cboTipePurchaseReturn.SelectedIndex = 1
                End If
            Else
                Me.cboTipePurchaseReturn.SelectedIndex = 0
            End If

            Me.KontrolKolom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

  

    ''' <summary>
    ''' data grn, yang sudah di kurangkan dengan purchase return yang lain
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub IsikanDGV_Default()
        Dim row As DseTransaksiGRN.DataTable2_MultiRow
        Dim TBL As New DseTransaksiGRN.DataTable2_MultiDataTable
        Dim QtyDGV As Double = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""

        Try
            Me.DGVGRN.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DGVGRN.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Dim strR As String = ""

            objGRN.Isikantblantaraviewgrnnpurchasereturn_view(TBL, Me.txtPrimaryGRN.Text, Me.TxtPrimaryPurchaseReturn.Text, strR)
            Me.DGVGRN.Rows.Clear()

            If strR <> "" Then
                MsgBox("Stock in no. " & Me.TxtStockInNo.Text & " is already finish returned by purchase return no : " & strR & ".", MsgBoxStyle.Information, "Information")
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "1"
                Me.txtGRNNo.Clear()
                Me.txtPrimaryGRN.Text = "-1"
                Me.txtAopNo.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "-1"
                Me.txtPrimaryCurrTax.Text = "-1"
                Me.txtPrimaryAccount.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtSupplierDONumber.Clear()
                Me.TxtTaxReturnNotes.Clear()
                Me.TxtSupplierNPWPNo.Clear()
                Me.TxtSupplierInvoiceNumber.Clear()
                Exit Sub
            End If



            For Each row In TBL.Rows
                If row.Primary_PO <> PrimaryPO Then
                    PrimaryPO = row.Primary_PO
                    PO_Number = Me.ObjPO.GetNoTransaksiPO(row.Primary_PO)
                End If
                row.PO_Number = PO_Number

                'untuk yang edit, maka qty yang dimasukkan ke dgv GRN adalah qty ditambahakan qty pada dgv
                If Me.baru = False Then
                    QtyDGV = Me.GetJml_Multi_MasihDefault(row.Primary_Product, row.Primary_PO, row.Batch_No, row.Expiry_Date, row.Secondary_Qty)
                End If

                Dim ExpiryDate As String = Format(row.Expiry_Date, "dd MMM yyyy")
                If row.Total_Qty <> 0 Then
                    Me.DGVGRN.Rows.Add("0", row.Product_Code, row.Product_Desc, "-1", "-1", _
                    row.PO_Number, row.Batch_No, ExpiryDate, "0", "0", "-1", "0", ObjCD.DigitGrouping(row.Total_Qty, row._Decimal), _
                    row.Unit_Code, row.Secondary_Qty, row.Sec_Unit, "", _
                    row.Primary_Product, "-1", row.Primary_Units, row.Primary_Secondary_Units, _
                    row.Primary_PO)
                End If
            Next

            ObjCD.SetNumber(Me.DGVGRN)
            Me.FormatDGVGRN()

            Me.DGVGRN.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsikanDGV_Multi()
        Dim row As DseTransaksiGRN.DataTable2_MultiRow
        Dim TBL As New DseTransaksiGRN.DataTable2_MultiDataTable
        Dim QtyDGV As Double = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""

        Dim RowDgv As DataGridViewRow

        Try
            Me.DGVGRN.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DGVGRN.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Dim strR As String = ""
            objGRN.Isikantblantaraviewgrnnpurchasereturn_view(TBL, Me.txtPrimaryGRN.Text, Me.TxtPrimaryPurchaseReturn.Text, strR)
            Me.DGVGRN.Rows.Clear()

            If strR <> "" Then
                MsgBox("Stock in no. " & Me.TxtStockInNo.Text & " is already finish returned by purchase return no : " & strR & ".", MsgBoxStyle.Information, "Information")
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "1"
                Me.txtGRNNo.Clear()
                Me.txtPrimaryGRN.Text = "-1"
                Me.txtAopNo.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.TxtPrimaryCurrency.Text = "-1"
                Me.txtPrimaryCurrTax.Text = "-1"
                Me.txtPrimaryAccount.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtSupplierDONumber.Clear()
                Me.TxtTaxReturnNotes.Clear()
                Me.TxtSupplierNPWPNo.Clear()
                Me.TxtSupplierInvoiceNumber.Clear()
                Exit Sub
            End If

            For Each row In TBL.Rows
                If row.Primary_PO <> PrimaryPO Then
                    PrimaryPO = row.Primary_PO
                    PO_Number = Me.ObjPO.GetNoTransaksiPO(row.Primary_PO)
                End If
                row.PO_Number = PO_Number

                'untuk yang edit, maka qty yang dimasukkan ke dgv GRN adalah qty ditambahakan qty pada dgv
                If Me.baru = False Then
                    QtyDGV = Me.GetJml_Multi_MasihDefault(row.Primary_Product, row.Primary_PO, row.Batch_No, row.Expiry_Date, row.Secondary_Qty)
                End If

                Dim ExpiryDate As String = Format(row.Expiry_Date, "dd MMM yyyy")
                If row.Total_Qty <> 0 Then
                    ''@rosadinnn
                    ''@MUaT
                    Me.DGVGRN.Rows.Add()
                    RowDgv = Me.DGVGRN.Rows(Me.DGVGRN.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColNo_GRN.Index).Value = ""
                        .Cells(Me.colCodeNoDGVGRN.Index).Value = row.Product_Code
                        .Cells(Me.colDescDGVGRN.Index).Value = row.Product_Desc
                        .Cells(Me.ColSuppCodeNoDGVGRN.Index).Value = "-1"
                        .Cells(Me.ColSuppItemDescDGVGRN.Index).Value = "-1"
                        .Cells(Me.colPONoDGVGRN.Index).Value = row.PO_Number
                        .Cells(Me.colBatchNoDGVGRN.Index).Value = row.Batch_No
                        .Cells(Me.colExpiryDateDGVGRN.Index).Value = ExpiryDate
                        .Cells(Me.colQtyMasterDGVGRN.Index).Value = "0"
                        .Cells(Me.colQtyPerMasterDGVGRN.Index).Value = "0"
                        .Cells(Me.colMasterUnitDGVGRN.Index).Value = "-1"
                        .Cells(Me.colQTYDGVGRN.Index).Value = "0"
                        .Cells(Me.colTotalQtyDGVGRN.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, row._Decimal)
                        .Cells(Me.colUnitDGVGRN.Index).Value = row.Unit_Code
                        .Cells(Me.colSecQtyDGVGRN.Index).Value = row.Secondary_Qty
                        .Cells(Me.colSecUnitDGVGRN.Index).Value = row.Sec_Unit
                        .Cells(Me.colRemarksDGVGRN.Index).Value = ""
                        .Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value = row.Primary_Product
                        .Cells(Me.colPRimaryMasterUnitDGVGRN.Index).Value = "-1"
                        .Cells(Me.colPrimaryUnitDGVGRN.Index).Value = row.Primary_Units
                        .Cells(Me.colPrimarySecUnitDGVGRN.Index).Value = row.Primary_Secondary_Units
                        .Cells(Me.colPrimaryPODGVGRN.Index).Value = row.Primary_PO
                        .Cells(Me.ColQtyMulti_GRN.Index).Value = ObjCD.DigitGrouping(row.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(Me.ColUnitsCodeMulti_GRN.Index).Value = row.Unit_Code_Multi
                        If row.Primary_Units <> row.Primary_Unit_Multi Then
                            .Cells(Me.ColRatioDesc_GRN.Index).Value = Me.ObjCD.DigitGrouping(row.Ratio_Reference, _
                                    ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                    & " " & row.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(row.Ratio_Default, _
                                    Me.ObjProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                        Else
                            .Cells(Me.ColRatioDesc_GRN.Index).Value = "-"
                        End If
                        .Cells(Me.ColRatioReference_GRN.Index).Value = row.Ratio_Reference
                        .Cells(Me.ColRatioDefault_GRN.Index).Value = row.Ratio_Default
                        .Cells(Me.ColPrimaryUnitMulti_GRN.Index).Value = row.Primary_Unit_Multi
                    End With
                End If
            Next

            ObjCD.SetNumber(Me.DGVGRN)
            Me.FormatDGVGRN()

            Me.DGVGRN.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' menyalin data dari dgv grn ke dalam dgv item
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub IsiDGVItem_Default()
        Dim rowDGV1 As DataGridViewRow
        Dim Clone As DataGridViewRow

        Dim TipeExcRate As ClsTransaksiAOP.EnumCalculatesBy = ClsTransaksiAOP.EnumCalculatesBy.AOP
        Dim DTA_AOP As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim TBL_AOP As New DseTransaksiAOP.tbltrans_aop_item_exchangerateDataTable
        Dim vPrimaryAOP As Double = -1
        Dim Rounding As Integer = 0
        Dim RoundingAOP As Integer = 0
        Dim vPrimaryCurrency, vTaxExchangeRate, vTaxPercent As Double
        Dim vUseTax As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.DataTable13TableAdapter
        Dim TBL1 As New DseTransaksiAOP.DataTable13DataTable
        Dim rowTBL As DseTransaksiAOP.DataTable13Row
        Dim Filter As String = ""
        Dim row_dgv As DataGridViewRow

        Try
            Me.dgvItem.Rows.Clear()

            For Each rowDGV1 In Me.DGVGRN.Rows
                Me.dgvItem.Rows.Add()
                Clone = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                With Clone
                    .Cells(Me.ColCodeNo.Index).Value = rowDGV1.Cells(Me.colCodeNoDGVGRN.Index).Value
                    .Cells(Me.ColDescription.Index).Value = rowDGV1.Cells(Me.colDescDGVGRN.Index).Value
                    .Cells(Me.ColPONO.Index).Value = rowDGV1.Cells(Me.colPONoDGVGRN.Index).Value
                    .Cells(Me.ColBatchNo.Index).Value = rowDGV1.Cells(Me.colBatchNoDGVGRN.Index).Value
                    .Cells(Me.ColExpiryDate.Index).Value = rowDGV1.Cells(Me.colExpiryDateDGVGRN.Index).Value
                    .Cells(Me.ColQtyMaster.Index).Value = rowDGV1.Cells(Me.colQtyPerMasterDGVGRN.Index).Value
                    .Cells(Me.ColQtyPerMaster.Index).Value = rowDGV1.Cells(Me.colQtyPerMasterDGVGRN.Index).Value
                    .Cells(Me.ColMasterUnits.Index).Value = rowDGV1.Cells(Me.colMasterUnitDGVGRN.Index).Value
                    .Cells(Me.ColQty.Index).Value = rowDGV1.Cells(Me.colQTYDGVGRN.Index).Value
                    .Cells(Me.ColTotalQty.Index).Value = rowDGV1.Cells(Me.colTotalQtyDGVGRN.Index).Value
                    .Cells(Me.ColUnit.Index).Value = rowDGV1.Cells(Me.colUnitDGVGRN.Index).Value
                    .Cells(Me.ColSecQty.Index).Value = rowDGV1.Cells(Me.colSecQtyDGVGRN.Index).Value
                    .Cells(Me.ColSecUnit.Index).Value = rowDGV1.Cells(Me.colSecUnitDGVGRN.Index).Value
                    .Cells(Me.colRemarks.Index).Value = ""
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGV1.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value
                    .Cells(Me.ColPrimaryPO.Index).Value = rowDGV1.Cells(Me.colPrimaryPODGVGRN.Index).Value
                    .Cells(Me.ColPrimaryMasterUnits.Index).Value = rowDGV1.Cells(Me.colPRimaryMasterUnitDGVGRN.Index).Value
                    .Cells(Me.ColPrimaryUnit.Index).Value = rowDGV1.Cells(Me.colPrimaryUnitDGVGRN.Index).Value
                    .Cells(Me.ColPrimarySecUnit.Index).Value = rowDGV1.Cells(Me.colPrimarySecUnitDGVGRN.Index).Value
                End With
            Next

            'isikan data harga pada item
            DTA.Fill(TBL1, Me.txtPrimaryGRN.Text)
            vPrimaryAOP = Me.ObjAOP.GetPrimaryAOP_ByPrimaryGRN(Me.txtPrimaryGRN.Text)
            DTA_AOP.FillBy_PrimaryAOP(TBL_AOP, vPrimaryAOP)
            'isikan data currency, tipe pajak, tax percent, tax exchannge rate
            If vPrimaryAOP = -1 Then
                Me.TxtPrimaryCurrency.Text = -1
                Me.TxtCurrency.Text = ""
                Me.chkTax.Checked = False
                Me.txtTax.Text = 0
                Me.txtTaxAmount.Text = 0
                Me.txtTaxExchangeRate.Text = 1
                Me.txtSubTotalAfterDP.Text = 0
                Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = 0
                RoundingAOP = 0
            Else
                Me.ObjAOP.Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(vPrimaryAOP, vTaxPercent, vUseTax, _
                vPrimaryCurrency, vTaxExchangeRate)
                Me.TxtPrimaryCurrency.Text = vPrimaryCurrency
                Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(vPrimaryCurrency)
                RoundingAOP = Me.ObjCurrency.GetDefaultPrecission(vPrimaryCurrency)
                Me.chkTax.Checked = vUseTax
                Me.txtTax.Text = vTaxPercent
                Me.txtTaxAmount.Text = 0
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(vTaxExchangeRate)
                Else
                    Me.txtTaxExchangeRate.Text = 0
                End If
                Me.txtSubTotalAfterDP.Text = 0
                'Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = 0
            End If

            For Each row_dgv In Me.dgvItem.Rows
                With row_dgv
                    If TBL1.Rows.Count <> 0 Then
                        Filter = "[Primary Product] =" + row_dgv.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString _
                        + " AND [Primary PO] = " + row_dgv.Cells(Me.ColPrimaryPO.Index).Value.ToString
                        TBL1.DefaultView.RowFilter = Filter
                        If TBL1.DefaultView.Count > 0 Then
                            rowTBL = TBL1.DefaultView(0).Row
                            .Cells(Me.ColPrimaryCurrency.Index).Value = rowTBL.Primary_Currency
                            Rounding = Me.ObjCurrency.GetDefaultPrecission(row_dgv.Cells(Me.ColPrimaryCurrency.Index).Value)
                            .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Unit_Price, Rounding))
                            .Cells(Me.ColDiscount.Index).Value = rowTBL.Discount
                            .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Nett_Price, Rounding))
                            .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(TBL_AOP, Me.ObjAOP.GetTipePerhitungan(vPrimaryAOP), _
                            row_dgv.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryGRN.Text, vPrimaryAOP))
                            .Cells(Me.ColCurrency.Index).Value = rowTBL.Currency_Code
                            .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColNettPrice.Index).Value * .Cells(Me.ColTotalQty.Index).Value)
                        Else
                            .Cells(Me.ColUnitPrice.Index).Value = 0
                            .Cells(Me.ColDiscount.Index).Value = 0
                            .Cells(Me.ColNettPrice.Index).Value = 0
                            .Cells(Me.ColExchangeRate.Index).Value = 1
                            .Cells(Me.ColCurrency.Index).Value = ""
                            .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                            .Cells(Me.ColAmount.Index).Value = 0
                        End If
                    Else
                        .Cells(Me.ColUnitPrice.Index).Value = 0
                        .Cells(Me.ColDiscount.Index).Value = 0
                        .Cells(Me.ColNettPrice.Index).Value = 0
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColCurrency.Index).Value = ""
                        .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                        .Cells(Me.ColAmount.Index).Value = 0
                    End If
                End With
            Next

            'hitung total
            Me.HitungTotal_Multi()

            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_AOP.Dispose()
            TBL_AOP.Dispose()
            DTA.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsiDGVItem_Multi()
        Dim rowDGV1 As DataGridViewRow
        Dim Clone As DataGridViewRow

        Dim TipeExcRate As ClsTransaksiAOP.EnumCalculatesBy = ClsTransaksiAOP.EnumCalculatesBy.AOP
        Dim DTA_AOP As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim TBL_AOP As New DseTransaksiAOP.tbltrans_aop_item_exchangerateDataTable
        Dim vPrimaryAOP As Double = -1
        Dim Rounding As Integer = 0
        Dim RoundingAOP As Integer = 0
        Dim vPrimaryCurrency, vTaxExchangeRate, vTaxPercent As Double
        Dim vUseTax As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.DataTable13_MultiTableAdapter
        Dim TBL1 As New DseTransaksiAOP.DataTable13_MultiDataTable
        Dim rowTBL As DseTransaksiAOP.DataTable13_MultiRow
        Dim Filter As String = ""
        Dim row_dgv As DataGridViewRow

        Try
            Me.dgvItem.Rows.Clear()

            For Each rowDGV1 In Me.DGVGRN.Rows
                Me.dgvItem.Rows.Add()
                Clone = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                With Clone
                    .Cells(Me.ColCodeNo.Index).Value = rowDGV1.Cells(Me.colCodeNoDGVGRN.Index).Value
                    .Cells(Me.ColDescription.Index).Value = rowDGV1.Cells(Me.colDescDGVGRN.Index).Value
                    .Cells(Me.ColPONO.Index).Value = rowDGV1.Cells(Me.colPONoDGVGRN.Index).Value
                    .Cells(Me.ColBatchNo.Index).Value = rowDGV1.Cells(Me.colBatchNoDGVGRN.Index).Value
                    .Cells(Me.ColExpiryDate.Index).Value = rowDGV1.Cells(Me.colExpiryDateDGVGRN.Index).Value
                    .Cells(Me.ColQtyMaster.Index).Value = rowDGV1.Cells(Me.colQtyPerMasterDGVGRN.Index).Value
                    .Cells(Me.ColQtyPerMaster.Index).Value = rowDGV1.Cells(Me.colQtyPerMasterDGVGRN.Index).Value
                    .Cells(Me.ColMasterUnits.Index).Value = rowDGV1.Cells(Me.colMasterUnitDGVGRN.Index).Value
                    .Cells(Me.ColQty.Index).Value = rowDGV1.Cells(Me.colQTYDGVGRN.Index).Value
                    .Cells(Me.ColTotalQty.Index).Value = rowDGV1.Cells(Me.colTotalQtyDGVGRN.Index).Value
                    .Cells(Me.ColUnit.Index).Value = rowDGV1.Cells(Me.colUnitDGVGRN.Index).Value
                    .Cells(Me.ColSecQty.Index).Value = rowDGV1.Cells(Me.colSecQtyDGVGRN.Index).Value
                    .Cells(Me.ColSecUnit.Index).Value = rowDGV1.Cells(Me.colSecUnitDGVGRN.Index).Value
                    .Cells(Me.colRemarks.Index).Value = ""
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGV1.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value
                    .Cells(Me.ColPrimaryPO.Index).Value = rowDGV1.Cells(Me.colPrimaryPODGVGRN.Index).Value
                    .Cells(Me.ColPrimaryMasterUnits.Index).Value = rowDGV1.Cells(Me.colPRimaryMasterUnitDGVGRN.Index).Value
                    .Cells(Me.ColPrimaryUnit.Index).Value = rowDGV1.Cells(Me.colPrimaryUnitDGVGRN.Index).Value
                    .Cells(Me.ColPrimarySecUnit.Index).Value = rowDGV1.Cells(Me.colPrimarySecUnitDGVGRN.Index).Value

                    ''@rosadinnn
                    ''@MUaT
                    .Cells(Me.ColQtyMulti.Index).Value = rowDGV1.Cells(Me.ColQtyMulti_GRN.Index).Value
                    .Cells(Me.ColUnitCodeMulti.Index).Value = rowDGV1.Cells(Me.ColUnitsCodeMulti_GRN.Index).Value
                    .Cells(Me.ColRatioDesc.Index).Value = rowDGV1.Cells(Me.ColRatioDesc_GRN.Index).Value
                    .Cells(Me.ColRatioReference.Index).Value = rowDGV1.Cells(Me.ColRatioReference_GRN.Index).Value
                    .Cells(Me.ColRatioDefault.Index).Value = rowDGV1.Cells(Me.ColRatioDefault_GRN.Index).Value
                    .Cells(Me.ColPrimaryUnitMulti.Index).Value = rowDGV1.Cells(Me.ColPrimaryUnitMulti_GRN.Index).Value
                End With
            Next

            'isikan data harga pada item
            DTA.Fill(TBL1, Me.txtPrimaryGRN.Text)
            vPrimaryAOP = Me.ObjAOP.GetPrimaryAOP_ByPrimaryGRN(Me.txtPrimaryGRN.Text)
            DTA_AOP.FillBy_PrimaryAOP(TBL_AOP, vPrimaryAOP)
            'isikan data currency, tipe pajak, tax percent, tax exchannge rate
            If vPrimaryAOP = -1 Then
                Me.TxtPrimaryCurrency.Text = -1
                Me.TxtCurrency.Text = ""
                Me.chkTax.Checked = False
                Me.txtTax.Text = 0
                Me.txtTaxAmount.Text = 0
                Me.txtTaxExchangeRate.Text = 1
                Me.txtSubTotalAfterDP.Text = 0
                Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = 0
                RoundingAOP = 0
            Else
                Me.ObjAOP.Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(vPrimaryAOP, vTaxPercent, vUseTax, _
                vPrimaryCurrency, vTaxExchangeRate)
                Me.TxtPrimaryCurrency.Text = vPrimaryCurrency
                Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(vPrimaryCurrency)
                RoundingAOP = Me.ObjCurrency.GetDefaultPrecission(vPrimaryCurrency)
                Me.chkTax.Checked = vUseTax
                Me.txtTax.Text = vTaxPercent
                Me.txtTaxAmount.Text = 0
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(vTaxExchangeRate)
                Else
                    Me.txtTaxExchangeRate.Text = 0
                End If
                Me.txtSubTotalAfterDP.Text = 0
                'Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = 0
            End If

            For Each row_dgv In Me.dgvItem.Rows
                With row_dgv
                    If TBL1.Rows.Count <> 0 Then
                        Filter = "[Primary Product] =" + row_dgv.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString _
                        + " AND [Primary PO] = " + row_dgv.Cells(Me.ColPrimaryPO.Index).Value.ToString
                        TBL1.DefaultView.RowFilter = Filter
                        If TBL1.DefaultView.Count > 0 Then
                            rowTBL = TBL1.DefaultView(0).Row
                            .Cells(Me.ColPrimaryCurrency.Index).Value = rowTBL.Primary_Currency
                            Rounding = Me.ObjCurrency.GetDefaultPrecission(row_dgv.Cells(Me.ColPrimaryCurrency.Index).Value)
                            .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Unit_Price, Rounding))
                            .Cells(Me.ColDiscount.Index).Value = rowTBL.Discount
                            .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Nett_Price, Rounding))
                            .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(TBL_AOP, Me.ObjAOP.GetTipePerhitungan(vPrimaryAOP), _
                            row_dgv.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryGRN.Text, vPrimaryAOP))
                            .Cells(Me.ColCurrency.Index).Value = rowTBL.Currency_Code

                            ''@rosadinnn
                            ''@MUaT
                            .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColNettPrice.Index).Value * .Cells(Me.ColQtyMulti.Index).Value)

                        Else
                            .Cells(Me.ColUnitPrice.Index).Value = 0
                            .Cells(Me.ColDiscount.Index).Value = 0
                            .Cells(Me.ColNettPrice.Index).Value = 0
                            .Cells(Me.ColExchangeRate.Index).Value = 1
                            .Cells(Me.ColCurrency.Index).Value = ""
                            .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                            .Cells(Me.ColAmount.Index).Value = 0
                        End If
                    Else
                        .Cells(Me.ColUnitPrice.Index).Value = 0
                        .Cells(Me.ColDiscount.Index).Value = 0
                        .Cells(Me.ColNettPrice.Index).Value = 0
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColCurrency.Index).Value = ""
                        .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                        .Cells(Me.ColAmount.Index).Value = 0
                    End If
                End With
            Next

            'hitung total
            Me.HitungTotal_Multi()

            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.FormatDGV()

            ''@rosadinnn
            ''@MUaT
            Me.SetMultiUnitProperty()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_AOP.Dispose()
            TBL_AOP.Dispose()
            DTA.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Function GetJml_Default(ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, ByVal vBatchNo As String, _
    ByVal vExpiry As Date, ByVal vSecQty As Double) As Double
        Dim row As DataGridViewRow
        Dim Qty As Double = 0
        Dim Lanjut As Boolean = False
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                    'cek ed
                    If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                        If vExpiry = #1/1/1950# Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    Else
                        If row.Cells(Me.ColExpiryDate.Index).Value = vExpiry Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If
                    'cek sec qty
                    If Lanjut = True Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = vSecQty Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        Else
                            If vSecQty = -1 Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        End If
                    End If
                End If
            Next

            Return Qty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_Multi_MasihDefault(ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, ByVal vBatchNo As String, _
    ByVal vExpiry As Date, ByVal vSecQty As Double) As Double
        Dim row As DataGridViewRow
        Dim Qty As Double = 0
        Dim Lanjut As Boolean = False
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                    'cek ed
                    If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                        If vExpiry = #1/1/1950# Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    Else
                        If row.Cells(Me.ColExpiryDate.Index).Value = vExpiry Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If
                    'cek sec qty
                    If Lanjut = True Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = vSecQty Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        Else
                            If vSecQty = -1 Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        End If
                    End If
                End If
            Next

            Return Qty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DGVGRN_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVGRN.CellBeginEdit
        'ga boleh diedit
        e.Cancel = True
    End Sub

    Private Sub KontrolKolom()
        Try
            If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                'revisi aop
                Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = False
                Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = False
                Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = False
                Me.dgvItem.Columns(Me.ColAmount.Index).Visible = False
                Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = False
                Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = False 
            Else
                'nota return
                Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = True
                Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = True
                Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = True
                Me.dgvItem.Columns(Me.ColAmount.Index).Visible = True
                Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = True
                Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = True 
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTipePurchaseReturn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipePurchaseReturn.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ChkInvoiceCancellation.Checked = False
                If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                    'revisi aop
                    Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColAmount.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = False
                    Me.TxtTaxReturnNotes.Text = ""
                    Me.TxtTaxReturnNotes.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtTaxReturnNotes.ReadOnly = True
                    Me.TxtTaxReturnNotes.TabStop = False

                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                    Me.cboPrimaryAccount.Enabled = False
                    Me.cboAccount.Enabled = False

                    'Me.cboCurrTax.Enabled = False
                    Me.cboPrimaryCurrTax.Enabled = False
                    Me.cboCurrTax.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1

                    Me.ColPriceByMultiUnit_str.Visible = False
                Else
                    'nota return
                    Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColAmount.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = True
                    'Me.TxtTaxReturnNotes.Text = ""
                    Me.TxtTaxReturnNotes.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.TxtTaxReturnNotes.ReadOnly = False
                    Me.TxtTaxReturnNotes.TabStop = True



                    'Me.chkTax.Enabled = True
                    Me.cboCurrTax.Enabled = True
                    Me.cboPrimaryCurrTax.Enabled = True
                    Me.cboPrimaryAccount.Enabled = True
                    Me.cboAccount.Enabled = True
                    'Me.txtPrimaryAccount.Visible = True
                    '  Me.TxtPrimaryCurrency.Text = 1
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
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

                    If Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
                        Me.ChkInvoiceCancellation.Checked = True
                    End If

                    Me.ColPriceByMultiUnit_str.Visible = True
                End If
            End If
           

            If Me.TxtRemark.Enabled = True Then
                Me.IsikanDataHarga()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataHarga()
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiAOPTableAdapters.DataTable13TableAdapter
        Dim TBL As New DseTransaksiAOP.DataTable13DataTable
        Dim rowTBL As DseTransaksiAOP.DataTable13Row
        Dim Filter As String = ""
        Dim TipeExcRate As ClsTransaksiAOP.EnumCalculatesBy = ClsTransaksiAOP.EnumCalculatesBy.AOP
        Dim DTA_AOP As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim TBL_AOP As New DseTransaksiAOP.tbltrans_aop_item_exchangerateDataTable
        Dim vPrimaryAOP As Double = -1
        Dim Rounding As Integer = 0
        Dim RoundingInvoice As Integer = 0
        Dim vPrimaryCurrency, vTaxExchangeRate, vTaxPercent, vPrimaryCurrTax, vPrimaryAccount As Double
        Dim vUseTax As Boolean

        Try
            'isikan data harga pada item
            DTA.Fill(TBL, Me.txtPrimaryGRN.Text)
            vPrimaryAOP = Me.ObjAOP.GetPrimaryAOP_ByPrimaryGRN(Me.txtPrimaryGRN.Text)
            DTA_AOP.FillBy_PrimaryAOP(TBL_AOP, vPrimaryAOP)
            'isikan data currency, tipe pajak, tax percent, tax exchannge rate
            If vPrimaryAOP = -1 Then
                Rounding = Me.ObjCurrency.GetDefaultPrecission(1)
                Me.TxtPrimaryCurrency.Text = -1
                Me.TxtCurrency.Text = ""
                Me.chkTax.Checked = False
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(0, 2)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(1, 2)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
                RoundingInvoice = Rounding
                Me.cboPPNDipungut.SelectedIndex = 0
                Me.TxtPaymentTerm.Text = 0
                Me.txtPrimaryCurrTax.Text = -1
                Me.txtPrimaryAccount.Text = -1
                Me.cboCurrTax.DataSource = Nothing
                Me.cboAccount.DataSource = Nothing
                Me.cboPrimaryCurrTax.DataSource = Nothing
                Me.cboPrimaryCurrTax.DataSource = Nothing
            Else
                Me.ObjAOP.Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(vPrimaryAOP, vTaxPercent, vUseTax, _
                vPrimaryCurrency, vTaxExchangeRate)
                Me.TxtPrimaryCurrency.Text = vPrimaryCurrency
                Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(vPrimaryCurrency)
                RoundingInvoice = Me.ObjCurrency.GetDefaultPrecission(vPrimaryCurrency)
                Me.chkTax.Checked = vUseTax
                If Me.ObjAOP.ApakahVATDipungut(vPrimaryAOP) = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If

                If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                    vPrimaryCurrTax = -1
                    vPrimaryAccount = -1
                Else
                    vPrimaryCurrTax = Me.ObjAOP.GetPrimaryCurrTax(vPrimaryAOP)
                    vPrimaryAccount = Me.ObjAOP.GetPrimaryAccountAP(vPrimaryAOP)
                End If

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
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(vTaxPercent, 2)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, Me.ObjCurrency.GetDefaultPrecission(vPrimaryCurrTax))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, Me.ObjCurrency.GetDefaultPrecission(vPrimaryCurrTax))
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(vTaxExchangeRate, 2)
                Else
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, RoundingInvoice)
                'Me.TxtDutystamp.Text = 0
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, RoundingInvoice)

                Me.TxtPaymentTerm.Text = Me.ObjSupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)
            End If

            For Each row In Me.dgvItem.Rows
                With row
                    If TBL.Rows.Count <> 0 Then
                        Filter = "[Primary Product] =" + row.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString _
                        + " AND [Primary PO] = " + row.Cells(Me.ColPrimaryPO.Index).Value.ToString
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count > 0 Then
                            rowTBL = TBL.DefaultView(0).Row
                            .Cells(Me.ColPrimaryCurrency.Index).Value = rowTBL.Primary_Currency
                            Rounding = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                            .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Unit_Price, Rounding))
                            .Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(rowTBL.Discount, 2)
                            .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowTBL.Nett_Price, Rounding))
                            .Cells(Me.ColExchangeRate.Index).Value = Me.ObjCD.DigitGrouping(Me.GetExchangeRate(TBL_AOP, Me.ObjAOP.GetTipePerhitungan(vPrimaryAOP), _
                            row.Cells(Me.ColPrimaryPO.Index).Value, Me.txtPrimaryGRN.Text, vPrimaryAOP))
                            .Cells(Me.ColCurrency.Index).Value = rowTBL.Currency_Code
                            .Cells(Me.ColPriceByMultiUnit.Index).Value = rowTBL.price_by_multi_unit
                            If rowTBL.price_by_multi_unit = 1 Then
                                .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColNettPrice.Index).Value) * CDec(.Cells(Me.ColQtyMulti.Index).Value), Rounding)
                                .Cells(Me.ColPriceByMultiUnit_str.Index).Value = "/" + .Cells(Me.ColUnitCodeMulti.Index).Value
                            Else
                                .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColNettPrice.Index).Value) * CDec(.Cells(Me.ColTotalQty.Index).Value), Rounding)
                                .Cells(Me.ColPriceByMultiUnit_str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                            End If
                        Else
                            .Cells(Me.ColUnitPrice.Index).Value = 0
                            .Cells(Me.ColDiscount.Index).Value = 0
                            .Cells(Me.ColNettPrice.Index).Value = 0
                            .Cells(Me.ColExchangeRate.Index).Value = 1
                            .Cells(Me.ColCurrency.Index).Value = ""
                            .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                            .Cells(Me.ColAmount.Index).Value = 0
                            .Cells(Me.ColPriceByMultiUnit_str.Index).Value = "'"
                            .Cells(Me.ColPriceByMultiUnit.Index).Value = 1
                        End If
                    Else
                        .Cells(Me.ColUnitPrice.Index).Value = 0
                        .Cells(Me.ColDiscount.Index).Value = 0
                        .Cells(Me.ColNettPrice.Index).Value = 0
                        .Cells(Me.ColExchangeRate.Index).Value = 1
                        .Cells(Me.ColCurrency.Index).Value = ""
                        .Cells(Me.ColPrimaryCurrency.Index).Value = -1
                        .Cells(Me.ColAmount.Index).Value = 0
                        .Cells(Me.ColPriceByMultiUnit.Index).Value = 1
                        .Cells(Me.ColPriceByMultiUnit_str.Index).Value = ""
                    End If
                End With
            Next

            'hitung total
            Me.HitungTotal_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_AOP.Dispose()
            TBL_AOP.Dispose()
        End Try
    End Sub

    'Private Sub HitungTotal_Default()
    '    Dim rowdgv As DataGridViewRow
    '    Dim SubTotal, TaxAmount, Total As Double
    '    Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
    '    Dim DefPrecIDR As Integer = 0
    '    Dim AmountInIDR As Decimal = 0
    '    Dim SubTotal1 As Decimal = 0

    '    Try
    '        DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

    '        If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then
    '            'hitung subtotal
    '            For Each rowdgv In Me.dgvItem.Rows
    '                SubTotal += Math.Round(CDbl(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '                CDbl(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
    '                Rounding, MidpointRounding.AwayFromZero)
    '            Next
    '            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
    '            'hitung pajak
    '            If Me.cboPPNDipungut.SelectedIndex = 0 Then
    '                TaxAmount = Me.ObjCD.DigitGrouping(SubTotal * CDbl(Me.txtTax.Text) * CDbl(Me.txtTaxExchangeRate.Text) / 100, _
    '                DefPrecIDR, ClsCekData.EnumTipeRounding.Down)
    '            Else
    '                TaxAmount = 0
    '            End If
    '            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecIDR)
    '            'dutystamp
    '            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Math.Round(CDbl(Me.TxtDutystamp.Text), DefPrecIDR, MidpointRounding.AwayFromZero), DefPrecIDR)
    '            If Me.TxtPrimaryCurrency.Text = 1 Then
    '                'jika rupiah, maka total adalah = sub total + tax + duty stamp
    '                Total = SubTotal + TaxAmount + CDbl(Me.TxtDutystamp.Text)
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Total, Rounding))
    '            Else
    '                Total = SubTotal
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Total, Rounding))
    '            End If
    '        Else
    '            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
    '            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, DefPrecIDR)
    '            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
    '            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(1, 2)
    '            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, DefPrecIDR)
    '            Me.txtTax.Text = Me.ObjCD.DigitGrouping(0, 2)
    '        End If

    '        'hitung purchase price yang akan digunakan untuk jurnal closing
    '        AmountInIDR = 0
    '        SubTotal1 = 0
    '        For Each rowdgv In Me.dgvItem.Rows
    '            SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
    '            Rounding, MidpointRounding.AwayFromZero)
    '        Next
    '        SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

    '        For Each rowdgv In Me.dgvItem.Rows
    '            If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
    '                rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = SubTotal - AmountInIDR
    '            Else
    '                If SubTotal1 <> 0 Then
    '                    rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
    '                    DefPrecIDR, MidpointRounding.AwayFromZero)
    '                    AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchaseInIDR.Index).Value)
    '                End If
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub   

    Private Function GetExchangeRate(ByVal vTBL As DseTransaksiAOP.tbltrans_aop_item_exchangerateDataTable, _
   ByVal vTipePerhitungan As ClsTransaksiAOP.EnumCalculatesBy, ByVal vPrimaryPO As Double, _
   ByVal vPrimaryGRN As Double, ByVal vPrimaryAOP As Double) As Double
        Dim Filter As String = ""
        Dim row As DseTransaksiAOP.tbltrans_aop_item_exchangerateRow

        Try
            If vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
                If vTBL.Rows.Count > 0 Then
                    row = vTBL.Rows(0)
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimaryGRN.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.PO Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimaryPO.ToString
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

    Private Sub TxtDutystamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Leave
        Try

            Me.ObjCD.Number_Leave(Me.TxtDutystamp, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text), Me.ObjCD)
            Me.HitungTotal_Multi()
            ' Me.IsikanDataHarga()
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
                    Me.TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
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
                    Me.TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
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
                    Me.TxtPaymentTerm.BackColor = Me.txtReadOnly.BackColor
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
            Me.ObjCD.Number_KeyUp(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentTerm.ReadOnly = True Then
                    Me.TxtPaymentTerm.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPaymentTerm.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetPONumber()
        Dim ListPrimaryPO As New List(Of Double)
        Dim rowdgv As DataGridViewRow

        Try
            If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                For Each rowdgv In Me.DGVGRN.Rows
                    'untuk setiap item GRN, cari no po dan qty sesuai dengan product code yang diberikan
                    If rowdgv.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value = Me.txtPrimaryCodeNo.Text Then
                        'kurangkan dengan item yang ada di dalam dgv item
                        If CDbl(rowdgv.Cells(Me.colTotalQtyDGVGRN.Index).Value) - Me.GetJmlDGV(rowdgv.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value, _
                        rowdgv.Cells(Me.colPrimaryPODGVGRN.Index).Value) > 0 Then
                            If ListPrimaryPO.Count > 0 Then
                                'kalo ada 2 po, user harus masukkin sendiri
                                Exit Sub
                            Else
                                ListPrimaryPO.Add(rowdgv.Cells(Me.colPrimaryPODGVGRN.Index).Value)
                            End If
                        End If
                    End If
                Next

                If ListPrimaryPO.Count > 0 Then
                    Me.txtPrimaryPO.Text = ListPrimaryPO.Item(0)
                    Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPO = Nothing
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.objPurchaseReturn.CariDetail(vPrimary, Me.dgvItem, Me) = True Then
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

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                'Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))

                ''@rosadinnn
                ''@MUaT
                If Me.txtPrimaryUnitsMulti.Text = "" Then
                    Me.txtPrimaryUnitsMulti.Text = "-1"
                End If
                Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.SetQtyDefault()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmCetakTransaksiAOP

        Try
            If Me.TxtPrimaryPurchaseReturn.Text <> "" Then
                If IsNumeric(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                    If CDbl(Me.TxtPrimaryPurchaseReturn.Text) >= 0 Then
                        'jika quick purchase return maka tidak boleh diedit.
                        If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                            MsgBox("Please print this data from quick purchase return form.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        If objPurchaseReturn.Cetak(ij.crv1, Me.TxtPrimaryPurchaseReturn.Text, False, -1) = True Then
                            ij.Text = "Purchase Return"
                            ij.ShowDialog()
                        End If
                        Exit Sub
                    Else
                        Me.btnFind.PerformClick()
                    End If
                Else
                    Me.btnFind.PerformClick()
                End If
            Else
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryPurchaseReturn.Text <> "" And Me.TxtPrimaryPurchaseReturn.Text <> "-1" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                    MsgBox("Please print this data from quick purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                objPurchaseReturn.Cetak(ij.crv1, Me.TxtPrimaryPurchaseReturn.Text, False, -1)
                ij.Text = "Purchase Return"
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxReturnNotes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxReturnNotes.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxReturnNotes.Enabled = True Then
                    If Me.TxtTaxReturnNotes.ReadOnly = True Then
                        Me.TxtTaxReturnNotes.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtTaxReturnNotes.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxReturnNotes_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxReturnNotes.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxReturnNotes.Enabled = True Then
                    If Me.TxtTaxReturnNotes.ReadOnly = True Then
                        Me.TxtTaxReturnNotes.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtTaxReturnNotes.BackColor = Me.txtPrimaryCategory.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSupplierNPWPNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierNPWPNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtSupplierNPWPNo.Enabled = True Then
                    If Me.TxtSupplierNPWPNo.ReadOnly = True Then
                        Me.TxtSupplierNPWPNo.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtSupplierNPWPNo.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSupplierNPWPNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierNPWPNo.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtSupplierNPWPNo.Enabled = True Then
                    If Me.TxtSupplierNPWPNo.ReadOnly = True Then
                        Me.TxtSupplierNPWPNo.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtSupplierNPWPNo.BackColor = Me.txtPrimaryCategory.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Function IsContainMatchData(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Boolean
        Try
            For Each RowDgv As DataGridViewRow In Me.DGVGRN.Rows
                With RowDgv
                    If CDbl(.Cells(Me.colPrimaryPODGVGRN.Index).Value) = vPrimaryPO And CDbl(.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value = vPrimaryProduct) Then
                        Return True
                    End If
                End With
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Enum EnumGetOption As Integer
        PRIMARY_UNIT_DEFAULT
        PRIMARY_UNIT_MULTI
        UNIT_CODE_DEFAULT
        UNIT_CODE_MULTI
        RATIO_REFERENCE
        RATIO_DEFAULT
        RATIO_DESC
    End Enum

    Private Function GetMultiUnitProperty(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, ByVal vEnumGetOption As EnumGetOption) As String
        Try
            For Each RowDgv As DataGridViewRow In Me.DGVGRN.Rows
                With RowDgv
                    If CDbl(.Cells(Me.colPrimaryPODGVGRN.Index).Value) = vPrimaryPO And CDbl(.Cells(Me.colPrimaryCodeNoDGVGRN.Index).Value = vPrimaryProduct) Then
                        Select Case vEnumGetOption
                            Case EnumGetOption.PRIMARY_UNIT_DEFAULT
                                Return .Cells(Me.colPrimaryUnitDGVGRN.Index).Value.ToString
                            Case EnumGetOption.PRIMARY_UNIT_MULTI
                                Return .Cells(Me.ColPrimaryUnitMulti_GRN.Index).Value.ToString
                            Case EnumGetOption.RATIO_DEFAULT
                                Return ObjCD.DigitGrouping(.Cells(Me.ColRatioDefault_GRN.Index).Value.ToString, ObjProduct.GetAngkaDecimal(vPrimaryProduct))
                            Case EnumGetOption.RATIO_DESC
                                Return .Cells(Me.ColRatioDesc_GRN.Index).Value
                            Case EnumGetOption.RATIO_REFERENCE
                                Return ObjCD.DigitGrouping(.Cells(Me.ColRatioReference_GRN.Index).Value, ObjProduct.GetAngkaDecimal(vPrimaryProduct, CDbl(.Cells(Me.ColPrimaryUnitMulti_GRN.Index).Value)))
                            Case EnumGetOption.UNIT_CODE_DEFAULT
                                Return .Cells(Me.colUnitDGVGRN.Index).Value
                            Case EnumGetOption.UNIT_CODE_MULTI
                                Return .Cells(Me.ColUnitsCodeMulti_GRN.Index).Value
                        End Select
                    End If
                End With
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SetMultiUnitProperty()
        Try
            If Me.txtPrimaryStockInPurchase.Text <> "-1" And Me.txtPrimaryStockInPurchase.Text <> "" And _
                    Me.txtPrimaryPO.Text <> "-1" And Me.txtPrimaryPO.Text <> "" And _
                    Me.txtPrimaryCodeNo.Text <> "-1" And Me.txtPrimaryCodeNo.Text <> "" Then
                If IsContainMatchData(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text) = True Then
                    Me.txtPrimaryUnitsDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.PRIMARY_UNIT_DEFAULT)
                    Me.txtPrimaryUnitsMulti.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.PRIMARY_UNIT_MULTI)
                    Me.TxtRatioReference.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.RATIO_REFERENCE)
                    Me.TxtRatioDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.RATIO_DEFAULT)
                    Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.lblUnitsMulti.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.UNIT_CODE_MULTI)
                    Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
                    Me.lblUnitsDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.UNIT_CODE_DEFAULT)
                    Me.LblQty.Text = "Qty (" & Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, EnumGetOption.RATIO_DESC) & ")"
                    SetQtyDefault()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryPO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimaryPO.TextChanged
        Try
            Me.SetMultiUnitProperty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimaryCodeNo.TextChanged
        Try
            Me.SetMultiUnitProperty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            If Me.TxtRatioDefault.Text <> "0" And Me.TxtRatioReference.Text <> "0" And Me.TxtRatioDefault.Text <> "" And Me.TxtRatioReference.Text <> "" Then
                Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    'Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged

    '    Try
    '        If Me.txtPrimaryCurrTax.Text = -1 Then
    '            If Me.txtPrimaryCurrency.Text <= 1 Then
    '                Me.cboCurrTax.Enabled = False
    '            Else
    '                Me.cboCurrTax.Enabled = True
    '            End If

    '            Me.IsiCBOCurrencyTax()
    '            Me.IsiCBOArAkunByTipePerkiraan1()
    '        End If


    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

 
    Private Sub chkTax_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                'If Me.chkTax.Checked = True Then
                '    Me.txtTax.ReadOnly = False
                '    Me.txtTax.TabStop = True
                'Else
                '    Me.txtTax.ReadOnly = True
                '    Me.txtTax.TabStop = False
                'End If

                If Me.cboTipePurchaseReturn.SelectedIndex = 1 Or Me.cboTipePurchaseReturn.SelectedIndex = 2 Then

                    Me.HitungTotal_Multi()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_PReturn_Create(ByVal vPrimaryTransaksi As Double, ByVal vPrimarySupp As Double, _
    ByVal vTblPrimaryGrn As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)
        Dim TblTempPO As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objPurchaseReturn.ResetObject(Me.dgvItem)
            Me.objPurchaseReturn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me)
            Me.objPurchaseReturn.PrimaryGRN = ObjCutOff_Core.GetPrimaryUpdate(Me.objPurchaseReturn.PrimaryGRN, vTblPrimaryGrn)
            Me.objPurchaseReturn.PrimarySupplier = vPrimarySupp

            'Update Primary PO
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPO.Index).Value) = False And rowD.Cells(Me.ColPrimaryPO.Index).Value <> "-1" Then
                    TblTempPO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value)
                    If TblTempPO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), vTblPrimaryPO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), PrimaryUpdate, TblTempPO)
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), TblTempPO)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_PReturn_Save() As Double
        Try
            Me.baru = True
            Me.objPurchaseReturn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me, False, False, True)
            Return Me.objPurchaseReturn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class