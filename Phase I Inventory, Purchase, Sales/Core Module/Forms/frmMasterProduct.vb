Option Explicit On
Option Strict Off

Imports system.io

Public Class frmMasterProduct
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public EditDrillDown As Boolean = False
    Public NewDrillDown As Boolean = False
    Public PrimaryDrillDown As Double = -1

    Dim baru As Boolean
    Dim objProduct As clsMasterProduct
    Dim ObjCekData As New ClsCekData
    Dim ObjCategory As New ClsMasterCategory
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData
    Dim ObjLicense As New ClsLicense
    Dim ObjAkun As New ClsMasterAkun1
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjWorkUnit As New ClsMasterWorkUnit

    'terkait kompetensi
    Dim ObjKompetensi As New ClsMasterKompetensi_Product
    Dim mDgv As DataGridView
    Dim mEraseRowIndex As Integer = 0
    Dim EraseRowIndex As Integer = 0
    Dim ObjUser As New Core_Module.ClsUser

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage5.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            'add weight & volume
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer8.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer8.Panel2.Controls, Me.txtPrimaryProduct.BackColor, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel2.Controls, True, True, True, True, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.chkUseSecondaryUnit.Enabled = True
            Me.chkLocation.Enabled = True
            Me.chkBarcode.Enabled = True
            Me.chkSizing.Enabled = True
            If Me.chkSizing.Checked = True Then
                Me.cboSizing.Enabled = True
            Else
                Me.cboSizing.Enabled = False
            End If
            Me.chkBatchNo.Enabled = True
            Me.chkPDOControlled.Enabled = True
            Me.chkCategoryBindCode.Enabled = True
            Me.chkExpiryDate.Enabled = True
            Me.chkCategoryBindDecs.Enabled = True
            If Me.chkBatchNo.Checked = False Then
                Me.chkExpiryDate.Checked = False
                Me.chkExpiryDate.Enabled = False
            End If
            Me.rboGoods.Enabled = True
            Me.rboService.Enabled = True
           

            Me.chkSizing.Enabled = True
            If Me.chkSizing.Checked = True Then
                Me.cboSizing.Enabled = True
            Else
                Me.cboSizing.Enabled = False
            End If

            'untuk inventory, jika sudah closing pertama atau sudah ada aop, maka tidak bisa diedit
            If Me.baru = False Then
                If Me.rboGoods.Checked = True Then
                    If Me.TxtPrimaryInventoryAcc.Text <> "-1" Then 'hanya kalo primary inventory account =-1, supaya memudahkan edit datanya ketika peralihan
                        If Me.objProduct.ApakahSudahPernahJurnal(Me.txtPrimaryProduct.Text) = True Then
                            Me.TxtInventoryAcc.ReadOnly = True
                            Me.TxtInventoryAcc.BackColor = Me.TxtReadOnly.BackColor
                            Me.TxtInventoryAcc.TabStop = False
                            Me.btnFindInventoryAcc.Enabled = False
                        Else
                            If Me.objProduct.ApakahSudahPernahClosing(Me.txtPrimaryProduct.Text) = True Then
                                Me.TxtInventoryAcc.ReadOnly = True
                                Me.TxtInventoryAcc.BackColor = Me.TxtReadOnly.BackColor
                                Me.TxtInventoryAcc.TabStop = False
                                Me.btnFindInventoryAcc.Enabled = False
                            Else
                                Me.TxtInventoryAcc.ReadOnly = False
                                Me.TxtInventoryAcc.BackColor = Me.txtPrimaryProduct.BackColor
                                Me.TxtInventoryAcc.TabStop = True
                                Me.btnFindInventoryAcc.Enabled = True
                            End If
                        End If
                    Else
                        Me.TxtInventoryAcc.ReadOnly = False
                        Me.TxtInventoryAcc.BackColor = Me.txtPrimaryProduct.BackColor
                        Me.TxtInventoryAcc.TabStop = True
                        Me.btnFindInventoryAcc.Enabled = True
                    End If
                End If
            Else
                Me.TxtInventoryAcc.ReadOnly = False
                Me.TxtInventoryAcc.BackColor = Me.txtPrimaryProduct.BackColor
                Me.TxtInventoryAcc.TabStop = True
                Me.btnFindInventoryAcc.Enabled = True
            End If
            'berkaitan dengan goods-service
            If Me.rboGoods.Checked = False Then
                Me.chkBatchNo.Checked = False
                Me.chkExpiryDate.Checked = False
                Me.chkLocation.Checked = False
                Me.chkSizing.Checked = False

                If Me.txtUnit.Enabled = True Then
                    Me.chkBatchNo.Enabled = False
                    Me.chkExpiryDate.Enabled = False
                    Me.chkLocation.Enabled = False
                    Me.chkSizing.Enabled = False
                End If
            Else
                If Me.txtUnit.Enabled = True Then
                    Me.chkBatchNo.Enabled = True
                    If Me.chkBatchNo.Checked = True Then
                        Me.chkExpiryDate.Enabled = True
                    Else
                        Me.chkExpiryDate.Enabled = False
                    End If
                    Me.chkLocation.Enabled = True
                    Me.chkSizing.Enabled = True
                End If
            End If

            'jika edit dan product sudah pernah digunakan maka check box expiry date, batch no, location, 
            'use secondary unit
            If Me.baru = False Then
                If Me.objProduct.ApakahSudahPernahDigunakan(Me.txtPrimaryProduct.Text) = True Then
                    Me.chkBatchNo.Enabled = False
                    Me.chkExpiryDate.Enabled = False
                    Me.chkLocation.Enabled = False
                    Me.chkPDOControlled.Enabled = False
                    Me.chkUseSecondaryUnit.Enabled = False
                    Me.rboGoods.Enabled = False
                    Me.rboService.Enabled = False
                End If
                'penambahan untuk perbaikan master product
                If Me.objProduct.ApakahSudahPernahDigunakan2(Me.txtPrimaryProduct.Text) = True Then
                    Me.txtUnit.Enabled = False
                    Me.btnFindUnit.Enabled = False
                End If

                'add weight & volume
                If Me.objProduct.ApakahBolehEditUnitWeight(Me.txtPrimaryProduct.Text, Me.txtPrimaryUnit.Text, Me.TxtPrimaryUnitWeight.Text) = False Then
                    Me.TxtUnitWeight.Enabled = False
                    Me.BtnFindUnitWeight.Enabled = False
                End If
                If Me.objProduct.ApakahBolehEditUnitVolume(Me.txtPrimaryProduct.Text, Me.txtPrimaryUnit.Text, Me.TxtPrimaryUnitWeight.Text) = False Then
                    Me.TxtUnitVolume.Enabled = False
                    Me.BtnFindUnitVolume.Enabled = False
                End If
            Else
                Me.rboGoods.Enabled = True
                Me.rboService.Enabled = True
                If Me.txtProductCode.Text = "" Then
                    Me.rboGoods.Checked = True
                End If
            End If

            'berkaitan dengan bill of material
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, False, False, False, False)
            Me.dgvItemBOM.AllowUserToAddRows = True
            Me.dgvItemBOM.AllowUserToDeleteRows = True
            Me.dgvItemBOM.ReadOnly = False
            Me.SetTabBOM()

            Me.dgvItem_MultiUnit.AllowUserToAddRows = True
            Me.dgvItem_MultiUnit.AllowUserToDeleteRows = True
            Me.dgvItem_MultiUnit.ReadOnly = False

            Me.LlbEditPriceList.Enabled = False

            Me.DGVKompetensi.AllowUserToAddRows = True
            Me.DGVKompetensi.ReadOnly = False
            Me.DGVKompetensi.AllowUserToDeleteRows = True

            Me.DGVAttanchment.ReadOnly = False
            Me.DGVAttanchment.AllowUserToAddRows = True
            Me.DGVAttanchment.AllowUserToDeleteRows = True

            Me.CboTypeFinishGood.Enabled = True
            Me.DgvWorkUnit.ReadOnly = False
            Me.DgvWorkUnit.AllowUserToAddRows = True
            Me.DgvWorkUnit.AllowUserToDeleteRows = True

            If Me.rboService.Checked = True Then
                Me.CboTypeFinishGood.Enabled = False
                Me.CboTypeFinishGood.SelectedIndex = 0
            Else
                Me.CboTypeFinishGood.Enabled = True
            End If

            Me.llbCopy.Enabled = False
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
            Me.btnUp1.Enabled = True
            Me.btnDown1.Enabled = True
            Me.cboDependency.Enabled = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "frmMasterProduct_InputWeightNVolume") = False Then
                Me.GboDetailVolume.Enabled = False
                Me.GboDetailWeight.Enabled = False
            Else
                Me.GboDetailVolume.Enabled = True
                Me.GboDetailWeight.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage5.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer8.Panel2.Controls, True, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnExit
            Me.chkUseSecondaryUnit.Enabled = False
            Me.chkLocation.Enabled = False
            Me.chkBatchNo.Enabled = False
            Me.chkPDOControlled.Enabled = False
            Me.chkBarcode.Enabled = False
            Me.chkCategoryBindCode.Enabled = False
            Me.chkExpiryDate.Enabled = False
            Me.chkCategoryBindDecs.Enabled = False
            Me.cboSizing.Enabled = False
            Me.chkSizing.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage5.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.rboGoods.Enabled = False
            Me.rboService.Enabled = False

            'berkaitan dengan bill of material
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, False, False, False, False)
            Me.dgvItemBOM.AllowUserToAddRows = False
            Me.dgvItemBOM.AllowUserToDeleteRows = False
            Me.dgvItemBOM.ReadOnly = True
            Me.SetTabBOM()

            If Me.EditDrillDown = True Or Me.NewDrillDown = True Then
                Me.Close()
            End If

            Me.dgvItem_MultiUnit.AllowUserToAddRows = False
            Me.dgvItem_MultiUnit.AllowUserToDeleteRows = False
            Me.dgvItem_MultiUnit.ReadOnly = True
            Me.LlbEditPriceList.Enabled = True

            Me.DGVKompetensi.AllowUserToAddRows = False
            Me.DGVKompetensi.AllowUserToDeleteRows = False
            Me.DGVKompetensi.AllowUserToResizeColumns = False
            Me.DGVKompetensi.AllowUserToResizeRows = False
            Me.DGVKompetensi.ReadOnly = True

            Me.DGVAttanchment.ReadOnly = True
            Me.DGVAttanchment.AllowUserToAddRows = False
            Me.DGVAttanchment.AllowUserToDeleteRows = False

            Me.DgvWorkUnit.ReadOnly = True
            Me.DgvWorkUnit.AllowUserToAddRows = False
            Me.DgvWorkUnit.AllowUserToDeleteRows = False
            Me.CboTypeFinishGood.Enabled = False
            Me.llbCopy.Enabled = True
            Me.btnUP.Enabled = False
            Me.btnDown.Enabled = False
            Me.btnUp1.Enabled = False
            Me.btnDown1.Enabled = False
            Me.cboDependency.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.TabPage5.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, False, False)

            Me.txtPrimaryProduct.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryMasterUnit.Text = "-1"
            Me.TxtPrimaryInventoryAcc.Text = "-1"
            Me.TxtPrimarySalesAccount.Text = "-1"
            Me.TxtPrimarySalesReturnAccount.Text = "-1"
            Me.TxtPrimaryCOGSAccount.Text = "-1"
            Me.TxtPrimaryAdjustmentAccount.Text = "-1"
            Me.TxtPrimaryStockTakingAccount.Text = "-1"
            Me.TxtPrimaryTaxProductRef.Text = "-1"

            Me.txtPrimaryUnit.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtQtyPerMasterUnit.Text = "0"
            Me.txtQtyPerMasterUnit.ReadOnly = True
            Me.txtQtyPerMasterUnit.TabStop = False
            Me.txtQtyPerMasterUnit.BackColor = Me.txtDescriptionPatokanWarna.BackColor
            Me.txtUnitPrice.Text = "0"
            Me.txtDiscount.Text = "0"
            Me.txtLeadTime.Text = "0"
            Me.txtMinimumPurchase.Text = "0"
            Me.txtMinimumStock.Text = "0"
            Me.TxtDefaultPrecision.Text = "0"

            Me.pbxProduct.Image = Nothing
            Me.chkBatchNo.Checked = False
            Me.chkLocation.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.chkPDOControlled.Checked = False
            Me.chkExpiryDate.Checked = False
            Me.chkCategoryBindDecs.Checked = True
            Me.chkCategoryBindCode.Checked = True
            Me.txtSecondaryUnit.Enabled = False
            Me.btnFindSecondaryUnit.Enabled = False
            Me.dgvPricing.Rows.Clear()
            Me.chkSizing.Checked = False

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.LblProduct.Text = "Product Code - Description"

            Me.Kosongkan_BOM()
            Me.SetTabBOM()
            Me.dgvItem_MultiUnit.Rows.Clear()
            Me.DGVKompetensi.Rows.Clear()
            Me.DGVAttanchment.Rows.Clear()
            Me.DgvWorkUnit.Rows.Clear()

            Me.CboTypeFinishGood.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TipeAkun As ClsMasterAkun1.EnumTipeAkun2 = ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim TipeAkun1 As ClsMasterAkun1.EnumTipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset

        Try
            If CDec(Me.TxtQtyRef.Text) = 0 Then
                Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping("1", Me.TxtDefaultPrecision.Text)
            End If
            If CDec(Me.TxtGoodPercentage.Text) = 0 Then
                Me.TxtGoodPercentage.Text = "100.00"
                Me.txtTolerance.Text = "0.00"
            End If

            'trim
            Me.txtDescription.Text = Me.txtDescription.Text.Trim
            Me.txtProductCode.Text = Me.txtProductCode.Text.Trim

            'isikan data akun jika semua data akun kosong
            If Me.rboGoods.Checked = True Then
                If Me.TxtInventoryAcc.Text = "" And Me.TxtCOGSAccount.Text = "" And _
                Me.TxtAdjustmentAccount.Text = "" And Me.TxtStockTakingAccount.Text = "" And _
                Me.TxtSalesAccount.Text = "" And Me.TxtSalesReturnAccount.Text = "" Then
                    Me.IsikanAkunDefault()
                    MsgBox("Please check related account first.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage3
                    Return False
                End If
            Else
                If Me.TxtInventoryAcc.Text = "" And _
                Me.TxtSalesAccount.Text = "" And Me.TxtSalesReturnAccount.Text = "" Then
                    Me.IsikanAkunDefault()
                    MsgBox("Please check related account first.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage3
                    Return False
                End If
            End If

            'cek data kosong
            KoleksiNamaData.Add("product code")
            KoleksiTextBox.Add(Me.txtProductCode)
            KoleksiNamaData.Add("category code")
            KoleksiTextBox.Add(Me.txtCategoryCode)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.txtDescription)
            KoleksiNamaData.Add("units")
            KoleksiTextBox.Add(Me.txtUnit)
            If Me.chkUseSecondaryUnit.Checked = True Then
                KoleksiNamaData.Add("secondary units")
                KoleksiTextBox.Add(Me.txtSecondaryUnit)
            End If
            If Me.rboGoods.Checked = True Then
                KoleksiNamaData.Add("inventory account")
                KoleksiTextBox.Add(Me.TxtInventoryAcc)
                KoleksiNamaData.Add("cost of sales account")
                KoleksiTextBox.Add(Me.TxtCOGSAccount)
                KoleksiNamaData.Add("adjustment account")
                KoleksiTextBox.Add(Me.TxtAdjustmentAccount)
                KoleksiNamaData.Add("stock taking account")
                KoleksiTextBox.Add(Me.TxtStockTakingAccount)
            Else
                KoleksiNamaData.Add("expense account")
                KoleksiTextBox.Add(Me.TxtInventoryAcc)
            End If
            KoleksiNamaData.Add("sales account")
            KoleksiTextBox.Add(Me.TxtSalesAccount)
            KoleksiNamaData.Add("sales return account")
            KoleksiTextBox.Add(Me.TxtSalesReturnAccount)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data salah bila:
            '1.tidak ada data pricing
            'If Me.dgvPricing.RowCount = 0 Then
            '    MsgBox("Please input pricing data.", MsgBoxStyle.Information, "Information")
            '    Me.TabControl1.SelectedIndex = 1
            '    Me.txtCurrencyCode.Focus()
            '    Return False
            'End If

            '2.cek data yang salah, double dalam kode
            Dim ProductCodeTest As String = ""
            If Me.chkCategoryBindCode.Checked = True Then
                ProductCodeTest = Me.txtCategoryCode.Text + Me.txtProductCode.Text
            Else
                ProductCodeTest = Me.txtProductCode.Text
            End If
            If Me.baru = True Then
                If Me.objProduct.ApakahSudahAda(ProductCodeTest) = True Then
                    MsgBox("Duplication in product code, please change your code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtProductCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objProduct.GetProductCode(Me.txtPrimaryProduct.Text)
                If ProductGroupLama <> ProductCodeTest Then
                    If Me.objProduct.ApakahSudahAda(ProductCodeTest) = True Then
                        MsgBox("Duplication in product code, please change your code", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtProductCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objProduct.GetTimeStamp(Me.txtPrimaryProduct.Text)
                If Me.objProduct.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objProduct.CariDetail(Me.DGVKompetensi, Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
                        Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. qty per master unit negatif
            'kalo master unit tidak kosong maka qty per master unit harus positif
            'update: 2016-06-09-iwan: qty per master unit harusnya nol saja dan primary-nya -1
            If Me.txtPrimaryMasterUnit.Text <> "-1" Then
                Me.txtPrimaryMasterUnit.Text = -1
                If Me.ObjCekData.CekNonNegatif(Me.txtQtyPerMasterUnit.Text) = False Then
                    Me.txtQtyPerMasterUnit.Text = 0

                    'MsgBox("Quantity per master must be a non negatif number.", MsgBoxStyle.Information, _
                    '"Information")
                    'Me.txtQtyPerMasterUnit.Focus()
                    'Return False
                End If

                'If Me.ObjCekData.CekBilanganPositif(Me.txtQtyPerMasterUnit.Text) = False Then
                '    MsgBox("Please enter a positif number for quantity per master unit data.", _
                '    MsgBoxStyle.Information, "Information")
                '    Me.txtQtyPerMasterUnit.Focus()
                '    Return False
                'End If
            Else
                If Me.ObjCekData.CekNonNegatif(Me.txtQtyPerMasterUnit.Text) = False Then
                    Me.txtQtyPerMasterUnit.Text = 0

                    MsgBox("Quantity per master must be a non negatif number.", MsgBoxStyle.Information, _
                    "Information")
                    Me.txtQtyPerMasterUnit.Focus()
                    Return False
                End If
            End If

            '4. Data salah bila expiry date control tetapi tidak batch no control
            If Me.chkExpiryDate.Checked = True And Me.chkBatchNo.Checked = False Then
                MsgBox("To use expiry date control feature, please check the batch no checkbox also.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'default precision harus non negatif
            If Me.ObjCekData.CekNonNegatif(Me.TxtDefaultPrecision.Text) = False Then
                MsgBox("Default precision be a non negatif number.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtDefaultPrecision.Focus()
                Return False
            End If

            'default precision maximal bernilai 6
            If CInt(Me.TxtDefaultPrecision.Text) > 6 Then
                MsgBox("Default precision maximum value is 6.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtDefaultPrecision.Text = "6"
                Me.TxtDefaultPrecision.Focus()
                Return False
            End If

            'cek tipe akun
            TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryInventoryAcc.Text)
            If Me.rboGoods.Checked = True Then
                'tipe akun inventory haruslah inventory --> goods
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Inventory Then
                    MsgBox("Please specify an inventory type account for inventory account", MsgBoxStyle.Information, "Information")
                    Me.TxtInventoryAcc.Focus()
                    Return False
                End If
            Else
                'tipe akun inventory haruslah expense/other expense --> service
                TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryInventoryAcc.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    'boleh juga akun asset
                    TipeAkun1 = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryInventoryAcc.Text)
                    If TipeAkun1 <> ClsMasterAkun1.EnumTipeAkun.Asset Then
                        MsgBox("Please specify an expense/asset type account for expense account", MsgBoxStyle.Information, "Information")
                        Me.TxtInventoryAcc.Focus()
                        Return False
                    End If
                End If
            End If
            'tipe akun sales dan sales return haruslah sales return
            TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimarySalesAccount.Text)
            If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                MsgBox("Please specify an income/expense type account for sales account", MsgBoxStyle.Information, "Information")
                Me.TxtSalesAccount.Focus()
                Return False
            End If
            TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimarySalesReturnAccount.Text)
            If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                MsgBox("Please specify an income/expense type account for sales return account", MsgBoxStyle.Information, "Information")
                Me.TxtSalesReturnAccount.Focus()
                Return False
            End If
            If Me.rboGoods.Checked = True Then
                'tipe akun cogs haruslah cogs
                TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryCOGSAccount.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an cost of sales type account for cost of sales account", MsgBoxStyle.Information, "Information")
                    Me.TxtCOGSAccount.Focus()
                    Return False
                End If
                'tipe akun adjustment dan stock taking account harus lah expense, other expense
                TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryAdjustmentAccount.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an expense type account for adjustment account", MsgBoxStyle.Information, "Information")
                    Me.TxtAdjustmentAccount.Focus()
                    Return False
                End If
                TipeAkun = Me.ObjAkun.GetTipeAccount2(Me.TxtPrimaryStockTakingAccount.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                 TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an expense type account for stock taking account", MsgBoxStyle.Information, "Information")
                    Me.TxtStockTakingAccount.Focus()
                    Return False
                End If
            End If

            'qty ref harus positive
            If Me.ObjCekData.CekBilanganPositif(Me.TxtQtyRef.Text) = False Then
                MsgBox("Please enter a positif number for reference qty.", _
                MsgBoxStyle.Information, "Information")

                Me.TabControl1.SelectedIndex = 3
                Me.TxtQtyRef.Focus()
                Return False
            End If

            'good precentage harus positive
            If Me.ObjCekData.CekBilanganPositif(Me.TxtGoodPercentage.Text) = False Then
                MsgBox("Please enter a positif number for good percentage.", _
                MsgBoxStyle.Information, "Information")

                Me.TabControl1.SelectedIndex = 3
                Me.TxtQtyRef.Focus()
                Return False
            End If

            'kalo ada processnya, maka harus isikan minimal 1 raw material
            If Me.txtPrimaryProses.Text <> "" And Me.txtPrimaryProses.Text <> "-1" Then
                If Me.dgvItemBOM.Rows.Count <= 1 Then
                    MsgBox("Please specify minimal 1 material", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'semua item bill of material harus goods
            For Each rowdgv As DataGridViewRow In Me.dgvItemBOM.Rows
                If rowdgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value <> -1 Then
                    If Me.objProduct.ApakahGoods(Me.dgvItemBOM.Item(Me.ColBOM_PrimaryProduct.Index, rowdgv.Index).Value) = False Then
                        MsgBox("Each code no in bill of material must be goods", MsgBoxStyle.Information, "Information")

                        Me.dgvItemBOM.Focus()
                        Me.dgvItemBOM.ClearSelection()
                        rowdgv.Selected = True
                        Me.TabControl1.SelectedIndex = 3

                        Return False
                    End If
                End If
            Next

            'cek nilai pada dgv kompetensi
            Dim rowDgv2 As DataGridViewRow
            For Each rowDgv2 In Me.DGVKompetensi.Rows
                If rowDgv2.Cells(Me.ColCompetence_Code.Index).Value <> "" Then
                    If String.IsNullOrEmpty(rowDgv2.Cells(Me.ColCompetence_Value.Index).Value) = True Then
                        Me.DGVKompetensi.ClearSelection()
                        rowDgv2.Selected = True
                        Me.TabControl1.SelectedIndex = 4
                        MsgBox("Value data for custom value: '" + rowDgv2.Cells(Me.ColCempetence_Description.Index).Value.ToString + "' cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            For Each rowDgv2 In Me.DGVAttanchment.Rows
                If rowDgv2.Cells(Me.ColAtt_Description.Index).Value <> "" Then
                    If String.IsNullOrEmpty(rowDgv2.Cells(Me.ColAtt_LocationFile.Index).Value) = True Then
                        Me.DGVAttanchment.ClearSelection()
                        rowDgv2.Selected = True
                        Me.TabControl1.SelectedIndex = 5
                        MsgBox("File path data for '" + rowDgv2.Cells(Me.ColAtt_Description.Index).Value.ToString + "' cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            'jika tipenya raw material, tidak boleh memiliki item bill of material
            If Me.CboTypeFinishGood.SelectedIndex = CInt(clsMasterProduct.EnumTypeGoods.Raw_Material) Then
                If Me.dgvItemBOM.Rows.Count > 1 Then
                    MsgBox("Goods with type: 'raw material' can't have any bill of material item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika tipenya service, juga tidak boleh memiliki item bill of material
            If Me.rboService.Checked = CInt(clsMasterProduct.EnumTypeGoods.Raw_Material) Then
                If Me.dgvItemBOM.Rows.Count > 1 Then
                    MsgBox("'Service' data can't have any bill of material item.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            ''@aul 18/09/2015
            Dim ListBoM As New List(Of Double)
            Dim ListChild As New List(Of Double)
            If Me.dgvItemBOM.RowCount > 1 Then
                For Each rowDgv As DataGridViewRow In Me.dgvItemBOM.Rows
                    If IsNothing(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value) = False Then
                        If rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                            If ListBoM.Contains(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value.ToString) = False Then
                                ListBoM.Add(CDbl(rowDgv.Cells(Me.ColBOM_PrimaryProduct.Index).Value.ToString))
                            End If
                        End If
                    End If
                Next
            End If
            If ListBoM.Count > 0 Then
                GetChild(Me.txtPrimaryProduct.Text, ListChild)
                If ListChild.Count > 0 Then
                    Dim e As String = ""
                    For Each d As Double In ListBoM
                        If ListChild.Contains(d) Then
                            e = objProduct.GetProductCode(d)
                            MsgBox("Invalid bill of material data, the current data will cause infinitely long bill of material." & vbCrLf & "Code No : " & e, MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            End If

            'add weight & volume
            If CDec(Me.TxtNettWeight.Text) > 0 Or CDec(Me.TxtGrossWeight.Text) > 0 Or CInt(Me.TxtDefaultPrecisionWeight.Text) > 0 Then
                If Me.TxtUnitWeight.Text = "" Then
                    MsgBox("Unit weight cannot be an empty string.", MsgBoxStyle.Information, "Error")
                    Me.TabControl1.SelectedIndex = 0
                    Me.TxtUnitWeight.Focus()
                    Return False
                End If
                'nett tidak boleh lebih besar dari gross
                If CDec(Me.TxtNettWeight.Text) > CDec(Me.TxtGrossWeight.Text) Then
                    MsgBox("Nett weight cannot be greater than gross weight.", MsgBoxStyle.Information, "Error")
                    Me.TabControl1.SelectedIndex = 0
                    Me.TxtNettWeight.Focus()
                    Return False
                End If
            End If
            If CDec(Me.TxtNettVolume.Text) > 0 Or CDec(Me.TxtGrossVolume.Text) > 0 Or CInt(Me.TxtDefaultPrecisionVolume.Text) > 0 Then
                If Me.TxtUnitVolume.Text = "" Then
                    MsgBox("Unit volume cannot be an empty string.", MsgBoxStyle.Information, "Error")
                    Me.TabControl1.SelectedIndex = 0
                    Me.TxtUnitVolume.Focus()
                    Return False
                End If
                'nett tidak boleh lebih besar dari gross
                If CDec(Me.TxtNettVolume.Text) > CDec(Me.TxtGrossVolume.Text) Then
                    MsgBox("Nett volume cannot be greater than gross volume.", MsgBoxStyle.Information, "Error")
                    Me.TabControl1.SelectedIndex = 0
                    Me.TxtNettVolume.Focus()
                    Return False
                End If
            End If
            For Each rd As DataGridViewRow In Me.dgvItem_MultiUnit.Rows
                If String.IsNullOrEmpty(rd.Cells(Me.ColUnit_Unit.Index).Value) = False Then
                    If CDec(rd.Cells(Me.ColUnit_NettWeight.Index).Value) > 0 Or CDec(rd.Cells(Me.ColUnit_GrossWeight.Index).Value) > 0 Or CInt(rd.Cells(Me.ColUnit_PrecisionWeight.Index).Value) > 0 Then
                        If String.IsNullOrEmpty(rd.Cells(Me.ColUnit_UnitWeight.Index).Value) = True Then
                            MsgBox("Unit weight cannot be an empty string.", MsgBoxStyle.Information, "Error")
                            Me.TabControl1.SelectedIndex = 3
                            rd.Cells(Me.ColUnit_UnitWeight.Index).Selected = True
                            Return False
                        End If
                        'nett tidak boleh lebih besar dari gross
                        If CDec(rd.Cells(Me.ColUnit_NettWeight.Index).Value) > CDec(rd.Cells(Me.ColUnit_GrossWeight.Index).Value) Then
                            MsgBox("Nett weight cannot be greater than gross weight.", MsgBoxStyle.Information, "Error")
                            Me.TabControl1.SelectedIndex = 3
                            rd.Cells(Me.ColUnit_NettWeight.Index).Selected = True
                            Return False
                        End If
                    End If
                    If CDec(rd.Cells(Me.ColUnit_NettVolume.Index).Value) > 0 Or CDec(rd.Cells(Me.ColUnit_GrossVolume.Index).Value) > 0 Or CInt(rd.Cells(Me.ColUnit_PrecisionVolume.Index).Value) > 0 Then
                        If String.IsNullOrEmpty(rd.Cells(Me.ColUnit_UnitVolume.Index).Value) = True Then
                            MsgBox("Unit volume cannot be an empty string.", MsgBoxStyle.Information, "Error")
                            Me.TabControl1.SelectedIndex = 3
                            rd.Cells(Me.ColUnit_UnitVolume.Index).Selected = True
                            Return False
                        End If
                        'nett tidak boleh lebih besar dari gross
                        If CDec(rd.Cells(Me.ColUnit_NettVolume.Index).Value) > CDec(rd.Cells(Me.ColUnit_GrossVolume.Index).Value) Then
                            MsgBox("Nett volume cannot be greater than gross volume.", MsgBoxStyle.Information, "Error")
                            Me.TabControl1.SelectedIndex = 3
                            rd.Cells(Me.ColUnit_NettVolume.Index).Selected = True
                            Return False
                        End If
                    End If
                End If
            Next
            '----------------------------------------------------------------------------------------------------

            'FC
            If Me.txtPrimaryCategory.Text = "-1" Then
                MsgBox("category code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.txtPrimaryUnit.Text = "-1" Then
                MsgBox("unit code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Return False
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

    Private Sub GetChild(ByVal vPrimaryProduct As Double, ByVal vListChild As List(Of Double))
        Dim DtaBoM As New dseMasterProduct_BOMTableAdapters.tblmaster_product_item_bomTableAdapter
        Dim TblBoM As New dseMasterProduct_BOM.tblmaster_product_item_bomDataTable
        Dim RowBoM As dseMasterProduct_BOM.tblmaster_product_item_bomRow

        Try
            DtaBoM.Connection.ConnectionString = Core_Module.GetConstring
            DtaBoM.FillBy_PrimaryParent(TblBoM, vPrimaryProduct)

            If vListChild.Contains(vPrimaryProduct) = False Then
                vListChild.Add(vPrimaryProduct)
            End If
            For Each RowBoM In TblBoM
                If vListChild.Contains(RowBoM.primary_product) = False Then
                    vListChild.Add(RowBoM.primary_product)
                End If
                Me.GetChild(RowBoM.primary_product, vListChild)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBoM.Dispose()
            TblBoM.Dispose()
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.objProduct
                .AngkaDecimal = Me.TxtDefaultPrecision.Text
                .BatchNoControl = Me.chkBatchNo.Checked
                .LocationControl = Me.chkLocation.Checked
                .PdONumberControlled = Me.chkPDOControlled.Checked
                .ExpiryDateControlled = Me.chkExpiryDate.Checked
                .CategoryCodeBounded = Me.chkCategoryBindCode.Checked
                .CategoryDescBounded = Me.chkCategoryBindDecs.Checked

                .PrimaryCategory = Me.txtPrimaryCategory.Text
                Me.txtPrimaryMasterUnit.Text = "-1"
                .PrimaryMasterUnit = Me.txtPrimaryMasterUnit.Text
                .PrimaryUnit = Me.txtPrimaryUnit.Text
                .Barcode = Me.TxtBarcode.Text
                If Me.chkCategoryBindCode.Checked = True Then
                    .ProductCode = Me.txtCategoryCode.Text + Trim(Me.txtProductCode.Text)
                Else
                    .ProductCode = Trim(Me.txtProductCode.Text)
                End If
                .QtyPerMasterUnit = Me.txtQtyPerMasterUnit.Text
                Me.txtPrimarySecondaryUnit.Text = "-1"
                .PrimarySecondaryUnit = Me.txtPrimarySecondaryUnit.Text
                .Specification = Me.txtSpecification.Text
                .UseSecondaryUnits = Me.chkUseSecondaryUnit.Checked
                .MinimumPurchaseQty = Me.txtMinimumPurchase.Text
                .LeadTime = Me.txtLeadTime.Text
                .MinimumStock = Me.txtMinimumStock.Text
                If Me.chkCategoryBindDecs.Checked = False Then
                    .Description = Me.txtDescription.Text
                End If

                Me.objProduct.UseSizing = Me.chkSizing.Checked
                If Me.cboSizing.SelectedIndex = 0 Then
                    Me.objProduct.Size = 36
                ElseIf Me.cboSizing.SelectedIndex = 1 Then
                    Me.objProduct.Size = 37
                ElseIf Me.cboSizing.SelectedIndex = 2 Then
                    Me.objProduct.Size = 38
                ElseIf Me.cboSizing.SelectedIndex = 3 Then
                    Me.objProduct.Size = 39
                ElseIf Me.cboSizing.SelectedIndex = 4 Then
                    Me.objProduct.Size = 40
                ElseIf Me.cboSizing.SelectedIndex = 5 Then
                    Me.objProduct.Size = 41
                ElseIf Me.cboSizing.SelectedIndex = 6 Then
                    Me.objProduct.Size = 42
                ElseIf Me.cboSizing.SelectedIndex = 7 Then
                    Me.objProduct.Size = 43
                ElseIf Me.cboSizing.SelectedIndex = 8 Then
                    Me.objProduct.Size = 44
                ElseIf Me.cboSizing.SelectedIndex = 9 Then
                    Me.objProduct.Size = 45
                Else
                    Me.objProduct.Size = 46
                End If

                .PrimaryAkunInventory = Me.TxtPrimaryInventoryAcc.Text
                .PrimaryAkunSales = Me.TxtPrimarySalesAccount.Text
                .PrimaryAkunSalesReturn = Me.TxtPrimarySalesReturnAccount.Text
                .PrimaryAkunCOGS = Me.TxtPrimaryCOGSAccount.Text
                .PrimaryAkunAdjustment = Me.TxtPrimaryAdjustmentAccount.Text
                .PrimaryAkunStockTaking = Me.TxtPrimaryStockTakingAccount.Text
                If Me.rboGoods.Checked = True Then
                    .ProductType = clsMasterProduct.EnumTipeProduct.Goods
                Else
                    .ProductType = clsMasterProduct.EnumTipeProduct.Services
                End If

                'berkaitan dengan bill of material
                Me.objProduct.PrimaryProcess = Me.txtPrimaryProses.Text
                Me.objProduct.QtyRef = Me.TxtQtyRef.Text
                Me.objProduct.GoodPercentage = Me.TxtGoodPercentage.Text

                'berkaitan dengan verifikasi akun
                Me.objProduct.IsVerified = Me.TxtIsVerified.Text
                Me.objProduct.VerifiedBy = Me.TxtVerifiedBy.Text
                Me.objProduct.VerifiedTime = #1/1/1950#
                Me.objProduct.TypeGoods = Me.CboTypeFinishGood.SelectedIndex
                .TipeDependency = Me.cboDependency.SelectedIndex
                .DelayDependency = Me.ObjCD.Convert_ToMinutes(Me.txtDependency.Text)
                .Tolerance = Me.txtTolerance.Text

                'add weight & volume
                .NettWeight = Me.TxtNettWeight.Text
                .GrossWeight = Me.TxtGrossWeight.Text
                .PrimaryUnitWeight = Me.TxtPrimaryUnitWeight.Text
                .PrecisionWeight = Me.TxtDefaultPrecisionWeight.Text
                .NettVolume = Me.TxtNettVolume.Text
                .GrossVolume = Me.TxtGrossVolume.Text
                .PrimaryUnitVolume = Me.TxtPrimaryUnitVolume.Text
                .PrecisionVolume = Me.TxtDefaultPrecisionVolume.Text

                .PrimaryTaxProductRef = Me.TxtPrimaryTaxProductRef.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim ObjCat As New ClsMasterCategory
        Dim ObjUnit As New ClsMasterUnit
        Dim ObjTaxProduct As New ClsMasterTaxProductReference

        Try
            With Me.objProduct
                Me.chkBatchNo.Checked = .BatchNoControl
                Me.chkExpiryDate.Enabled = False
                Me.chkLocation.Checked = .LocationControl
                Me.chkPDOControlled.Checked = .PdONumberControlled
                Me.chkExpiryDate.Checked = .ExpiryDateControlled
                Me.chkCategoryBindCode.Checked = .CategoryCodeBounded
                Me.chkCategoryBindDecs.Checked = .CategoryDescBounded
                Me.TxtDefaultPrecision.Text = .AngkaDecimal

                If .Barcode <> "" Then
                    Me.chkBarcode.Checked = True
                    Me.TxtBarcode.Text = .Barcode
                Else
                    Me.chkBarcode.Checked = False
                    Me.TxtBarcode.Text = ""
                End If
                Me.txtPrimaryCategory.Text = .PrimaryCategory
                If Me.txtPrimaryCategory.Text <> "-1" Then
                    Me.txtCategoryCode.Text = ObjCat.GetCategoryCode(Me.txtPrimaryCategory.Text)
                    Me.txtCategoryCode2.Text = Me.txtCategoryCode.Text
                End If
                Me.txtPrimaryMasterUnit.Text = .PrimaryMasterUnit
                If Me.txtPrimaryMasterUnit.Text <> "-1" Then
                    Me.txtMasterUnit.Text = ObjUnit.GetUnitCode(Me.txtPrimaryMasterUnit.Text)
                End If
                Me.txtPrimaryUnit.Text = .PrimaryUnit
                If Me.txtPrimaryUnit.Text <> "-1" Then
                    Me.txtUnit.Text = ObjUnit.GetUnitCode(Me.txtPrimaryUnit.Text)
                End If
                If Me.chkCategoryBindCode.Checked = True Then
                    Me.txtProductCode.Text = .ProductCode.Substring(Me.txtCategoryCode.Text.Length)
                Else
                    Me.txtProductCode.Text = .ProductCode
                End If
                Me.txtPrimaryProduct.Text = .Primary
                Me.txtQtyPerMasterUnit.Text = .QtyPerMasterUnit
                Me.txtPrimarySecondaryUnit.Text = .PrimarySecondaryUnit
                If Me.txtPrimarySecondaryUnit.Text <> "-1" Then
                    Me.txtSecondaryUnit.Text = ObjUnit.GetUnitCode(Me.txtPrimarySecondaryUnit.Text)
                End If
                Me.txtSpecification.Text = .Specification
                Me.chkUseSecondaryUnit.Checked = .UseSecondaryUnits
                If Me.txtCategoryCode.Enabled = True Then
                    If Me.chkUseSecondaryUnit.Checked = True Then
                        Me.txtSecondaryUnit.Enabled = True
                        Me.btnFindSecondaryUnit.Enabled = True
                    Else
                        Me.txtSecondaryUnit.Clear()
                        Me.txtSecondaryUnit.Enabled = False
                        Me.txtPrimarySecondaryUnit.Text = "-1"
                        Me.btnFindSecondaryUnit.Enabled = False
                    End If
                End If
                Me.txtMinimumPurchase.Text = Me.ObjCekData.DigitGrouping(CStr(.MinimumPurchaseQty))
                Me.txtLeadTime.Text = Me.ObjCekData.DigitGrouping(CStr(.LeadTime))
                Me.txtMinimumStock.Text = Me.ObjCekData.DigitGrouping(CStr(.MinimumStock))
                If Me.chkCategoryBindDecs.Checked = False Then
                    Me.txtDescription.Text = .Description
                End If

                'format lbl
                Dim Code, Desc As String
                If Me.chkCategoryBindCode.Checked = True Then
                    Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                Else
                    Code = Me.txtProductCode.Text
                End If
                If Me.chkCategoryBindDecs.Checked = True Then
                    Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                    Me.txtDescription.Text = Desc
                Else
                    Desc = Me.txtDescription.Text
                End If
                Me.LblProduct.Text = Code + " - " + Desc

                'format harga untuk setiap dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvPricing.Rows
                    row.Cells(3).Value = Me.ObjCekData.DigitGrouping(row.Cells(3).Value)
                Next
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                If .ProductType = clsMasterProduct.EnumTipeProduct.Goods Then
                    Me.rboGoods.Checked = True
                    Me.rboService.Checked = False
                Else
                    Me.rboGoods.Checked = False
                    Me.rboService.Checked = True
                End If
                Me.TxtPrimaryInventoryAcc.Text = .PrimaryAkunInventory
                Me.TxtInventoryAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryInventoryAcc.Text)
                Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                Me.TxtPrimarySalesAccount.Text = .PrimaryAkunSales
                Me.TxtSalesAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesAccount.Text)
                Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                Me.TxtPrimarySalesReturnAccount.Text = .PrimaryAkunSalesReturn
                Me.TxtSalesReturnAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesReturnAccount.Text)
                Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                Me.TxtPrimaryCOGSAccount.Text = .PrimaryAkunCOGS
                Me.TxtCOGSAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryCOGSAccount.Text)
                Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                Me.TxtPrimaryAdjustmentAccount.Text = .PrimaryAkunAdjustment
                Me.TxtAdjustmentAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAdjustmentAccount.Text)
                Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                Me.TxtPrimaryStockTakingAccount.Text = .PrimaryAkunStockTaking
                Me.TxtStockTakingAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryStockTakingAccount.Text)
                Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                Me.cboDependency.SelectedIndex = .TipeDependency
                Me.txtDependency.ReadOnly = False
                Me.txtDependency.BackColor = Me.txtPrimaryCategory.BackColor
                Me.txtDependency.Text = Me.ObjCD.Convert_ToString(.DelayDependency)
                Me.SetAccount()

                'add weight & volume
                Me.TxtNettWeight.Text = Me.ObjCD.DigitGrouping(.NettWeight, .PrecisionWeight)
                Me.TxtGrossWeight.Text = Me.ObjCD.DigitGrouping(.GrossWeight, .PrecisionWeight)
                Me.TxtPrimaryUnitWeight.Text = .PrimaryUnitWeight
                Me.TxtUnitWeight.Text = Me.ObjUnit.GetUnitCode(.PrimaryUnitWeight)
                Me.TxtDefaultPrecisionWeight.Text = .PrecisionWeight
                Me.TxtNettVolume.Text = Me.ObjCD.DigitGrouping(.NettVolume, .PrecisionVolume)
                Me.TxtGrossVolume.Text = Me.ObjCD.DigitGrouping(.GrossVolume, .PrecisionVolume)
                Me.TxtPrimaryUnitVolume.Text = .PrimaryUnitVolume
                Me.TxtUnitVolume.Text = Me.ObjUnit.GetUnitCode(.PrimaryUnitVolume)
                Me.TxtDefaultPrecisionVolume.Text = .PrecisionVolume

                Me.TxtPrimaryTaxProductRef.Text = .PrimaryTaxProductRef
                Me.TxtTaxProductRef.Text = ObjTaxProduct.GetCode(.PrimaryTaxProductRef)
                Me.TxtTaxProductDesc.Text = ObjTaxProduct.GetDesc(.PrimaryTaxProductRef)
            End With

            Me.chkSizing.Checked = Me.objProduct.UseSizing
            If Me.objProduct.Size = 36 Then
                Me.cboSizing.SelectedIndex = 0
            ElseIf Me.objProduct.Size = 37 Then
                Me.cboSizing.SelectedIndex = 1
            ElseIf Me.objProduct.Size = 38 Then
                Me.cboSizing.SelectedIndex = 2
            ElseIf Me.objProduct.Size = 39 Then
                Me.cboSizing.SelectedIndex = 3
            ElseIf Me.objProduct.Size = 40 Then
                Me.cboSizing.SelectedIndex = 4
            ElseIf Me.objProduct.Size = 41 Then
                Me.cboSizing.SelectedIndex = 5
            ElseIf Me.objProduct.Size = 42 Then
                Me.cboSizing.SelectedIndex = 6
            ElseIf Me.objProduct.Size = 43 Then
                Me.cboSizing.SelectedIndex = 7
            ElseIf Me.objProduct.Size = 44 Then
                Me.cboSizing.SelectedIndex = 8
            ElseIf Me.objProduct.Size = 45 Then
                Me.cboSizing.SelectedIndex = 9
            Else
                Me.cboSizing.SelectedIndex = 10
            End If

            'berkaitan dengan raw material
            Me.txtPrimaryProses.Text = Me.objProduct.PrimaryProcess
            Me.txtProcessCode.Text = Me.ObjProses.GetCode(Me.txtPrimaryProses.Text)
            Me.TxtProcessDesc.Text = Me.ObjProses.GetDesc(Me.txtPrimaryProses.Text)
            Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping(Me.objProduct.QtyRef, Me.TxtDefaultPrecision.Text)
            Me.TxtGoodPercentage.Text = Me.ObjCD.DigitGrouping(Me.objProduct.GoodPercentage, 2)
            Me.txtTolerance.Text = Me.ObjCD.DigitGrouping(Me.objProduct.Tolerance, 2)
            Me.Update_TreeView()

            'berkaitan dengan verifikasi akun
            Me.TxtIsVerified.Text = Me.objProduct.IsVerified
            Me.TxtVerifiedBy.Text = Me.objProduct.VerifiedBy
            Me.TxtVerifiedTime.Text = Format(Me.objProduct.VerifiedTime, "dd MMM yyyy HH:mm:ss")
            If Me.TxtIsVerified.Text = "0" Then
                Me.TxtAccountVerificationStatus.Text = "Not verified"
            Else
                Me.TxtAccountVerificationStatus.Text = "Verified by: " + Me.TxtVerifiedBy.Text + " on " + Me.TxtVerifiedTime.Text
            End If

            Me.CboTypeFinishGood.SelectedIndex = Me.objProduct.TypeGoods

            If Me.txtUnit.Text = "" Then
                Me.ColWorkUnit_Rate.HeaderText = "Rate (Unit/Hour)"
            Else
                Me.ColWorkUnit_Rate.HeaderText = "Rate (" + Me.txtUnit.Text + "/Hour)"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCat = Nothing
            ObjUnit = Nothing
        End Try
    End Sub

    Private Sub frmMasterProduct_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim DefPrec As Integer = 0
        Dim DataType As Integer

        Try
            For Each rowDgv As DataGridViewRow In Me.DGVKompetensi.Rows
                With rowDgv
                    DefPrec = Me.ObjKompetensi.GetDefPrec(.Cells(Me.ColCompetence_Primary.Index).Value)
                    DataType = Me.ObjKompetensi.GetDataType(.Cells(Me.ColCompetence_Primary.Index).Value)
                    .Cells(Me.ColCompetence_Code.Index).Value = Me.ObjKompetensi.GetDataCode(.Cells(Me.ColCompetence_Primary.Index).Value)
                    .Cells(Me.ColCempetence_Description.Index).Value = Me.ObjKompetensi.GetDesc(.Cells(Me.ColCompetence_Primary.Index).Value)
                    .Cells(Me.ColCompetence_StandarValue.Index).Value = Me.ObjKompetensi.GetValue_CustomFormatted(.Cells(Me.ColCompetence_Primary.Index).Value, Me.ObjCekData, DataType)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjUser As New ClsUser

        Try
            Me.dgvPricing.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem_MultiUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVKompetensi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWorkUnit.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.objProduct = New clsMasterProduct
            Me.baru = True
            Me.NonAktifkanDetail()
            Dim Obj1 As New clsMasterListedSoftware
            If Obj1.IsInventoryInstalled = True Then
                Me.chkBatchNo.Visible = True
                Me.chkExpiryDate.Visible = True
                Me.chkLocation.Visible = True

                'tidak ada modul produksi, chk PDO tidak aktif
                Me.chkLocation.Left = 423
                Me.chkExpiryDate.Left = 199
                Me.chkPDOControlled.Visible = False
                Me.chkPDOControlled.TabStop = False
            Else
                Me.chkBatchNo.Visible = False
                Me.chkExpiryDate.Visible = False
                Me.chkLocation.Visible = False
                Me.chkPDOControlled.Visible = False
            End If

            Me.cboSizing.SelectedIndex = 0
            Obj1 = Nothing

            ''ok-rosadin-r5
            ''dependency
            Me.cboDependency.Items.Clear()
            Me.cboDependency.Items.Add("End-Start")
            Me.cboDependency.Items.Add("Start-Start")
            Me.cboDependency.Items.Add("Start-Start (Strict)")
            Me.cboDependency.SelectedIndex = 0
            Me.txtDependency.ReadOnly = False
            Me.txtDependency.BackColor = Me.txtPrimaryCategory.BackColor

            If Me.EditDrillDown = True Then
                If Me.PrimaryDrillDown <> -1 Then
                    Me.EditCodeNoDrillDown(Me.PrimaryDrillDown)
                    Me.TabControl1.SelectedTab = Me.TabPage4
                Else
                    Me.TabControl1.SelectedTab = Me.TabPage4
                End If
                Exit Sub
            End If
            If Me.NewDrillDown = True Then
                Me.btnNew.PerformClick()
                Exit Sub
            End If


            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "frmMasterProduct_PriceList") = True Then
                Me.LlbEditPriceList.Visible = True
            Else
                Me.LlbEditPriceList.Visible = False
            End If

            Me.ColCempetence_Identifier.Visible = False
            Me.TabPage8.Text = "Work Unit Production Rate"

            ''ok-rosadin-r5
            '-4
            'Me.txtDependency.ReadOnly = True

            ''dependency
            'Me.cboDependency.Items.Clear()
            'Me.cboDependency.Items.Add("End-Start")
            'Me.cboDependency.Items.Add("Start-Start")
            'Me.cboDependency.Items.Add("Start-Start (Strict)")
            'Me.cboDependency.SelectedIndex = 0
            'Me.txtDependency.ReadOnly = False
            'Me.txtDependency.BackColor = Me.txtPrimaryCategory.BackColor
            Me.ColBOM_Tolerance.HeaderText = "Tolerance"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objProduct.SimpanBaru(Me.DGVKompetensi, Me.UserCode, Me.pbxProduct, Me.dgvPricing, Me.dgvItemBOM, Me, Me.dgvItem_MultiUnit, _
                        Me.DGVAttanchment, Me.DgvWorkUnit) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.objProduct.CariDetail(Me.DGVKompetensi, Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
                            Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit)
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objProduct.SimpanPerubahan(Me.DGVKompetensi, Me.UserCode, Me.pbxProduct, Me.dgvPricing, Me.dgvItemBOM, Me, _
                        Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.objProduct.CariDetail(Me.DGVKompetensi, Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
                           Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit)
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
            If Not Me.txtProductCode.Text = "" And Me.txtDescription.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objProduct.CariDetail(Me.DGVKompetensi, Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, Me, _
                        Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit)
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
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TabControl1.SelectedIndex = 0
            Me.txtCategoryCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'cari apakah kode terakhir yang disimpan menggunakan barcode
            If Me.objProduct.ApakahKodeTerakhirMenggunakanBarcode = True Then
                Me.chkBarcode.Checked = False
                Me.chkBarcode.Checked = True
            Else
                Me.chkBarcode.Checked = True
                Me.chkBarcode.Checked = False
            End If

            'secara default product
            Me.rboGoods.Checked = True

            'barang baru, tidak ada verifikasi akun product
            Me.TxtIsVerified.Text = 0
            Me.TxtVerifiedBy.Text = ""
            Me.TxtVerifiedTime.Text = Format(#1/1/1950#, "dd MMM yyyy HH:mm:ss")
            Me.TxtAccountVerificationStatus.Text = "Not verified"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New frmCariMasterProduct

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryProduct = -1 Then
                If Me.objProduct.CariDetail(Me.DGVKompetensi, hj.PrimaryProduct, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
                Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
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

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.txtProductCode.Text = "" Then
            Me.btnFind.PerformClick()
        End If

        If Me.txtProductCode.Text <> "" Then
            Me.baru = False
            Me.TabControl1.SelectedIndex = 0
            Me.AktifkanDetail()
            Me.txtCategoryCode.Focus()
        End If
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtProductCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtProductCode.Text <> "" Then
                If Me.objProduct.ApakahBolehHapus(Me.txtPrimaryProduct.Text) = False Then
                    'MsgBox("You cannot delete this data, there are other people using this data.", _
                    '                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objProduct.Hapus(Me.txtPrimaryProduct.Text, Me.UserCode) Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
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
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtProductCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objProduct.SimpanBaru(Me.DGVKompetensi, Me.UserCode, Me.pbxProduct, Me.dgvPricing, Me.dgvItemBOM, _
                            Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
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
                            If Me.objProduct.SimpanPerubahan(Me.DGVKompetensi, Me.UserCode, Me.pbxProduct, Me.dgvPricing, Me.dgvItemBOM, _
                            Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
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

    Private Sub txtQtyPerMasterUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyPerMasterUnit.Enter
        Try
            Me.buffer = Me.txtQtyPerMasterUnit.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyPerMasterUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyPerMasterUnit.Leave
        Try
            If ModCekData.CekNonNegatif(Me.txtQtyPerMasterUnit.Text) = False Then
                MsgBox("Quantity value must be a number", MsgBoxStyle.Information, "Infomation")
                Me.txtQtyPerMasterUnit.Text = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBrowseImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseImage.Click
        Me.OpenFileDialog1.ShowDialog()
        Try
            If Me.OpenFileDialog1.FileName <> "" Then
                'cek ukuran gambar yang dipilih
                Dim myStream As IO.FileStream = New FileStream(Me.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
                If myStream.Length > 0.5 * 1024 * 1024 Then
                    MsgBox("We only use picture smaller than 500 kb, please change your picture.", MsgBoxStyle.Information, "Information")
                    myStream.Close()
                    Exit Sub
                Else
                    Me.pbxProduct.ImageLocation = Me.OpenFileDialog1.FileName
                End If
                myStream.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.OpenFileDialog1.Dispose()
        End Try
    End Sub

    Private Sub btnClearImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearImage.Click
        Me.pbxProduct.Image = Nothing
    End Sub

    Private Sub txtCategoryCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.Enter
        Try
            Me.buffer = Me.txtPrimaryCategory.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategoryCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCategoryCode.KeyDown
        Dim i As New frmCariMasterCategory
        Dim Obj As New ClsMasterCategory
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                    Me.txtPrimaryCategory.Text = Obj.GetPrimary(Me.txtCategoryCode.Text)
                    Me.txtProductCode.Focus()
                Else
                    If Me.txtCategoryCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCategoryCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCategory > 0 Then
                        Me.txtPrimaryCategory.Text = i.PrimaryCategory
                        Me.txtCategoryCode.Text = Obj.GetCategoryCode(Me.txtPrimaryCategory.Text)
                        Me.txtProductCode.Focus()
                    Else
                        Me.txtCategoryCode.Clear()
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtCategoryCode.Focus()
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = "Product Code"
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = "Description"
                            Me.txtDescription.Clear()
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                        Me.Update_TreeView_FirstNode()
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

    Private Sub txtCategoryCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.Leave
        Dim i As New frmCariMasterCategory
        Dim obj As New ClsMasterCategory
        i.BackColorAll = Me.BackColorAll
        Dim UpdateAkun As Boolean = True

        Try
            If Me.txtCategoryCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                    Me.txtPrimaryCategory.Text = obj.GetPrimary(Me.txtCategoryCode.Text)
                    If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                            Me.txtDescription.Text = Desc
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                        Me.Update_TreeView_FirstNode()
                    End If
                Else
                    If Me.txtCategoryCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCategoryCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCategory > 0 Then
                        Me.txtPrimaryCategory.Text = i.PrimaryCategory
                        Me.txtCategoryCode.Text = obj.GetCategoryCode(Me.txtPrimaryCategory.Text)
                        Me.txtProductCode.Focus()
                        If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                            Dim Code, Desc As String
                            If Me.chkCategoryBindCode.Checked = True Then
                                Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                            Else
                                Code = Me.txtProductCode.Text
                            End If
                            If Me.chkCategoryBindDecs.Checked = True Then
                                Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                                Me.txtDescription.Text = Desc
                            Else
                                Desc = Me.txtDescription.Text
                            End If
                            Me.LblProduct.Text = Code + " - " + Desc
                            Me.Update_TreeView_FirstNode()
                        End If
                    Else
                        Me.txtCategoryCode.Clear()
                        Me.txtPrimaryCategory.Text = "-1"
                        If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                            Dim Code, Desc As String
                            If Me.chkCategoryBindCode.Checked = True Then
                                Code = "Product Code"
                            Else
                                Code = Me.txtProductCode.Text
                            End If
                            If Me.chkCategoryBindDecs.Checked = True Then
                                Desc = "Description"
                                Me.txtDescription.Clear()
                            Else
                                Desc = Me.txtDescription.Text
                            End If
                            Me.LblProduct.Text = Code + " - " + Desc
                            Me.Update_TreeView_FirstNode()
                        End If
                        Me.txtCategoryCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCategory.Text = "-1"
                If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                    Dim Code, Desc As String
                    If Me.chkCategoryBindCode.Checked = True Then
                        Code = "Product Code"
                    Else
                        Code = Me.txtProductCode.Text
                    End If
                    If Me.chkCategoryBindDecs.Checked = True Then
                        Desc = "Description"
                        Me.txtDescription.Clear()
                    Else
                        Desc = Me.txtDescription.Text
                    End If
                    Me.LblProduct.Text = Code + " - " + Desc
                    Me.Update_TreeView_FirstNode()
                End If
            End If

            If Me.buffer <> Me.txtPrimaryCategory.Text Then
                If Me.rboGoods.Checked = True Then
                    If Me.objProduct.ApakahSudahPernahJurnal(Me.txtPrimaryProduct.Text) = True Then
                        UpdateAkun = False
                    End If
                End If
                If UpdateAkun = True Then
                    Me.IsikanAkunDefault()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindCategory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCategory.Click
        Dim i As New frmCariMasterCategory
        Dim obj As New ClsMasterCategory
        Dim UpdateAkun As Boolean = True

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCategoryCode.Text <> "" Then
                i.txtCari.Text = Me.txtCategoryCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCategory > 0 Then
                Me.txtPrimaryCategory.Text = i.PrimaryCategory
                Me.txtCategoryCode.Text = obj.GetCategoryCode(Me.txtPrimaryCategory.Text)
                Me.txtProductCode.Focus()
                If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                    Dim Code, Desc As String
                    If Me.chkCategoryBindCode.Checked = True Then
                        Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                    Else
                        Code = Me.txtProductCode.Text
                    End If
                    If Me.chkCategoryBindDecs.Checked = True Then
                        Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                        Me.txtDescription.Text = Desc
                    Else
                        Desc = Me.txtDescription.Text
                    End If
                    Me.LblProduct.Text = Code + " - " + Desc
                    Me.Update_TreeView_FirstNode()
                End If
            Else
                Me.txtCategoryCode.Clear()
                Me.txtCategoryCode.Focus()
                If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                    Dim Code, Desc As String
                    If Me.chkCategoryBindCode.Checked = True Then
                        Code = "Product Code"
                    Else
                        Code = Me.txtProductCode.Text
                    End If
                    If Me.chkCategoryBindDecs.Checked = True Then
                        Desc = "Description"
                        Me.txtDescription.Clear()
                    Else
                        Desc = Me.txtDescription.Text
                    End If
                    Me.LblProduct.Text = Code + " - " + Desc
                    Me.Update_TreeView_FirstNode()
                End If
            End If

            If Me.buffer <> Me.txtPrimaryCategory.Text Then
                If Me.rboGoods.Checked = True Then
                    If Me.objProduct.ApakahSudahPernahJurnal(Me.txtPrimaryProduct.Text) = True Then
                        UpdateAkun = False
                    End If
                End If
                If UpdateAkun = True Then
                    Me.IsikanAkunDefault()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnit.KeyDown
        Dim i As New FrmCariMasterUnit
        Dim Obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtUnit.Text) = True Then
                    Me.txtPrimaryUnit.Text = Obj.GetPrimary(Me.txtUnit.Text)
                    Me.txtMasterUnit.Focus()
                Else
                    If Me.txtUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnit.Text = i.PrimaryUnit
                        Me.txtUnit.Text = Obj.GetUnitCode(Me.txtPrimaryUnit.Text)
                        Me.txtMasterUnit.Focus()
                    Else
                        Me.txtUnit.Clear()
                        Me.txtPrimaryUnit.Text = "-1"
                        Me.txtUnit.Focus()
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

    Private Sub txtUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnit.Leave
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit
        Dim rowDGV As DataGridViewRow

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtUnit.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtUnit.Text) = True Then
                    Me.txtPrimaryUnit.Text = obj.GetPrimary(Me.txtUnit.Text)
                Else
                    If Me.txtUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryUnit.Text = i.PrimaryUnit
                        Me.txtUnit.Text = obj.GetUnitCode(Me.txtPrimaryUnit.Text)
                        Me.txtMasterUnit.Focus()
                    Else
                        Me.txtUnit.Clear()
                        Me.txtPrimaryUnit.Text = "-1"
                        Me.txtUnit.Focus()
                    End If
                End If

                For Each rowDGV In Me.dgvItem_MultiUnit.Rows
                    If rowDGV.Cells(Me.ColUnit_Unit.Index).Value <> "" Then
                        If rowDGV.Cells(Me.ColUnit_Unit.Index).Value = Me.txtUnit.Text Then
                            MsgBox("Unit code should not be the same as the default unit, please change your code", MsgBoxStyle.Information, "Information")
                            rowDGV.Selected = True
                            Me.txtUnit.Clear()
                            Me.txtUnit.Focus()
                            Exit Sub
                        Else
                            rowDGV.Cells(Me.ColUnit_Desc.Index).Value = rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value _
                            + " " + rowDGV.Cells(Me.ColUnit_Unit.Index).Value _
                            + " = " + FormatNumber(rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value, CInt(Me.TxtDefaultPrecision.Text)) _
                            + " " + Me.txtUnit.Text
                        End If
                    Else
                        rowDGV.Cells(Me.ColUnit_Desc.Index).Value = ""
                    End If
                Next
            Else
                Me.txtPrimaryUnit.Text = "-1"
            End If

            If Me.btnOK.Enabled = True Then
                If Me.txtUnit.Text = "" Then
                    Me.ColWorkUnit_Rate.HeaderText = "Rate (Unit/Hour)"
                Else
                    Me.ColWorkUnit_Rate.HeaderText = "Rate (" + Me.txtUnit.Text + "/Hour)"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindUnit.Click
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.txtUnit.Text <> "" Then
                i.txtCari.Text = Me.txtUnit.Text
            End If

            i.ShowDialog()
            If i.PrimaryUnit > 0 Then
                Me.txtPrimaryUnit.Text = i.PrimaryUnit
                Me.txtUnit.Text = obj.GetUnitCode(Me.txtPrimaryUnit.Text)
                Me.txtMasterUnit.Focus()
            Else
                Me.txtUnit.Clear()
                Me.txtUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtMasterUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMasterUnit.KeyDown
        Dim i As New FrmCariMasterUnit
        Dim Obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtMasterUnit.Text) = True Then
                    Me.txtPrimaryMasterUnit.Text = Obj.GetPrimary(Me.txtMasterUnit.Text)
                    Me.txtQtyPerMasterUnit.Focus()
                Else
                    If Me.txtMasterUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtMasterUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryMasterUnit.Text = i.PrimaryUnit
                        Me.txtMasterUnit.Text = Obj.GetUnitCode(Me.txtPrimaryMasterUnit.Text)
                        Me.txtQtyPerMasterUnit.Focus()
                    Else
                        Me.txtMasterUnit.Clear()
                        Me.txtPrimaryMasterUnit.Text = "-1"
                        Me.txtMasterUnit.Focus()
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

    Private Sub txtMasterUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMasterUnit.Leave
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtMasterUnit.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtMasterUnit.Text) = True Then
                    Me.txtPrimaryMasterUnit.Text = obj.GetPrimary(Me.txtMasterUnit.Text)
                    Me.txtQtyPerMasterUnit.TabStop = True
                    Me.txtQtyPerMasterUnit.ReadOnly = False
                    Me.txtQtyPerMasterUnit.BackColor = Me.txtProductCode.BackColor
                    Me.txtQtyPerMasterUnit.Focus()
                Else
                    If Me.txtMasterUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtMasterUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimaryMasterUnit.Text = i.PrimaryUnit
                        Me.txtMasterUnit.Text = obj.GetUnitCode(Me.txtPrimaryMasterUnit.Text)
                        Me.txtQtyPerMasterUnit.TabStop = True
                        Me.txtQtyPerMasterUnit.ReadOnly = False
                        Me.txtQtyPerMasterUnit.BackColor = Me.txtProductCode.BackColor
                        Me.txtQtyPerMasterUnit.Focus()
                    Else
                        Me.txtMasterUnit.Clear()
                        Me.txtPrimaryMasterUnit.Text = "-1"
                        Me.txtQtyPerMasterUnit.TabStop = False
                        Me.txtQtyPerMasterUnit.ReadOnly = True
                        Me.txtQtyPerMasterUnit.Text = "0"
                        Me.txtQtyPerMasterUnit.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                        Me.txtMasterUnit.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryMasterUnit.Text = "-1"
                Me.txtQtyPerMasterUnit.TabStop = False
                Me.txtQtyPerMasterUnit.ReadOnly = True
                Me.txtQtyPerMasterUnit.Text = "0"
                Me.txtQtyPerMasterUnit.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                Me.chkUseSecondaryUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindMasterUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMasterUnit.Click
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.txtMasterUnit.Text <> "" Then
                i.txtCari.Text = Me.txtMasterUnit.Text
            End If

            i.ShowDialog()
            If i.PrimaryUnit > 0 Then
                Me.txtPrimaryMasterUnit.Text = i.PrimaryUnit
                Me.txtMasterUnit.Text = obj.GetUnitCode(Me.txtPrimaryMasterUnit.Text)
                Me.txtQtyPerMasterUnit.TabStop = True
                Me.txtQtyPerMasterUnit.ReadOnly = False
                Me.txtQtyPerMasterUnit.BackColor = Me.txtProductCode.BackColor
                Me.txtQtyPerMasterUnit.Focus()
            Else
                Me.txtMasterUnit.Clear()
                Me.txtMasterUnit.Focus()
                Me.txtQtyPerMasterUnit.TabStop = False
                Me.txtQtyPerMasterUnit.ReadOnly = True
                Me.txtQtyPerMasterUnit.Text = "0"
                Me.txtQtyPerMasterUnit.BackColor = Me.txtDescriptionPatokanWarna.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtSecondaryUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryUnit.KeyDown
        Dim i As New FrmCariMasterUnit
        Dim Obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtSecondaryUnit.Text) = True Then
                    Me.txtPrimarySecondaryUnit.Text = Obj.GetPrimary(Me.txtSecondaryUnit.Text)
                    If Me.chkBatchNo.Visible = True Then
                        Me.chkBatchNo.Focus()
                    ElseIf Me.chkLocation.Visible = True Then
                        Me.chkLocation.Focus()
                    ElseIf chkPDOControlled.Visible = True Then
                        Me.chkPDOControlled.Focus()
                    Else
                        Me.TabControl1.SelectedIndex = 1
                        Me.txtCurrencyCode.Focus()
                    End If
                Else
                    If Me.txtSecondaryUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtSecondaryUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimarySecondaryUnit.Text = i.PrimaryUnit
                        Me.txtSecondaryUnit.Text = Obj.GetUnitCode(Me.txtPrimarySecondaryUnit.Text)
                        If Me.chkBatchNo.Visible = True Then
                            Me.chkBatchNo.Focus()
                        ElseIf Me.chkLocation.Visible = True Then
                            Me.chkLocation.Focus()
                        ElseIf chkPDOControlled.Visible = True Then
                            Me.chkPDOControlled.Focus()
                        Else
                            Me.TabControl1.SelectedIndex = 1
                            Me.txtCurrencyCode.Focus()
                        End If
                    Else
                        Me.txtSecondaryUnit.Clear()
                        Me.txtPrimarySecondaryUnit.Text = "-1"
                        Me.txtSecondaryUnit.Focus()
                    End If
                End If
            ElseIf e.KeyValue = 9 Then    'tab button
                If Me.chkBatchNo.Visible = True Then
                    Me.chkBatchNo.Focus()
                ElseIf Me.chkLocation.Visible = True Then
                    Me.chkLocation.Focus()
                ElseIf chkPDOControlled.Visible = True Then
                    Me.chkPDOControlled.Focus()
                Else
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub txtSecondaryUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryUnit.Leave
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtSecondaryUnit.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtSecondaryUnit.Text) = True Then
                    Me.txtPrimarySecondaryUnit.Text = obj.GetPrimary(Me.txtSecondaryUnit.Text)
                    If Me.chkBatchNo.Visible = True Then
                        Me.chkBatchNo.Focus()
                    ElseIf Me.chkLocation.Visible = True Then
                        Me.chkLocation.Focus()
                    ElseIf chkPDOControlled.Visible = True Then
                        Me.chkPDOControlled.Focus()
                    Else
                        Me.TabControl1.SelectedIndex = 1
                        Me.txtCurrencyCode.Focus()
                    End If
                Else
                    If Me.txtSecondaryUnit.Text <> "" Then
                        i.txtCari.Text = Me.txtSecondaryUnit.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.txtPrimarySecondaryUnit.Text = i.PrimaryUnit
                        Me.txtSecondaryUnit.Text = obj.GetUnitCode(Me.txtPrimarySecondaryUnit.Text)
                        If Me.chkBatchNo.Visible = True Then
                            Me.chkBatchNo.Focus()
                        ElseIf Me.chkLocation.Visible = True Then
                            Me.chkLocation.Focus()
                        ElseIf chkPDOControlled.Visible = True Then
                            Me.chkPDOControlled.Focus()
                        Else
                            Me.TabControl1.SelectedIndex = 1
                            Me.txtCurrencyCode.Focus()
                        End If
                    Else
                        Me.txtSecondaryUnit.Clear()
                        Me.txtPrimarySecondaryUnit.Text = "-1"
                        Me.txtSecondaryUnit.Focus()
                    End If
                End If
            Else
                If Me.chkBatchNo.Visible = True Then
                    Me.chkBatchNo.Focus()
                ElseIf Me.chkLocation.Visible = True Then
                    Me.chkLocation.Focus()
                ElseIf chkPDOControlled.Visible = True Then
                    Me.chkPDOControlled.Focus()
                Else
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindSecondaryUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSecondaryUnit.Click
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.txtSecondaryUnit.Text <> "" Then
                i.txtCari.Text = Me.txtSecondaryUnit.Text
            End If

            i.ShowDialog()
            If i.PrimaryUnit > 0 Then
                Me.txtPrimarySecondaryUnit.Text = i.PrimaryUnit
                Me.txtSecondaryUnit.Text = obj.GetUnitCode(Me.txtPrimarySecondaryUnit.Text)
                If Me.chkBatchNo.Visible = True Then
                    Me.chkBatchNo.Focus()
                ElseIf Me.chkLocation.Visible = True Then
                    Me.chkLocation.Focus()
                ElseIf chkPDOControlled.Visible = True Then
                    Me.chkPDOControlled.Focus()
                Else
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                End If
            Else
                Me.txtSecondaryUnit.Clear()
                Me.txtSecondaryUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Leave
        Try
            If Me.chkBarcode.Enabled = True Then
                If Me.chkBarcode.Checked = True And Me.TxtBarcode.Text = "" Then
                    If Me.chkCategoryBindCode.Checked = True Then
                        Me.TxtBarcode.Text = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                    Else
                        Me.TxtBarcode.Text = Me.txtProductCode.Text
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductCode.TextChanged
        Try
            If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                If Me.txtPrimaryCategory.Text <> "" Then
                    If Me.txtPrimaryCategory.Text <> "-1" Then
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                        Me.Update_TreeView_FirstNode()
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                        Me.Update_TreeView_FirstNode()
                    End If
                End If
            Else
                Me.LblProduct.Text = Me.txtProductCode.Text + " - " + Me.txtDescription.Text
                Me.Update_TreeView_FirstNode()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSpecification_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSpecification.TextChanged
        Try
            If Me.chkCategoryBindDecs.Checked = True Or Me.chkCategoryBindCode.Checked = True Then
                If Me.txtPrimaryCategory.Text <> "" Then
                    If Me.txtPrimaryCategory.Text <> "-1" Then
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                            Me.txtDescription.Text = Desc
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                        Me.Update_TreeView_FirstNode()
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                        Me.Update_TreeView_FirstNode()
                        Me.txtDescription.Text = ""
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus
        txtProductCode.BackColor = BackColorAll
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus
        txtProductCode.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtCategoryCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.GotFocus
        txtCategoryCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCategoryCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.LostFocus
        txtCategoryCode.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtSpecification_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpecification.GotFocus
        txtSpecification.BackColor = BackColorAll
    End Sub

    Private Sub txtSpecification_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSpecification.LostFocus
        txtSpecification.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.GotFocus
        If Me.txtDescription.ReadOnly = False Then
            txtDescription.BackColor = BackColorAll
        End If
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        If Me.txtDescription.ReadOnly = False Then
            txtDescription.BackColor = Me.txtPrimaryCategory.BackColor
        End If
    End Sub

    Private Sub txtUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnit.GotFocus
        txtUnit.BackColor = BackColorAll
    End Sub

    Private Sub txtUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnit.LostFocus
        txtUnit.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtMasterUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMasterUnit.GotFocus
        txtMasterUnit.BackColor = BackColorAll
    End Sub

    Private Sub txtMasterUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMasterUnit.LostFocus
        txtMasterUnit.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtQtyPerMasterUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyPerMasterUnit.GotFocus
        If Me.txtQtyPerMasterUnit.ReadOnly = False Then
            txtQtyPerMasterUnit.BackColor = BackColorAll
        End If
    End Sub

    Private Sub txtQtyPerMasterUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyPerMasterUnit.LostFocus
        If Me.txtQtyPerMasterUnit.ReadOnly = False Then
            txtQtyPerMasterUnit.BackColor = Me.txtPrimaryProduct.BackColor
        End If
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtDefaultUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitPrice.GotFocus
        txtUnitPrice.BackColor = BackColorAll
    End Sub

    Private Sub txtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnitPrice.KeyDown
        Me.ObjCekData.Number_KeyDown(Me.txtUnitPrice, sender, e)
    End Sub

    Private Sub txtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnitPrice.KeyUp
        Me.ObjCekData.Number_KeyUp(Me.txtUnitPrice, sender, e)
    End Sub

    Private Sub txtDefaultUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitPrice.LostFocus
        txtUnitPrice.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtDefaultDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscount.GotFocus
        txtDiscount.BackColor = BackColorAll
    End Sub

    Private Sub txtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyDown
        Me.ObjCekData.Number_KeyDown(Me.txtDiscount, sender, e)
    End Sub

    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        Me.ObjCekData.Number_KeyUp(Me.txtDiscount, sender, e)
    End Sub

    Private Sub txtDefaultDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscount.LostFocus
        txtDiscount.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtSecondaryUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryUnit.GotFocus
        txtSecondaryUnit.BackColor = BackColorAll
    End Sub

    Private Sub ttxtSecondaryUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryUnit.LostFocus
        txtSecondaryUnit.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtCategoryCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCategoryCode.TextChanged
        Try
            If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                If Me.txtPrimaryCategory.Text <> "" Then
                    If Me.txtPrimaryCategory.Text <> "-1" Then
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                            Me.txtDescription.Text = Desc
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                    End If
                End If
                Me.txtCategoryCode2.Text = Me.txtCategoryCode.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        Dim Obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = Obj.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.txtUnitPrice.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = Obj.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.txtUnitPrice.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
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

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtCurrencyCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtUnitPrice.Focus()
                    Me.txtPrimaryCurrency.Text = obj.GetPrimary(Me.txtCurrencyCode.Text)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = obj.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.txtUnitPrice.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.txtUnitPrice.Focus()
                Me.txtPrimaryCurrency.Text = obj.GetPrimary(Me.txtCurrencyCode.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = obj.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.txtUnitPrice.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtQtyPerMasterUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQtyPerMasterUnit.KeyDown
        Try
            If Me.txtQtyPerMasterUnit.ReadOnly = False Then
                Me.ObjCekData.Number_KeyDown(Me.txtQtyPerMasterUnit, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyPerMasterUnit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQtyPerMasterUnit.KeyUp
        Try
            If Me.txtQtyPerMasterUnit.ReadOnly = False Then
                Me.ObjCekData.Number_KeyUp(Me.txtQtyPerMasterUnit, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Try
            'cek validitas data, data salah bila:
            '1. Jenis currency yang akan dimasukkan sudah ada
            Dim row As DataGridViewRow
            For Each row In Me.dgvPricing.Rows
                If row.Cells(2).Value = Me.txtCurrencyCode.Text Then
                    MsgBox("Currency duplication, you already input pricing data for: " + _
                    Me.txtCurrencyCode.Text, MsgBoxStyle.Information, "Information")
                    Me.txtCurrencyCode.Focus()
                    Exit Sub
                End If
            Next
            '2. nilai diskon > 100
            If CDbl(Me.txtDiscount.Text) > 100 Then
                MsgBox("Maximum discount value is 100%", MsgBoxStyle.Information, "Information")
                Me.txtDiscount.Focus()
                Exit Sub
            End If
            '3. nama mata uang tidak diisi
            If Me.txtCurrencyCode.Text = "" Then
                MsgBox("Currency code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtCurrencyCode.Focus()
                Exit Sub
            End If
            '4. diskon atau harga satuan bernilai negatif
            If Me.ObjCekData.CekNonNegatif(Me.txtUnitPrice.Text) = False Then
                MsgBox("Unit price must be a non-negative number", MsgBoxStyle.Information, "Information")
                Me.txtUnitPrice.Focus()
                Exit Sub
            End If
            If Me.ObjCekData.CekNonNegatif(Me.txtDiscount.Text) = False Then
                MsgBox("Discount must be a non-negative number", MsgBoxStyle.Information, "Information")
                Me.txtDiscount.Focus()
                Exit Sub
            End If
            'masukkan data ke dalam dgv
            Me.dgvPricing.Rows.Add("-", Me.txtPrimaryCurrency.Text, Me.txtCurrencyCode.Text, Me.txtUnitPrice.Text, _
            Me.txtDiscount.Text)
            'renumber
            Me.ObjCekData.SetNumber(Me.dgvPricing)
            'kosongkan data yang diiinput tadi
            Me.txtUnitPrice.Text = "0"
            Me.txtDiscount.Text = "0"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtCurrencyCode.Text = ""
            'fokus kembali
            Me.txtCurrencyCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPricing_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPricing.CellBeginEdit
        Try
            If Me.txtCategoryCode.Enabled = True Then
                'yang bisa diedit hanya cell ke 1, ke 3 dan ke 4/currency, harga dan diskon
                If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Or e.ColumnIndex = 4 Then
                    Me.buffer = Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value
                    Dim Str As String = Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value
                    While Str.Contains(",")
                        Str = Str.Remove(Str.IndexOf(","), 1)
                    End While
                    Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Str
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPricing_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPricing.CellEndEdit
        Try
            Dim Dummy As String = CStr(Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value)
            If e.ColumnIndex = 4 Then
                'untuk nilai persen harus antara 0 dan 100 persen
                If ModCekData.CekNonNegatif(Dummy) = False Then
                    MsgBox("Discount must be number between 0 - 100.", MsgBoxStyle.Information, "Information")
                    Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                ElseIf CDbl(Dummy) > 100 Then
                    MsgBox("Discount must be number between 0 - 100.", MsgBoxStyle.Information, "Information")
                    Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
            ElseIf e.ColumnIndex = 3 Then
                'untuk harga harus non negatif
                If ModCekData.CekNonNegatif(Dummy) = False Then
                    MsgBox("unit price must be a number.", MsgBoxStyle.Information, "Information")
                    Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
            ElseIf e.ColumnIndex = 2 Then
                Dim Obj As New clsMasterCurrency
                If Obj.ApakahSudahAda(Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterCurrency
                    If Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                        i.txtCari.Text = Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency <> -1 Then
                        Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetCurrencyCode(i.PrimaryCurrency)).ToUpper
                        Me.dgvPricing.Item(1, e.RowIndex).Value = i.PrimaryCurrency
                    Else
                        Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    i = Nothing
                Else
                    Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                    Me.dgvPricing.Item(1, e.RowIndex).Value = Obj.GetPrimary(Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
                Obj = Nothing
            End If

            'kalo sampe sini maka atur agar nilainya formatnya sesuai
            Dim Txt As String = Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value
            'hapuskan koma yang ada
            While Txt.Contains(",") = True
                Txt = Txt.Remove(Txt.IndexOf(","), 1)
            End While

            'digit grouping
            Dim Start As Integer = 9
            Dim Finish As Boolean = False

            'tambahkan koma
            While Start - 3 > 0
                'identifikasi titik
                If Txt.IndexOf(".") <> -1 Then
                    Start = Txt.IndexOf(".")
                Else
                    Start = Txt.Length
                End If

                If Txt.Contains(",") = True Then
                    Start = Txt.IndexOf(",")
                End If

                If Start - 3 > 0 Then
                    If Txt.Substring(Start - 3) <> "," Then
                        Txt = Txt.Insert(Start - 3, ",")
                        Start -= 3
                    Else
                        Start -= 3
                    End If
                Else
                    Exit While
                End If
            End While
            Me.dgvPricing.Item(e.ColumnIndex, e.RowIndex).Value = Txt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPricing_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvPricing.UserDeletingRow
        Try
            If Me.txtCategoryCode.Enabled = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseSecondaryUnit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseSecondaryUnit.CheckedChanged
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.txtSecondaryUnit.Enabled = True
                    Me.btnFindSecondaryUnit.Enabled = True
                Else
                    Me.txtSecondaryUnit.Clear()
                    Me.txtSecondaryUnit.Enabled = False
                    Me.txtPrimarySecondaryUnit.Text = "-1"
                    Me.btnFindSecondaryUnit.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtLeadTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadTime.GotFocus
        Me.txtLeadTime.BackColor = BackColorAll
    End Sub

    Private Sub txtLeadTime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeadTime.KeyDown
        Try
            Me.ObjCekData.Number_KeyDown(Me.txtLeadTime, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtLeadTime_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeadTime.KeyUp
        Try
            Me.ObjCekData.Number_KeyUp(Me.txtLeadTime, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtMinimumPurchase_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinimumPurchase.GotFocus
        Me.txtMinimumPurchase.BackColor = BackColorAll
    End Sub

    Private Sub txtMinimumPurchase_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMinimumPurchase.KeyDown
        Try
            Me.ObjCekData.Number_KeyDown(Me.txtMinimumPurchase, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtMinimumPurchase_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMinimumPurchase.KeyUp
        Try
            Me.ObjCekData.Number_KeyUp(Me.txtMinimumPurchase, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtMinimumStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinimumStock.GotFocus
        Me.txtMinimumStock.BackColor = BackColorAll
    End Sub

    Private Sub txtMinimumStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMinimumStock.KeyDown
        Try
            Me.ObjCekData.Number_KeyDown(Me.txtMinimumStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtMinimumStock_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMinimumStock.KeyUp
        Try
            Me.ObjCekData.Number_KeyUp(Me.txtMinimumStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtMinimumPurchase_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinimumPurchase.LostFocus
        Me.txtMinimumPurchase.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtMinimumStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinimumStock.LostFocus
        Me.txtMinimumStock.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtLeadTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadTime.LostFocus
        Me.txtLeadTime.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        Try
            If Me.chkCategoryBindCode.Checked = True Or Me.chkCategoryBindDecs.Checked = True Then
                If Me.txtPrimaryCategory.Text <> "" Then
                    If Me.txtPrimaryCategory.Text <> "-1" Then
                        Dim Code, Desc As String
                        If Me.chkCategoryBindCode.Checked = True Then
                            Code = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Code = Me.txtProductCode.Text
                        End If
                        If Me.chkCategoryBindDecs.Checked = True Then
                            Desc = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " + Me.txtSpecification.Text
                        Else
                            Desc = Me.txtDescription.Text
                        End If
                        Me.LblProduct.Text = Code + " - " + Desc
                        Me.Update_TreeView_FirstNode()
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                        Me.Update_TreeView_FirstNode()
                    End If
                End If
            Else
                Me.LblProduct.Text = Me.txtProductCode.Text + " - " + Me.txtDescription.Text
                Me.Update_TreeView_FirstNode()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkBatchNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBatchNo.CheckedChanged
        'expiry date controlled hanya bisa dilakukan ketika user melakukan batch no controlled
        Try
            If Me.chkBatchNo.Checked = True Then
                Me.chkExpiryDate.Enabled = True
            Else
                Me.chkExpiryDate.Enabled = False
                Me.chkExpiryDate.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseSecondaryUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseSecondaryUnit.Leave
        Try
            If Me.chkUseSecondaryUnit.Checked = False Then
                If Me.chkBatchNo.Visible = True Then
                    Me.chkBatchNo.Focus()
                ElseIf Me.chkLocation.Visible = True Then
                    Me.chkLocation.Focus()
                ElseIf chkPDOControlled.Visible = True Then
                    Me.chkPDOControlled.Focus()
                Else
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCategoryBindCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCategoryBindCode.CheckedChanged
        Try
            If Me.chkCategoryBindCode.Checked = True Then
                Me.txtCategoryCode2.Text = Me.txtCategoryCode.Text
                Me.txtCategoryCode2.Visible = True
                Me.txtProductCode.Left = Me.txtCategoryCode2.Left + Me.txtCategoryCode2.Width + 5
                Me.txtProductCode.Width = Me.txtDescription.Width - (Me.txtProductCode.Left - Me.txtDescription.Left) - _
                (Me.chkBarcode.Width + Me.TxtBarcode.Width + (Me.TxtBarcode.Left - Me.chkBarcode.Left - Me.chkBarcode.Width) + 5)
                If Me.txtPrimaryCategory.Text <> "-1" Then
                    If Me.txtProductCode.Text.StartsWith(Me.txtCategoryCode2.Text) = True Then
                        Me.txtProductCode.Text = Me.txtProductCode.Text.Substring(Me.txtCategoryCode2.Text.Length)
                    End If
                End If
                Me.txtProductCode_TextChanged(sender, e)
            Else
                Me.txtCategoryCode2.Visible = False
                Me.txtProductCode.Left = Me.txtCategoryCode.Left
                Me.txtProductCode.Width = Me.txtDescription.Width - _
                (Me.chkBarcode.Width + Me.TxtBarcode.Width + (Me.TxtBarcode.Left - Me.chkBarcode.Left - Me.chkBarcode.Width) + 5)
                Me.txtProductCode.Text = Me.txtCategoryCode2.Text + Me.txtProductCode.Text
                Me.txtProductCode_TextChanged(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCategoryBindDecs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCategoryBindDecs.CheckedChanged
        Try
            If Me.chkCategoryBindDecs.Checked = True Then
                Me.txtDescription.ReadOnly = True
                Me.txtDescription.TabStop = False
                If Me.txtPrimaryCategory.Text <> "" Then
                    Me.txtDescription.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                    Me.txtDescription.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text) + " " _
                    + Me.txtSpecification.Text
                End If
                Me.txtProductCode_TextChanged(sender, e)
            Else
                Me.txtDescription.BackColor = Me.txtPrimaryCategory.BackColor
                Me.txtDescription.ReadOnly = False
                Me.txtDescription.TabStop = True
                Me.txtProductCode_TextChanged(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDataAkun(ByVal row_AccGroup As DseMasterProductAccountGroup.tblmaster_product_account_groupRow, _
    ByVal TipeCodeNo As ClsMasterProductAccountGroup.EnumTipe)
        'ByVal GroupLevel As ClsMasterProductAccountGroup.EnumGroupLevel, _
        Try
            If TipeCodeNo = ClsMasterProductAccountGroup.EnumTipe.Goods Then
                Me.rboGoods.Checked = True

                Me.TxtPrimaryAdjustmentAccount.Text = row_AccGroup.primary_adjustment_account
                Me.TxtAdjustmentAccount.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_adjustment_account)
                Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_adjustment_account)
                Me.TxtPrimaryCOGSAccount.Text = row_AccGroup.primary_cost_of_sales_account
                Me.TxtCOGSAccount.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_cost_of_sales_account)
                Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_cost_of_sales_account)
                Me.TxtPrimaryStockTakingAccount.Text = row_AccGroup.primary_stock_taking_account
                Me.TxtStockTakingAccount.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_stock_taking_account)
                Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_stock_taking_account)
            Else
                Me.rboService.Checked = True
            End If

            Me.TxtPrimaryInventoryAcc.Text = row_AccGroup.primary_inventory_account
            Me.TxtInventoryAcc.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_inventory_account)
            Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_inventory_account)
            Me.TxtPrimarySalesAccount.Text = row_AccGroup.primary_sales_account
            Me.TxtSalesAccount.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_sales_account)
            Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_sales_account)
            Me.TxtPrimarySalesReturnAccount.Text = row_AccGroup.primary_sales_return_account
            Me.TxtSalesReturnAccount.Text = Me.ObjAkun.GetAccountCode(row_AccGroup.primary_sales_return_account)
            Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(row_AccGroup.primary_sales_return_account)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.txtCategoryCode.Enabled = True Then
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryProduct.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.txtPrimaryProduct.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryProduct.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage5.Controls, Me.txtPrimaryProduct.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)

                If Me.TabControl1.SelectedIndex = 0 Then
                    Me.txtCategoryCode.Focus()
                ElseIf Me.TabControl1.SelectedIndex = 1 Then
                    Me.txtCurrencyCode.Focus()
                ElseIf Me.TabControl1.SelectedIndex = 2 Then
                    If Me.TxtInventoryAcc.ReadOnly = False And Me.TxtInventoryAcc.Enabled = True Then
                        Me.TxtInventoryAcc.Focus()
                    Else
                        Me.TxtSalesAccount.Focus()
                    End If
                End If
            Else
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage5.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            End If

            If Me.TabControl1.SelectedTab.Name = Me.TabPage3.Name Then
                'jika kode akun inventory-nya kosong, maka langsung isikan tergantung sub group
                If Me.btnOK.Enabled = True Then
                    If Me.txtPrimaryCategory.Text <> "-1" And Me.txtPrimaryCategory.Text <> "" And _
                    (Me.TxtPrimaryInventoryAcc.Text = "-1" Or Me.TxtPrimaryInventoryAcc.Text = "") Then
                        Me.IsikanAkunDefault()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanAkunDefault()
        Dim DTA_AccGroup As New DseMasterProductAccountGroupTableAdapters.tblmaster_product_account_groupTableAdapter
        Dim TBL_AccGroup As New DseMasterProductAccountGroup.tblmaster_product_account_groupDataTable
        Dim row_AccGroup As DseMasterProductAccountGroup.tblmaster_product_account_groupRow
        Dim DTA_ProductGroup As New dseMasterProductTableAdapters.DataTable6TableAdapter
        Dim TBL_ProductGroup As New dseMasterProduct.DataTable6DataTable
        Dim row_ProductGroup As dseMasterProduct.DataTable6Row

        Dim Ketemu As Boolean = False
        Dim GroupLevel As ClsMasterProductAccountGroup.EnumGroupLevel = ClsMasterProductAccountGroup.EnumGroupLevel.Category
        Dim TipeCodeNo As ClsMasterProductAccountGroup.EnumTipe = ClsMasterProductAccountGroup.EnumTipe.Goods

        Try
            DTA_ProductGroup.Fill(TBL_ProductGroup, Me.txtPrimaryCategory.Text)
            DTA_AccGroup.Fill(TBL_AccGroup)

            If TBL_ProductGroup.Rows.Count > 0 And TBL_AccGroup.Rows.Count > 0 Then
                row_ProductGroup = TBL_ProductGroup.Rows(0)
                'cek category
                GroupLevel = ClsMasterProductAccountGroup.EnumGroupLevel.Category
                TBL_AccGroup.DefaultView.RowFilter = "[group level] = " + CInt(GroupLevel).ToString + _
                " AND [primary group] = " + row_ProductGroup.Primary_Category.ToString
                If TBL_AccGroup.DefaultView.Count > 0 Then
                    row_AccGroup = TBL_AccGroup.DefaultView(0).Row
                    Ketemu = True
                    TipeCodeNo = row_AccGroup.type
                    Me.SalinDataAkun(row_AccGroup, TipeCodeNo)
                End If

                'cek sub brand
                If Ketemu = False Then
                    GroupLevel = ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
                    TBL_AccGroup.DefaultView.RowFilter = "[group level] = " + CInt(GroupLevel).ToString + _
                    " AND [primary group] = " + row_ProductGroup.Primary_Sub_Brand.ToString
                    If TBL_AccGroup.DefaultView.Count > 0 Then
                        row_AccGroup = TBL_AccGroup.DefaultView(0).Row
                        Ketemu = True
                        TipeCodeNo = row_AccGroup.type
                        Me.SalinDataAkun(row_AccGroup, TipeCodeNo)
                    End If
                End If

                'cek brand
                If Ketemu = False Then
                    GroupLevel = ClsMasterProductAccountGroup.EnumGroupLevel.Brand
                    TBL_AccGroup.DefaultView.RowFilter = "[group level] = " + CInt(GroupLevel).ToString + _
                    " AND [primary group] = " + row_ProductGroup.Primary_Brand.ToString
                    If TBL_AccGroup.DefaultView.Count > 0 Then
                        row_AccGroup = TBL_AccGroup.DefaultView(0).Row
                        Ketemu = True
                        TipeCodeNo = row_AccGroup.type
                        Me.SalinDataAkun(row_AccGroup, TipeCodeNo)
                    End If
                End If

                'cek product group
                If Ketemu = False Then
                    GroupLevel = ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
                    TBL_AccGroup.DefaultView.RowFilter = "[group level] = " + CInt(GroupLevel).ToString + _
                    " AND [primary group] = " + row_ProductGroup.Primary_Product_Group.ToString
                    If TBL_AccGroup.DefaultView.Count > 0 Then
                        row_AccGroup = TBL_AccGroup.DefaultView(0).Row
                        Ketemu = True
                        TipeCodeNo = row_AccGroup.type
                        Me.SalinDataAkun(row_AccGroup, TipeCodeNo)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_AccGroup.Dispose()
            TBL_AccGroup.Dispose()
            DTA_ProductGroup.Dispose()
            TBL_ProductGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtDefaultPrecision_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.Enter
        Try
            If Me.btnNew.Enabled = False Then
                Me.buffer = Me.TxtDefaultPrecision.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecision_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecision.KeyDown
        Me.ObjCekData.Number_KeyDown(Me.TxtDefaultPrecision, sender, e)
    End Sub

    Private Sub TxtDefaultPrecision_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecision.KeyUp
        Me.ObjCekData.Number_KeyUp(Me.TxtDefaultPrecision, sender, e, False)
    End Sub

    Private Sub TxtDefaultPrecision_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.GotFocus
        TxtDefaultPrecision.BackColor = BackColorAll
        'Me.buffer = Me.TxtDefaultPrecision.Text
    End Sub

    Private Sub TxtDefaultPrecision_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.LostFocus
        TxtDefaultPrecision.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub TxtDefaultPrecision_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.Leave
        Dim rowDGV As DataGridViewRow

        Try
            If CDbl(Me.TxtDefaultPrecision.Text) > 6 Then
                MsgBox("Please input default precision less than 7", MsgBoxStyle.Information, "Information")
                Me.TxtDefaultPrecision.Text = Me.buffer
                Exit Sub
            Else
                If Me.btnNew.Enabled = False Then
                    If Me.baru = False Then
                        If Me.txtPrimaryUnit.Text <> -1 Then
                            If CInt(Me.TxtDefaultPrecision.Text) <> CInt(Me.buffer) Then
                                Dim DecProd As Integer = Me.TxtDefaultPrecision.Text
                                If Me.objProduct.ApakahBolehEditDecimalProduct(Me.txtPrimaryProduct.Text, Me.txtPrimaryUnit.Text, DecProd) = False Then
                                    MsgBox("Minimum value you must can enter is  " & Me.buffer, MsgBoxStyle.Information, "Information")
                                    Me.TxtDefaultPrecision.Text = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox("Please specify unit code.", MsgBoxStyle.Information, "Information")
                            Me.TxtDefaultPrecision.Text = Me.buffer
                        End If
                    End If
                End If

                For Each rowDGV In Me.dgvItem_MultiUnit.Rows
                    If rowDGV.Cells(Me.ColUnit_Unit.Index).Value <> "" Then
                        rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value = ObjCD.DigitGrouping(rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value, Me.TxtDefaultPrecision.Text)
                        'rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value = ObjCD.DigitGrouping(rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value, rowDGV.Cells(Me.ColUnit_Precision.Index).Value)
                        If rowDGV.Cells(Me.ColUnit_Ratio.Index).Value <> 0 Then
                            rowDGV.Cells(Me.ColUnit_Ratio.Index).Value = rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value / rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value
                            rowDGV.Cells(Me.ColUnit_Ratio.Index).Value = ObjCD.DigitGrouping(rowDGV.Cells(Me.ColUnit_Ratio.Index).Value, 6)
                            rowDGV.Cells(Me.ColUnit_Ratio.Index).Value = ObjCD.BuangNolDibelakangDecimal(rowDGV.Cells(Me.ColUnit_Ratio.Index).Value)
                        End If
                        rowDGV.Cells(Me.ColUnit_Desc.Index).Value = _
                        rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value + " " + rowDGV.Cells(Me.ColUnit_Unit.Index).Value + " = " + _
                        rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value + " " + Me.txtUnit.Text
                    End If
                Next
            End If

            'update default precision pada bill of material + rate per hour
            If Me.txtCategoryCode.Enabled = True Then
                Me.TxtDefaultPrecision.Text = Math.Floor(CDbl(Me.TxtDefaultPrecision.Text))
                Me.TxtQtyRef.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyRef.Text, Me.TxtDefaultPrecision.Text)

                For Each rowDGV In Me.DgvWorkUnit.Rows
                    rowDGV.Cells(Me.ColWorkUnit_Rate.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColWorkUnit_Rate.Index).Value, Me.TxtDefaultPrecision.Text)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub chkSizing_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSizing.CheckedChanged
        Try
            If Me.txtUnit.Enabled = True Then
                If Me.chkSizing.Checked = True Then
                    Me.cboSizing.Enabled = True
                Else
                    Me.cboSizing.SelectedIndex = 0
                    Me.cboSizing.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBarcode.CheckedChanged
        Try
            If Me.TxtDefaultPrecision.Enabled = True Then
                If Me.chkBarcode.Checked = True Then
                    Me.TxtBarcode.ReadOnly = False
                    Me.TxtBarcode.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.TxtBarcode.TabStop = True
                    If Me.TxtBarcode.Text = "" Then
                        If Me.chkCategoryBindCode.Checked = True Then
                            Me.TxtBarcode.Text = Me.txtCategoryCode.Text + Me.txtProductCode.Text
                        Else
                            Me.TxtBarcode.Text = Me.txtProductCode.Text
                        End If
                    End If
                Else
                    Me.TxtBarcode.Clear()
                    Me.TxtBarcode.ReadOnly = True
                    Me.TxtBarcode.TabStop = False
                    Me.TxtBarcode.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBarcode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBarcode.GotFocus
        Try
            If Me.TxtDefaultPrecision.Enabled = True Then
                If Me.TxtBarcode.ReadOnly = True Then
                    Me.TxtBarcode.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                Else
                    Me.TxtBarcode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBarcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBarcode.LostFocus
        Try
            If Me.TxtDefaultPrecision.Enabled = True Then
                If Me.TxtBarcode.ReadOnly = True Then
                    Me.TxtBarcode.BackColor = Me.txtDescriptionPatokanWarna.BackColor
                Else
                    Me.TxtBarcode.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboGoods_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboGoods.CheckedChanged
        Try
            Me.SetAccount()
            Me.SetTabBOM()
            If Me.btnOK.Enabled = True Then
                If Me.rboGoods.Checked = True Then
                    Me.CboTypeFinishGood.Enabled = True
                Else
                    Me.CboTypeFinishGood.Enabled = False
                    Me.CboTypeFinishGood.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetAccount()
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.rboGoods.Checked = False Then
                    Me.TxtCOGSAccount.Clear()
                    Me.TxtCOGSAccountDesc.Clear()
                    Me.TxtPrimaryCOGSAccount.Text = "-1"
                    Me.TxtCOGSAccount.ReadOnly = True
                    Me.TxtCOGSAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindCOGSAccount.Enabled = False
                    Me.TxtStockTakingAccount.Clear()
                    Me.TxtStockTakingAccountDesc.Clear()
                    Me.TxtPrimaryStockTakingAccount.Text = "-1"
                    Me.TxtStockTakingAccount.ReadOnly = True
                    Me.TxtStockTakingAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindStockTakingAccount.Enabled = False
                    Me.TxtAdjustmentAccount.Clear()
                    Me.TxtAdjustmentAccountDesc.Clear()
                    Me.TxtPrimaryAdjustmentAccount.Text = "-1"
                    Me.TxtAdjustmentAccount.ReadOnly = True
                    Me.TxtAdjustmentAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindAdjustmentAccount.Enabled = False
                    Me.Label17.Text = "Expense Account"
                Else
                    Me.TxtCOGSAccount.ReadOnly = False
                    Me.TxtCOGSAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.btnFindCOGSAccount.Enabled = True
                    Me.TxtStockTakingAccount.ReadOnly = False
                    Me.TxtStockTakingAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.btnFindStockTakingAccount.Enabled = True
                    Me.TxtAdjustmentAccount.ReadOnly = False
                    Me.TxtAdjustmentAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.btnFindAdjustmentAccount.Enabled = True
                    Me.Label17.Text = "Inventory Account"
                End If
            Else
                If Me.rboGoods.Checked = False Then
                    Me.TxtCOGSAccount.Clear()
                    Me.TxtCOGSAccountDesc.Clear()
                    Me.TxtPrimaryCOGSAccount.Text = "-1"
                    Me.TxtCOGSAccount.ReadOnly = True
                    Me.TxtCOGSAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtStockTakingAccount.Clear()
                    Me.TxtStockTakingAccountDesc.Clear()
                    Me.TxtPrimaryStockTakingAccount.Text = "-1"
                    Me.TxtStockTakingAccount.ReadOnly = True
                    Me.TxtStockTakingAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtAdjustmentAccount.Clear()
                    Me.TxtAdjustmentAccountDesc.Clear()
                    Me.TxtPrimaryAdjustmentAccount.Text = "-1"
                    Me.TxtAdjustmentAccount.ReadOnly = True
                    Me.TxtAdjustmentAccount.BackColor = Me.TxtReadOnly.BackColor
                    Me.Label17.Text = "Expense Account"
                Else
                    Me.TxtCOGSAccount.ReadOnly = False
                    Me.TxtCOGSAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.TxtStockTakingAccount.ReadOnly = False
                    Me.TxtStockTakingAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.TxtAdjustmentAccount.ReadOnly = False
                    Me.TxtAdjustmentAccount.BackColor = Me.txtPrimaryProduct.BackColor
                    Me.Label17.Text = "Inventory Account"
                End If
            End If
            If Me.rboGoods.Checked = False Then
                Me.chkBatchNo.Checked = False
                Me.chkExpiryDate.Checked = False
                Me.chkLocation.Checked = False
                Me.chkSizing.Checked = False

                If Me.txtUnit.Enabled = True Then
                    Me.chkBatchNo.Enabled = False
                    Me.chkExpiryDate.Enabled = False
                    Me.chkLocation.Enabled = False
                    Me.chkSizing.Enabled = False
                End If
            Else
                If Me.txtUnit.Enabled = True Then
                    Me.chkBatchNo.Enabled = True
                    If Me.chkBatchNo.Checked = True Then
                        Me.chkExpiryDate.Enabled = True
                    Else
                        Me.chkExpiryDate.Enabled = False
                    End If
                    Me.chkLocation.Enabled = True
                    Me.chkSizing.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInventoryAcc_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInventoryAcc.Enter
        Try
            Me.buffer = Me.TxtInventoryAcc.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInventoryAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInventoryAcc.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtInventoryAcc.ReadOnly = False Then
                    Me.TxtInventoryAcc.BackColor = Me.BackColorAll
                Else
                    Me.TxtInventoryAcc.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInventoryAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInventoryAcc.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtInventoryAcc.ReadOnly = False Then
                    Me.TxtInventoryAcc.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtInventoryAcc.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtInventoryAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtInventoryAcc.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindInventoryAcc.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtInventoryAcc.Text) = True Then
                    Me.TxtPrimaryInventoryAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtInventoryAcc.Text)
                    Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                    Me.TxtSalesAccount.Focus()
                Else
                    If Me.TxtInventoryAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtInventoryAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryInventoryAcc.Text = i.mPrimaryAkun
                        Me.TxtInventoryAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryInventoryAcc.Text)
                        Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                        Me.TxtSalesAccount.Focus()
                    Else
                        Me.TxtInventoryAcc.Clear()
                        Me.TxtInventoryAccountDesc.Clear()
                        Me.TxtPrimaryInventoryAcc.Text = "-1"
                        Me.TxtInventoryAcc.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtInventoryAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInventoryAcc.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtInventoryAcc.Text = Me.buffer
                Me.btnFindInventoryAcc.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtInventoryAcc.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtInventoryAcc.Text) = True Then
                    Me.TxtPrimaryInventoryAcc.Text = Me.ObjAkun.GetPrimary(Me.TxtInventoryAcc.Text)
                    Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                Else
                    If Me.TxtInventoryAcc.Text <> "" Then
                        i.txtCari.Text = Me.TxtInventoryAcc.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryInventoryAcc.Text = i.mPrimaryAkun
                        Me.TxtInventoryAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryInventoryAcc.Text)
                        Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                        Me.TxtSalesAccount.Focus()
                    Else
                        Me.TxtInventoryAcc.Clear()
                        Me.TxtInventoryAccountDesc.Clear()
                        Me.TxtPrimaryInventoryAcc.Text = "-1"
                        Me.TxtInventoryAcc.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryInventoryAcc.Text = "-1"
                Me.TxtInventoryAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindInventoryAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindInventoryAcc.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtInventoryAcc.Text <> "" Then
                i.txtCari.Text = Me.TxtInventoryAcc.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryInventoryAcc.Text = i.mPrimaryAkun
                Me.TxtInventoryAcc.Text = obj.GetAccountCode(Me.TxtPrimaryInventoryAcc.Text)
                Me.TxtInventoryAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryInventoryAcc.Text)
                Me.TxtSalesAccount.Focus()
            Else
                Me.TxtInventoryAcc.Clear()
                Me.TxtInventoryAccountDesc.Clear()
                Me.TxtInventoryAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtSalesAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAccount.Enter
        Try
            Me.buffer = Me.TxtSalesAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAccount.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtSalesAccount.ReadOnly = False Then
                    Me.TxtSalesAccount.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAccount.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtSalesAccount.ReadOnly = False Then
                    Me.TxtSalesAccount.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtSalesAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindSalesAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesAccount.Text) = True Then
                    Me.TxtPrimarySalesAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtSalesAccount.Text)
                    Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                    Me.TxtSalesReturnAccount.Focus()
                Else
                    If Me.TxtSalesAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesAccount.Text = i.mPrimaryAkun
                        Me.TxtSalesAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesAccount.Text)
                        Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                        Me.TxtSalesReturnAccount.Focus()
                    Else
                        Me.TxtSalesAccount.Clear()
                        Me.TxtSalesAccountDesc.Clear()
                        Me.TxtPrimarySalesAccount.Text = "-1"
                        Me.TxtSalesAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtSalesAccount.Text = Me.buffer
                Me.btnFindSalesAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesAccount.Text) = True Then
                    Me.TxtPrimarySalesAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtSalesAccount.Text)
                    Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                Else
                    If Me.TxtSalesAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesAccount.Text = i.mPrimaryAkun
                        Me.TxtSalesAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesAccount.Text)
                        Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                        Me.TxtSalesReturnAccount.Focus()
                    Else
                        Me.TxtSalesAccount.Clear()
                        Me.TxtSalesAccountDesc.Clear()
                        Me.TxtPrimarySalesAccount.Text = "-1"
                        Me.TxtSalesAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimarySalesAccount.Text = "-1"
                Me.TxtSalesAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtSalesAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimarySalesAccount.Text = i.mPrimaryAkun
                Me.TxtSalesAccount.Text = obj.GetAccountCode(Me.TxtPrimarySalesAccount.Text)
                Me.TxtSalesAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesAccount.Text)
                Me.TxtSalesReturnAccount.Focus()
            Else
                Me.TxtSalesAccount.Clear()
                Me.TxtSalesAccountDesc.Clear()
                Me.TxtSalesAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtSalesReturnAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnAccount.Enter
        Try
            Me.buffer = Me.TxtSalesReturnAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnAccount.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtSalesReturnAccount.ReadOnly = False Then
                    Me.TxtSalesReturnAccount.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesReturnAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnAccount.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtSalesReturnAccount.ReadOnly = False Then
                    Me.TxtSalesReturnAccount.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtSalesReturnAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesReturnAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesReturnAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan sales return 
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindSalesReturnAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesReturnAccount.Text) = True Then
                    Me.TxtPrimarySalesReturnAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtSalesReturnAccount.Text)
                    Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                    If Me.TxtCOGSAccount.Enabled = True And Me.TxtCOGSAccount.ReadOnly = False Then
                        Me.TxtCOGSAccount.Focus()
                    Else
                        Me.btnOK.Focus()
                    End If
                Else
                    If Me.TxtSalesReturnAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesReturnAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesReturnAccount.Text = i.mPrimaryAkun
                        Me.TxtSalesReturnAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesReturnAccount.Text)
                        Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                        If Me.TxtCOGSAccount.Enabled = True And Me.TxtCOGSAccount.ReadOnly = False Then
                            Me.TxtCOGSAccount.Focus()
                        Else
                            Me.btnOK.Focus()
                        End If
                    Else
                        Me.TxtSalesReturnAccount.Clear()
                        Me.TxtSalesReturnAccountDesc.Clear()
                        Me.TxtPrimarySalesReturnAccount.Text = "-1"
                        Me.TxtSalesReturnAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesReturnAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesReturnAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtSalesReturnAccount.Text = Me.buffer
                Me.btnFindSalesReturnAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesReturnAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesReturnAccount.Text) = True Then
                    Me.TxtPrimarySalesReturnAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtSalesReturnAccount.Text)
                    Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                Else
                    If Me.TxtSalesReturnAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesReturnAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesReturnAccount.Text = i.mPrimaryAkun
                        Me.TxtSalesReturnAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesReturnAccount.Text)
                        Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                        If Me.TxtCOGSAccount.Enabled = True And Me.TxtCOGSAccount.ReadOnly = False Then
                            Me.TxtCOGSAccount.Focus()
                        Else
                            Me.btnOK.Focus()
                        End If
                    Else
                        Me.TxtSalesReturnAccount.Clear()
                        Me.TxtSalesReturnAccountDesc.Clear()
                        Me.TxtPrimarySalesReturnAccount.Text = "-1"
                        Me.TxtSalesReturnAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimarySalesReturnAccount.Text = "-1"
                Me.TxtSalesReturnAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesReturnAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesReturnAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtSalesReturnAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtSalesReturnAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimarySalesReturnAccount.Text = i.mPrimaryAkun
                Me.TxtSalesReturnAccount.Text = obj.GetAccountCode(Me.TxtPrimarySalesReturnAccount.Text)
                Me.TxtSalesReturnAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesReturnAccount.Text)
                If Me.TxtCOGSAccount.Enabled = True And Me.TxtCOGSAccount.ReadOnly = False Then
                    Me.TxtCOGSAccount.Focus()
                Else
                    Me.btnOK.Focus()
                End If
            Else
                Me.TxtSalesReturnAccount.Clear()
                Me.TxtSalesReturnAccountDesc.Clear()
                Me.TxtSalesReturnAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtCOGSAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCOGSAccount.Enter
        Try
            Me.buffer = Me.TxtCOGSAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCOGSAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCOGSAccount.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtCOGSAccount.ReadOnly = False Then
                    Me.TxtCOGSAccount.BackColor = Me.BackColorAll
                Else
                    Me.TxtCOGSAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCOGSAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCOGSAccount.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtCOGSAccount.ReadOnly = False Then
                    Me.TxtCOGSAccount.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtCOGSAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCOGSAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCOGSAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan cogs
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindCOGSAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCOGSAccount.Text) = True Then
                    Me.TxtPrimaryCOGSAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtCOGSAccount.Text)
                    Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                    Me.TxtAdjustmentAccount.Focus()
                Else
                    If Me.TxtCOGSAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtCOGSAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCOGSAccount.Text = i.mPrimaryAkun
                        Me.TxtCOGSAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryCOGSAccount.Text)
                        Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                        Me.TxtAdjustmentAccount.Focus()
                    Else
                        Me.TxtCOGSAccount.Clear()
                        Me.TxtCOGSAccountDesc.Clear()
                        Me.TxtPrimaryCOGSAccount.Text = "-1"
                        Me.TxtCOGSAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCOGSAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCOGSAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtCOGSAccount.Text = Me.buffer
                Me.btnFindCOGSAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtCOGSAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCOGSAccount.Text) = True Then
                    Me.TxtPrimaryCOGSAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtCOGSAccount.Text)
                    Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                Else
                    If Me.TxtCOGSAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtCOGSAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCOGSAccount.Text = i.mPrimaryAkun
                        Me.TxtCOGSAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryCOGSAccount.Text)
                        Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                        Me.TxtAdjustmentAccount.Focus()
                    Else
                        Me.TxtCOGSAccount.Clear()
                        Me.TxtCOGSAccountDesc.Clear()
                        Me.TxtPrimaryCOGSAccount.Text = "-1"
                        Me.TxtCOGSAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCOGSAccount.Text = "-1"
                Me.TxtCOGSAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCOGSAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCOGSAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtCOGSAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtCOGSAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCOGSAccount.Text = i.mPrimaryAkun
                Me.TxtCOGSAccount.Text = obj.GetAccountCode(Me.TxtPrimaryCOGSAccount.Text)
                Me.TxtCOGSAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCOGSAccount.Text)
                Me.TxtAdjustmentAccount.Focus()
            Else
                Me.TxtCOGSAccount.Clear()
                Me.TxtCOGSAccountDesc.Clear()
                Me.TxtCOGSAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtAdjustmentAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAdjustmentAccount.Enter
        Try
            Me.buffer = Me.TxtAdjustmentAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAdjustmentAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAdjustmentAccount.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtAdjustmentAccount.ReadOnly = False Then
                    Me.TxtAdjustmentAccount.BackColor = Me.BackColorAll
                Else
                    Me.TxtAdjustmentAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAdjustmentAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAdjustmentAccount.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtAdjustmentAccount.ReadOnly = False Then
                    Me.TxtAdjustmentAccount.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtAdjustmentAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAdjustmentAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAdjustmentAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan adjustment
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindAdjustmentAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAdjustmentAccount.Text) = True Then
                    Me.TxtPrimaryAdjustmentAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtAdjustmentAccount.Text)
                    Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                    Me.TxtStockTakingAccount.Focus()
                Else
                    If Me.TxtAdjustmentAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtAdjustmentAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAdjustmentAccount.Text = i.mPrimaryAkun
                        Me.TxtAdjustmentAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAdjustmentAccount.Text)
                        Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                        Me.TxtStockTakingAccount.Focus()
                    Else
                        Me.TxtAdjustmentAccount.Clear()
                        Me.TxtAdjustmentAccountDesc.Clear()
                        Me.TxtPrimaryAdjustmentAccount.Text = "-1"
                        Me.TxtAdjustmentAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtAdjustmentAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAdjustmentAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtAdjustmentAccount.Text = Me.buffer
                Me.btnFindAdjustmentAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtAdjustmentAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtAdjustmentAccount.Text) = True Then
                    Me.TxtPrimaryAdjustmentAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtAdjustmentAccount.Text)
                    Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                Else
                    If Me.TxtAdjustmentAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtAdjustmentAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAdjustmentAccount.Text = i.mPrimaryAkun
                        Me.TxtAdjustmentAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryAdjustmentAccount.Text)
                        Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                        Me.TxtStockTakingAccount.Focus()
                    Else
                        Me.TxtAdjustmentAccount.Clear()
                        Me.TxtAdjustmentAccountDesc.Clear()
                        Me.TxtPrimaryAdjustmentAccount.Text = "-1"
                        Me.TxtAdjustmentAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAdjustmentAccount.Text = "-1"
                Me.TxtAdjustmentAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindAdjustmentAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAdjustmentAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtAdjustmentAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtAdjustmentAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAdjustmentAccount.Text = i.mPrimaryAkun
                Me.TxtAdjustmentAccount.Text = obj.GetAccountCode(Me.TxtPrimaryAdjustmentAccount.Text)
                Me.TxtAdjustmentAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAdjustmentAccount.Text)
                Me.TxtStockTakingAccount.Focus()
            Else
                Me.TxtAdjustmentAccount.Clear()
                Me.TxtAdjustmentAccountDesc.Clear()
                Me.TxtAdjustmentAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtStockTakingAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTakingAccount.Enter
        Try
            Me.buffer = Me.TxtStockTakingAccount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTakingAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTakingAccount.GotFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtStockTakingAccount.ReadOnly = False Then
                    Me.TxtStockTakingAccount.BackColor = Me.BackColorAll
                Else
                    Me.TxtStockTakingAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTakingAccount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTakingAccount.LostFocus
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TxtStockTakingAccount.ReadOnly = False Then
                    Me.TxtStockTakingAccount.BackColor = Me.txtPrimaryProduct.BackColor
                Else
                    Me.TxtStockTakingAccount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockTakingAccount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockTakingAccount.KeyDown
        Dim i As New FrmMasterAkun_Cari1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan stock taking
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Me.btnFindStockTakingAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtStockTakingAccount.Text) = True Then
                    Me.TxtPrimaryStockTakingAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtStockTakingAccount.Text)
                    Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                    Me.btnOK.Focus()
                Else
                    If Me.TxtStockTakingAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockTakingAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryStockTakingAccount.Text = i.mPrimaryAkun
                        Me.TxtStockTakingAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryStockTakingAccount.Text)
                        Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtStockTakingAccount.Clear()
                        Me.TxtStockTakingAccountDesc.Clear()
                        Me.TxtPrimaryStockTakingAccount.Text = "-1"
                        Me.TxtStockTakingAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtStockTakingAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockTakingAccount.Leave
        Dim i As New FrmMasterAkun_Cari1

        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                Me.TxtStockTakingAccount.Text = Me.buffer
                Me.btnFindStockTakingAccount.Focus()
                Exit Sub
            End If

            i.BackColorAll = Me.BackColorAll
            If Me.TxtStockTakingAccount.Text <> "" Then
                If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtStockTakingAccount.Text) = True Then
                    Me.TxtPrimaryStockTakingAccount.Text = Me.ObjAkun.GetPrimary(Me.TxtStockTakingAccount.Text)
                    Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                Else
                    If Me.TxtStockTakingAccount.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockTakingAccount.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryStockTakingAccount.Text = i.mPrimaryAkun
                        Me.TxtStockTakingAccount.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimaryStockTakingAccount.Text)
                        Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.TxtStockTakingAccount.Clear()
                        Me.TxtStockTakingAccountDesc.Clear()
                        Me.TxtPrimaryStockTakingAccount.Text = "-1"
                        Me.TxtStockTakingAccount.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryStockTakingAccount.Text = "-1"
                Me.TxtStockTakingAccountDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindStockTakingAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindStockTakingAccount.Click
        Dim i As New FrmMasterAkun_Cari1
        Dim obj As New ClsMasterAkun1

        Try
            'cek apakah user ybs boleh edit akun2 terhubung dengan inventory
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterProduct_ChangeAccountCode") = False Then
                MsgBox("You don't have authorization to edit account data.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            If Me.TxtStockTakingAccount.Text <> "" Then
                i.txtCari.Text = Me.TxtStockTakingAccount.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtPrimaryStockTakingAccount.Text = i.mPrimaryAkun
                Me.TxtStockTakingAccount.Text = obj.GetAccountCode(Me.TxtPrimaryStockTakingAccount.Text)
                Me.TxtStockTakingAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryStockTakingAccount.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtStockTakingAccount.Clear()
                Me.TxtStockTakingAccountDesc.Clear()
                Me.TxtStockTakingAccount.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub txtUnit_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnit.TextChanged
        Try
            If Me.txtUnit.Text = "" Then
                Me.LblUnitBOM.Text = "Unit"
            Else
                Me.LblUnitBOM.Text = Me.txtUnit.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem_MultiUnit.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.txtProductCode.Enabled = True Then
                rowDGV = Me.dgvItem_MultiUnit.Rows(e.RowIndex)
                rowDGV.Cells(Me.ColUnit_Precision.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value = Me.ObjCD.DigitGrouping(1, Me.TxtDefaultPrecision.Text)
                rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value = Me.ObjCD.DigitGrouping(1, 0)
                rowDGV.Cells(Me.ColUnit_Ratio.Index).Value = 1
                rowDGV.Cells(Me.ColUnit_Desc.Index).Value = ""
                rowDGV.Cells(Me.ColUnit_Remarks.Index).Value = ""

                'add weight & volume
                rowDGV.Cells(Me.ColUnit_PrimaryUnitWeight.Index).Value = -1
                rowDGV.Cells(Me.ColUnit_NettWeight.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_GrossWeight.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_UnitWeight.Index).Value = ""
                rowDGV.Cells(Me.ColUnit_PrecisionWeight.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_PrimaryUnitVolume.Index).Value = -1
                rowDGV.Cells(Me.ColUnit_NettVolume.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_GrossVolume.Index).Value = 0
                rowDGV.Cells(Me.ColUnit_UnitVolume.Index).Value = ""
                rowDGV.Cells(Me.ColUnit_PrecisionVolume.Index).Value = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem_MultiUnit.RowsRemoved
        Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem_MultiUnit.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimaryUnit As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem_MultiUnit.Rows
                If rowDGV.Cells(Me.ColUnit_PrimaryUnit.Index).Value = vPrimaryUnit Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem_MultiUnit.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvItem_MultiUnit(Me.ColUnit_Unit.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColUnit_Unit.Index Then
                    If Me.objProduct.ApakahUnitSudahPernahDigunakan(Me.txtPrimaryProduct.Text, _
                    Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value, True) = True Then
                        'unit sudah digunakan
                        MsgBox("This unit is already used in transaction data. You cannot edit this product", _
                        MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    Else
                        'unit belum digunakan
                        Me.buffer = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                ElseIf e.ColumnIndex = Me.ColUnit_QtyDef.Index Or e.ColumnIndex = Me.ColUnit_QtyRasio.Index Or _
                e.ColumnIndex = Me.ColUnit_Remarks.Index Or e.ColumnIndex = Me.ColUnit_Precision.Index Then
                    Me.buffer = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value

                    'add weight & volume
                ElseIf e.ColumnIndex = Me.ColUnit_NettWeight.Index Or e.ColumnIndex = Me.ColUnit_GrossWeight.Index Or _
                e.ColumnIndex = Me.ColUnit_UnitWeight.Index Or e.ColumnIndex = Me.ColUnit_PrecisionWeight.Index Or _
                e.ColumnIndex = Me.ColUnit_NettVolume.Index Or e.ColumnIndex = Me.ColUnit_GrossVolume.Index Or _
                e.ColumnIndex = Me.ColUnit_UnitVolume.Index Or e.ColumnIndex = Me.ColUnit_PrecisionVolume.Index Then
                    If e.ColumnIndex = Me.ColUnit_UnitWeight.Index Then
                        If Me.objProduct.ApakahBolehEditUnitWeight(Me.txtPrimaryProduct.Text, Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value, Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnitWeight.Index, e.RowIndex).Value) = False Then
                            MsgBox("This unit weight is already used in transaction data. You cannot edit this product", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColUnit_UnitVolume.Index Then
                        If Me.objProduct.ApakahBolehEditUnitVolume(Me.txtPrimaryProduct.Text, Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value, Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnitVolume.Index, e.RowIndex).Value) = False Then
                            MsgBox("This unit volume is already used in transaction data. You cannot edit this product", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                            Exit Sub
                        End If
                    Else
                        Me.buffer = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                    '-------------------------------------------------------------------------
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColUnit_Unit.Index Then
                    Me.buffer = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem_MultiUnit.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem_MultiUnit.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Dim s As String = ""

        Try
            With Row
                If IsNothing(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColUnit_Remarks.Index Then
                        If ObjCD.CekDataKosong(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Remarks data connot contain the following characters: " _
                            + "' , % , #, or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    ''-----Cell kode unit-----
                    If e.ColumnIndex = Me.ColUnit_Unit.Index Then
                        Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUnit.ApakahSudahAda(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode unit yang diketik ada --> load data deskripsi + primary unit
                            Index = Me.GetIndexItemInput(Me.ObjUnit.GetPrimary(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value))
                            ''cek dengan unit default
                            s = Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value
                            If s = Me.txtUnit.Text Then
                                MsgBox("Unit code should not be the same as the default unit, please change your code", MsgBoxStyle.Information, "Information")
                                Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                Else
                                    Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
                                Exit Sub
                            End If
                            If Index < 0 Then
                                Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value = Me.ObjUnit.GetPrimary(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                            Else
                                ''cek data double
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication in unit code, please change your code", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode unit yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmCariMasterUnit
                            ij.BackColorAll = Me.BackColorAll
                            ij.txtCari.Text = Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.PrimaryUnit = Me.txtPrimaryUnit.Text Then
                                ''cek dengan unit default
                                MsgBox("Unit code should not be the same as the default unit, please change your code", MsgBoxStyle.Information, "Information")
                                Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.buffer) = True Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                Else
                                    Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
                                Exit Sub
                            End If
                            If ij.PrimaryUnit <> -1 Then
                                '2.a. kode unit yang dicari ada --> load data deskripsi + primary unit + kode unit
                                Index = Me.GetIndexItemInput(ij.PrimaryUnit)
                                If Index < 0 Then
                                    Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value = Me.ObjUnit.GetUnitCode(ij.PrimaryUnit)
                                    Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value = ij.PrimaryUnit
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication in unit code, please change your code", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
                                        Exit Sub
                                    Else
                                        Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value = Me.ObjUnit.GetUnitCode(ij.PrimaryUnit)
                                        Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value = ij.PrimaryUnit
                                    End If
                                End If
                            ElseIf ij.PrimaryUnit = -1 Then
                                If Me.dgvItem_MultiUnit.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
                        Me.IsikanRatio(sender, e)
                    End If

                    'Hapus Baris Jika Kode Unit Kosong
                    If e.ColumnIndex = Me.ColUnit_Remarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value) = True Then
                            If Me.dgvItem_MultiUnit.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        End If
                    End If
                    ''--------

                    'colom qty default
                    If e.ColumnIndex = Me.ColUnit_QtyDef.Index Then
                        'cek angka positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyDef.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'cek tidak boleh lebih dari 100rb
                        If Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyDef.Index, e.RowIndex).Value > 100000 Then
                            MsgBox("The maximum qty allowed is " + Me.ObjCD.DigitGrouping(100000, Me.TxtDefaultPrecision.Text), MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Try
                        End If
                        'isi nilai rasio
                        Me.IsikanRatio(sender, e)
                    End If

                    'colom qty rasio
                    If e.ColumnIndex = Me.ColUnit_QtyRasio.Index Then
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyRasio.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        If Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyRasio.Index, e.RowIndex).Value > 100000 Then
                            MsgBox("The maximum qty allowed is " + Me.ObjCD.DigitGrouping(100000, Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value), MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Me.IsikanRatio(sender, e)
                    End If

                    'kolom default precision
                    If e.ColumnIndex = Me.ColUnit_Precision.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        If Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value > 6 Then
                            MsgBox("The maximum value allowed is 6", MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If

                        Dim DecProd As Integer = Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value
                        If Me.objProduct.ApakahBolehEditDecimalProduct(Me.txtPrimaryProduct.Text, _
                        Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnit.Index, e.RowIndex).Value, _
                        DecProd, True) = False Then
                            MsgBox("Minimum value you must can enter is " & DecProd, MsgBoxStyle.Information, "Information")
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If

                        Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value, 0)
                        Me.IsikanRatio(sender, e)
                    End If

                    'add weight & volume
                    If e.ColumnIndex = Me.ColUnit_NettWeight.Index Or e.ColumnIndex = Me.ColUnit_GrossWeight.Index Or e.ColumnIndex = Me.ColUnit_UnitWeight.Index Or e.ColumnIndex = Me.ColUnit_PrecisionWeight.Index Then
                        Me.EndEditWeight(Row, e.ColumnIndex)
                    ElseIf e.ColumnIndex = Me.ColUnit_NettVolume.Index Or e.ColumnIndex = Me.ColUnit_GrossVolume.Index Or e.ColumnIndex = Me.ColUnit_UnitVolume.Index Or e.ColumnIndex = Me.ColUnit_PrecisionVolume.Index Then
                        Me.EndEditVolume(Row, e.ColumnIndex)
                    End If
                Else
                    If e.ColumnIndex = Me.ColUnit_Remarks.Index Then
                        Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColUnit_Unit.Index Then
                            If String.IsNullOrEmpty(Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvItem_MultiUnit.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                            'add weight & volume
                        ElseIf e.ColumnIndex = Me.ColUnit_UnitWeight.Index Then
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnitWeight.Index, e.RowIndex).Value = -1
                        ElseIf e.ColumnIndex = Me.ColUnit_UnitVolume.Index Then
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvItem_MultiUnit.Item(Me.ColUnit_PrimaryUnitVolume.Index, e.RowIndex).Value = -1
                            '----------------------------------------------------------------
                        Else
                            Me.dgvItem_MultiUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'add weight & volume
    Sub EndEditWeight(ByVal rDgv As DataGridViewRow, ByVal vColIdx As Integer)
        Try
            If IsNothing(rDgv.Cells(vColIdx).Value) = False Then
                If vColIdx = Me.ColUnit_NettWeight.Index Or vColIdx = Me.ColUnit_GrossWeight.Index Or vColIdx = Me.ColUnit_PrecisionWeight.Index Then
                    If Me.ObjCD.CekNonNegatif(rDgv.Cells(vColIdx).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rDgv.Cells(vColIdx).Value = Me.buffer
                        Exit Sub
                    End If

                    If vColIdx = Me.ColUnit_NettWeight.Index Or vColIdx = Me.ColUnit_GrossWeight.Index Then
                        rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(CDec(rDgv.Cells(vColIdx).Value), CInt(rDgv.Cells(Me.ColUnit_PrecisionWeight.Index).Value))
                    ElseIf vColIdx = Me.ColUnit_PrecisionWeight.Index Then
                        'cek batasan ke transaksi
                        rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(CDec(rDgv.Cells(vColIdx).Value), 0)
                        If Me.objProduct.ApakahBolehEditDecimalWeight(Me.txtPrimaryProduct.Text, rDgv.Cells(Me.ColUnit_PrimaryUnit.Index).Value, rDgv.Cells(Me.ColUnit_PrimaryUnitWeight.Index).Value, rDgv.Cells(Me.ColUnit_PrecisionWeight.Index).Value) = False Then
                            rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(Me.buffer, 0)
                            Exit Sub
                        End If
                        rDgv.Cells(Me.ColUnit_NettWeight.Index).Value = Me.ObjCD.DigitGrouping(rDgv.Cells(Me.ColUnit_NettWeight.Index).Value, CInt(rDgv.Cells(vColIdx).Value))
                        rDgv.Cells(Me.ColUnit_GrossWeight.Index).Value = Me.ObjCD.DigitGrouping(rDgv.Cells(Me.ColUnit_GrossWeight.Index).Value, CInt(rDgv.Cells(vColIdx).Value))
                    End If
                ElseIf vColIdx = Me.ColUnit_UnitWeight.Index Then
                    rDgv.Cells(vColIdx).Value = rDgv.Cells(vColIdx).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUnit.ApakahSudahAda(rDgv.Cells(vColIdx).Value) = True Then
                        rDgv.Cells(Me.ColUnit_PrimaryUnitWeight.Index).Value = Me.ObjUnit.GetPrimary(rDgv.Cells(vColIdx).Value)
                    Else
                        Dim ij As New FrmCariMasterUnit
                        ij.BackColorAll = Me.BackColorAll
                        ij.txtCari.Text = rDgv.Cells(vColIdx).Value
                        ij.ShowDialog()
                        If ij.PrimaryUnit <> -1 Then
                            rDgv.Cells(vColIdx).Value = Me.ObjUnit.GetUnitCode(ij.PrimaryUnit)
                            rDgv.Cells(Me.ColUnit_PrimaryUnitWeight.Index).Value = ij.PrimaryUnit
                        ElseIf ij.PrimaryUnit = -1 Then
                            rDgv.Cells(vColIdx).Value = ""
                            rDgv.Cells(Me.ColUnit_PrimaryUnitWeight.Index).Value = -1
                        End If
                        ij = Nothing
                    End If
                End If
            Else
                rDgv.Cells(vColIdx).Value = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'add weight & volume
    Sub EndEditVolume(ByVal rDgv As DataGridViewRow, ByVal vColIdx As Integer)
        Try
            If IsNothing(rDgv.Cells(vColIdx).Value) = False Then
                If vColIdx = Me.ColUnit_NettVolume.Index Or vColIdx = Me.ColUnit_GrossVolume.Index Or vColIdx = Me.ColUnit_PrecisionVolume.Index Then
                    If Me.ObjCD.CekNonNegatif(rDgv.Cells(vColIdx).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rDgv.Cells(vColIdx).Value = Me.buffer
                        Exit Sub
                    End If

                    If vColIdx = Me.ColUnit_NettVolume.Index Or vColIdx = Me.ColUnit_GrossVolume.Index Then
                        rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(CDec(rDgv.Cells(vColIdx).Value), CInt(rDgv.Cells(Me.ColUnit_PrecisionVolume.Index).Value))
                    ElseIf vColIdx = Me.ColUnit_PrecisionVolume.Index Then
                        'cek batasan ke transaksi
                        rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(CDec(rDgv.Cells(vColIdx).Value), 0)
                        If Me.objProduct.ApakahBolehEditDecimalVolume(Me.txtPrimaryProduct.Text, rDgv.Cells(Me.ColUnit_PrimaryUnit.Index).Value, rDgv.Cells(Me.ColUnit_PrimaryUnitVolume.Index).Value, rDgv.Cells(Me.ColUnit_PrecisionVolume.Index).Value) = False Then
                            rDgv.Cells(vColIdx).Value = Me.ObjCD.DigitGrouping(Me.buffer, 0)
                            Exit Sub
                        End If
                        rDgv.Cells(Me.ColUnit_NettVolume.Index).Value = Me.ObjCD.DigitGrouping(rDgv.Cells(Me.ColUnit_NettVolume.Index).Value, CInt(rDgv.Cells(vColIdx).Value))
                        rDgv.Cells(Me.ColUnit_GrossVolume.Index).Value = Me.ObjCD.DigitGrouping(rDgv.Cells(Me.ColUnit_GrossVolume.Index).Value, CInt(rDgv.Cells(vColIdx).Value))
                    End If
                ElseIf vColIdx = Me.ColUnit_UnitVolume.Index Then
                    rDgv.Cells(vColIdx).Value = rDgv.Cells(vColIdx).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjUnit.ApakahSudahAda(rDgv.Cells(vColIdx).Value) = True Then
                        rDgv.Cells(Me.ColUnit_PrimaryUnitVolume.Index).Value = Me.ObjUnit.GetPrimary(rDgv.Cells(vColIdx).Value)
                    Else
                        Dim ij As New FrmCariMasterUnit
                        ij.BackColorAll = Me.BackColorAll
                        ij.txtCari.Text = rDgv.Cells(vColIdx).Value
                        ij.ShowDialog()
                        If ij.PrimaryUnit <> -1 Then
                            rDgv.Cells(vColIdx).Value = Me.ObjUnit.GetUnitCode(ij.PrimaryUnit)
                            rDgv.Cells(Me.ColUnit_PrimaryUnitVolume.Index).Value = ij.PrimaryUnit
                        ElseIf ij.PrimaryUnit = -1 Then
                            rDgv.Cells(vColIdx).Value = ""
                            rDgv.Cells(Me.ColUnit_PrimaryUnitVolume.Index).Value = -1
                        End If
                        ij = Nothing
                    End If
                End If
            Else
                rDgv.Cells(vColIdx).Value = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanRatio(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim Qty As Decimal = 0
        Dim QtyRasio As Decimal = 0
        Dim Rasio As Decimal = 0
        Dim str As String = ""

        Try
            Qty = ObjCD.DigitGrouping(Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyDef.Index, e.RowIndex).Value, CInt(Me.TxtDefaultPrecision.Text))
            QtyRasio = ObjCD.DigitGrouping(Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyRasio.Index, e.RowIndex).Value, CInt(Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value))
            If QtyRasio <> 0 Then
                Rasio = ObjCD.DigitGrouping(Qty / QtyRasio, 6)
            Else
                Rasio = 0
            End If
            str = Me.ObjCD.BuangNolDibelakangDecimal(Rasio.ToString)

            If String.IsNullOrEmpty(str) = True Then
                Me.dgvItem_MultiUnit.Item(Me.ColUnit_Ratio.Index, e.RowIndex).Value = 0
            Else
                Me.dgvItem_MultiUnit.Item(Me.ColUnit_Ratio.Index, e.RowIndex).Value = ObjCD.DigitGrouping(str)
            End If
            Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyDef.Index, e.RowIndex).Value = ObjCD.DigitGrouping(Qty, CInt(Me.TxtDefaultPrecision.Text))
            Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyRasio.Index, e.RowIndex).Value = ObjCD.DigitGrouping(QtyRasio, CInt(Me.dgvItem_MultiUnit.Item(Me.ColUnit_Precision.Index, e.RowIndex).Value))
            Me.dgvItem_MultiUnit.Item(Me.ColUnit_Desc.Index, e.RowIndex).Value = _
            Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyRasio.Index, e.RowIndex).Value.ToString + " " + Me.dgvItem_MultiUnit.Item(Me.ColUnit_Unit.Index, e.RowIndex).Value + _
            " = " + Me.dgvItem_MultiUnit.Item(Me.ColUnit_QtyDef.Index, e.RowIndex).Value.ToString + " " + Me.txtUnit.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvItem_MultiUnit.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.dgvItem_MultiUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekQtyKosong() As Boolean
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem_MultiUnit.Rows
                If rowDGV.Cells(Me.ColUnit_Unit.Index).Value <> "" Then
                    If rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value = 0 Then
                        MsgBox("Please specify a non-negative number for default qty.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem_MultiUnit.Focus()
                        rowDGV.Cells(Me.ColUnit_QtyDef.Index).Selected = True
                        Return True
                        Exit Function
                    End If
                    If rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value = 0 Then
                        MsgBox("Please specify a non-negative number for ratio qty.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem_MultiUnit.Focus()
                        rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Selected = True
                        Return True
                        Exit Function
                    End If
                Else
                    Return False
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub UbahFormatAngkaDGV()
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem_MultiUnit.Rows
                If rowDGV.Cells(Me.ColUnit_Unit.Index).Value <> "" Then
                    If rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value <> 0 Then
                        rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value = FormatNumber(rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value, CInt(Me.TxtDefaultPrecision.Text))
                    End If
                    If rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value <> 0 Then
                        rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value = FormatNumber(rowDGV.Cells(Me.ColUnit_QtyRasio.Index).Value, CInt(rowDGV.Cells(Me.ColUnit_Precision.Index).Value))
                    End If
                    If rowDGV.Cells(Me.ColUnit_Ratio.Index).Value <> 0 Then
                        rowDGV.Cells(Me.ColUnit_Ratio.Index).Value = ObjCD.BuangNolDibelakangDecimal(rowDGV.Cells(Me.ColUnit_Ratio.Index).Value)
                    End If
                    rowDGV.Cells(Me.ColUnit_Desc.Index).Value = _
                    rowDGV.Cells(Me.ColUnit_QtyDef.Index).Value + " " + Me.txtUnit.Text + " = " + _
                    rowDGV.Cells(Me.ColUnit_Ratio.Index).Value + " " + rowDGV.Cells(Me.ColUnit_Unit.Index).Value
                Else
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbEditPriceList_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbEditPriceList.LinkClicked
        Dim i As New frmMasterProduct_PriceList

        Try
            If Me.txtProductCode.Enabled = False Then
                i.ShowLangsung = True
                If Me.txtPrimaryProduct.Text = "" Then
                    i.PrimaryProduct = -1
                Else
                    i.PrimaryProduct = Me.txtPrimaryProduct.Text
                End If
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub DGVKompetensi_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVKompetensi.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DGVKompetensi(Me.ColCompetence_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColCompetence_Value.Index Or e.ColumnIndex = Me.ColCompetence_NextCertificationDate.Index Or _
                e.ColumnIndex = Me.ColCompetence_Code.Index Or e.ColumnIndex = Me.ColCompetence_Remarks.Index Then
                    Me.buffer = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCompetence_Code.Index Then
                    Me.buffer = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVKompetensi.CellEndEdit
        Dim row As DataGridViewRow = Me.DGVKompetensi.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double
        Dim PrimaryItem As Double
        Dim DefPrec As Integer = 0
        Dim DataType As Integer
        Try
            With row
                '===code
                If e.ColumnIndex = Me.ColCompetence_Code.Index Then
                    If String.IsNullOrEmpty(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjKompetensi.ApakahSudahAda(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Primary = Me.ObjKompetensi.GetPrimary(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemInput2(Primary)
                            If index < 0 Then
                                DefPrec = Me.ObjKompetensi.GetDefPrec(Primary)
                                DataType = Me.ObjKompetensi.GetDataType(Primary)
                                Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value = Primary
                                Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(Primary)
                                Me.DGVKompetensi.Item(Me.ColCempetence_Description.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(Primary)
                                'If Me.ObjPost.GetJmlItem(Me.TxtPrimaryPosition.Text, Primary) > 0 Then
                                '    Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                '    Me.ObjPost.GetValue_CustomFormatted(Me.TxtPrimaryPosition.Text, Primary, Me.ObjCekData, DataType)
                                'Else
                                Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                Me.ObjKompetensi.GetValue_CustomFormatted(Primary, Me.ObjCekData, DataType)
                                'End If
                                Me.DGVKompetensi.Item(Me.ColCempetence_DataType.Index, e.RowIndex).Value = DataType
                                'If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, Primary) = 1 Then
                                '    Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = Format(Now, "dd MMM yyyy")
                                'Else
                                Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = "-"
                                'Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = Format _
                                '(CDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                                'End If
                                Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value = ""
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.mDgv = Me.DGVKompetensi
                                        Me.mEraseRowIndex = row.Index
                                        Me.BgwDeleteRowDgv.RunWorkerAsync()
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New FrmMasterKompetensi_Product_Cari
                            ij.txtCari.Text = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexItemInput2(ij.mPrimary)
                            If ij.mPrimary = -1 Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.mDgv = Me.DGVKompetensi
                                Me.mEraseRowIndex = row.Index
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                            Else
                                If index < 0 Then
                                    DefPrec = Me.ObjKompetensi.GetDefPrec(ij.mPrimary)
                                    DataType = Me.ObjKompetensi.GetDataType(ij.mPrimary)
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                                    Me.DGVKompetensi.Item(Me.ColCempetence_Description.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(ij.mPrimary)
                                    'If Me.ObjPost.GetJmlItem(Me.TxtPrimaryPosition.Text, ij.mPrimary) > 0 Then
                                    '    Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                    '    Me.ObjPost.GetValue_CustomFormatted(Me.TxtPrimaryPosition.Text, ij.mPrimary, Me.ObjCekData, DataType)
                                    'Else
                                    Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                    Me.ObjKompetensi.GetValue_CustomFormatted(ij.mPrimary, Me.ObjCekData, DataType)
                                    'End If
                                    Me.DGVKompetensi.Item(Me.ColCempetence_DataType.Index, e.RowIndex).Value = DataType
                                    'If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, ij.mPrimary) = 1 Then
                                    '    Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = Format(Now, "dd MMM yyyy")
                                    'Else
                                    Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = "-"
                                    'End If
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value = ""
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.mDgv = Me.DGVKompetensi
                                            Me.mEraseRowIndex = row.Index
                                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        Me.mDgv = Me.DGVKompetensi
                        Me.mEraseRowIndex = row.Index
                        Me.BgwDeleteRowDgv.RunWorkerAsync()
                    End If
                End If
                '===value
                If e.ColumnIndex = Me.ColCompetence_Value.Index Then
                    If String.IsNullOrEmpty(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        DefPrec = Me.ObjKompetensi.GetDefPrec(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value)
                        DataType = Me.ObjKompetensi.GetDataType(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value)
                        If DataType = ClsMasterKompetensi_Product.EnumDataType.Alfabetic Then
                            If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "A" Or _
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "B" Or _
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "C" Or _
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "D" Or _
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "E" Or _
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "F" Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            Else
                                MsgBox("Must input A ,B ,C ,D ,E ,F", MsgBoxStyle.Information, "Information")
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        ElseIf DataType = ClsMasterKompetensi_Product.EnumDataType.Number Then
                            If Me.ObjCekData.CekNumerik(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekData)
                            End If
                        ElseIf DataType = ClsMasterKompetensi_Product.EnumDataType.Tanggal Then
                            If IsDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                ''set region ke indonesia
                                'My.Application.ChangeCulture("id-ID")
                                'System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("id-ID", True)
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekData)
                                ''kembalikan ke region defaultnya
                                'My.Application.ChangeCulture("en-US")
                                'System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("en-US", True)
                            End If
                        ElseIf DataType = ClsMasterKompetensi_Product.EnumDataType.Text Then
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekData)
                        ElseIf DataType = ClsMasterKompetensi_Product.EnumDataType.YN Then
                            If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "YES" Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                            ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "NO" Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                            ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "Y" Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                            ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "N" Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                            Else
                                MsgBox("Must input Yes / No .", MsgBoxStyle.Information, "Information")
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            ''Aul 10 nov 2015
                        ElseIf DataType = ClsMasterKompetensi_Product.EnumDataType.DropdownList Then
                            If Me.ObjKompetensi.ApakahSudahAdaItem(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value, _
                            Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value) = False Then
                                'kode akun yang diketik tidak ada --> tampilkan form cari
                                Dim ij As New Core_Module.FrmMasterKompetensi_Product_Item_Cari
                                ij.mPrimaryCompetenceItem = Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value
                                ij.txtCari.Text = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.BackColorAll = Me.BackColorAll
                                ij.ShowDialog()
                                If ij.mPrimaryItem <> -1 Then
                                    'kode akun yang dicari ada --> load data 
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCodeItem(ij.mPrimaryItem)
                                    Me.DGVKompetensi.Item(Me.ColCempetence_Identifier.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataIdentifier(ij.mPrimaryItem)
                                ElseIf ij.mPrimaryItem = -1 Then
                                    If Me.DGVKompetensi.Rows.Count - 1 <> e.RowIndex Then
                                        Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    End If
                                End If
                                ij = Nothing
                            Else
                                PrimaryItem = Me.ObjKompetensi.GetPrimaryItem(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value, _
                                              Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value)
                                Me.DGVKompetensi.Item(Me.ColCempetence_Identifier.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataIdentifier(PrimaryItem)
                            End If
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                            ''---------  
                        End If
                    Else
                        Me.DGVKompetensi.Item(Me.ColCempetence_Identifier.Index, e.RowIndex).Value = ""
                    End If
                End If
                ''sertifikasi
                If e.ColumnIndex = Me.ColCompetence_NextCertificationDate.Index Then
                    If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                    Else
                        ''set region ke indonesia
                        'My.Application.ChangeCulture("id-ID")
                        'System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("id-ID", True)
                        If IsDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        ''kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("en-US", True)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVKompetensi.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DGVKompetensi.RowsRemoved
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVKompetensi.Sorted
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput2(ByVal vPrimary As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DGVKompetensi.Rows
                If rowDGV.Cells(Me.ColCompetence_Primary.Index).Value = vPrimary Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub BgwDeleteRowDgv_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDeleteRowDgv.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDeleteRowDgv_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDeleteRowDgv.RunWorkerCompleted
        Try
            Me.mDgv.Rows.RemoveAt(Me.mEraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVAttanchment.CellBeginEdit
        Try
            'jika baris baru, maka yang bisa diedit hanya desc-nya saja
            If Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAtt_Description.Index Then
                    Me.buffer = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'jik bukan baris baru, yang bisa diedit adalah:
                'desc, remarks, file path, 
                If e.ColumnIndex = Me.ColAtt_Description.Index Or _
                e.ColumnIndex = Me.ColAtt_LocationFile.Index Or _
                e.ColumnIndex = Me.ColAtt_Remarks.Index Then
                    Me.buffer = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemAttachment(ByVal vDesc As String, ByVal vIndexAcuan As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DGVAttanchment.Rows
                If rowDGV.Index = vIndexAcuan Or rowDGV.Cells(Me.ColAtt_Description.Index).Value = "" Then
                Else
                    If rowDGV.Cells(Me.ColAtt_Description.Index).Value = vDesc Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DGVAttanchment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchment.CellEndEdit
        Dim index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.DGVAttanchment.Rows(e.RowIndex)

            With row
                If IsNothing(Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim

                    If e.ColumnIndex = Me.ColAtt_Description.Index Then
                        'pastikan ada isinya
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                            MsgBox("Please specify a valid description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
                            Else
                                Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If
                        'upper case
                        Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'pastikan tidak ad duplikasi
                        index = GetIndex_ItemAttachment(Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value, e.RowIndex)
                        If index >= 0 Then
                            MsgBox("Duplication in description data, please input another description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
                            Else
                                Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_Remarks.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_LocationFile.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        'jika bukan URL, maka pastikan filenya exist
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then
                            If Me.ObjCD.ApakahUrlValid(Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                If My.Computer.FileSystem.FileExists((Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value)) = False Then
                                    MsgBox("The file path you specified is not a valid url path nor a valid file path.", MsgBoxStyle.Information, "Information")
                                    Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.buffer = "" Then
                                        Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                                        Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                                    End If
                                    Exit Sub
                                Else
                                    Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                                End If
                            Else
                                Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                            End If
                        Else
                            Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                            Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColAtt_Description.Index Then
                        If Me.buffer <> "" Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.EraseRowIndex = row.Index
                            Me.BackgroundWorker3.RunWorkerAsync()
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_LocationFile.Index Or _
                    e.ColumnIndex = Me.ColAtt_Remarks.Index Then
                        Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVAttanchment.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.btnOK.Enabled = True Then
                rowDGV = Me.DGVAttanchment.Rows(Me.DGVAttanchment.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColAtt_BtnOpenDelete.Index).Value = "add"
                    .Cells(Me.ColAtt_Description.Index).Value = ""
                    .Cells(Me.ColAtt_LocationFile.Index).Value = ""
                    .Cells(Me.ColAtt_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColAtt_Preview.Index).Value = "-"
                    .Cells(Me.ColAtt_Remarks.Index).Value = ""
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DGVAttanchment.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVAttanchment.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchment.CellClick
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value <> "" Then
                If e.ColumnIndex = Me.ColAtt_Preview.Index Then
                    If Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value <> "" Then
                        If Me.ObjCD.ApakahUrlValid(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value) = True Then
                            'jika valid url maka tampilkan ke browser
                            Process.Start(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                        Else
                            'jika file path, maka buka filenya
                            Dim fileExists As Boolean
                            fileExists = My.Computer.FileSystem.FileExists(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                            If fileExists = False Then
                                MsgBox("You are trying to open an inaccessible file or the file does not exist anymore.", MsgBoxStyle.Information, "Information")
                            Else
                                Process.Start(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColAtt_BtnOpenDelete.Index Then
                    If Me.btnOK.Enabled = True Then
                        If Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add" Then
                            'add file
                            If Me.OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
                                If Me.OpenFileDialog2.FileName <> "" Then
                                    Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value = Me.OpenFileDialog2.FileName
                                    Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                                End If
                            End If
                        Else
                            'delete file
                            If MsgBox("Remove file path?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value = ""
                                Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                                Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVAttanchment.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemWorkUnit(ByVal vPrimaryWorkUnit As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvWorkUnit.Rows.Count - 2
                RowDGV = Me.DgvWorkUnit.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColWorkUnit_PrimaryWorkUnit.Index).Value = vPrimaryWorkUnit Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvWorkUnit_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWorkUnit.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value) = False Then
                If (e.ColumnIndex = Me.ColWorkUnit_Code.Index Or e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index _
                Or e.ColumnIndex = Me.ColWorkUnit_Remarks.Index) Or e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Efficiency.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                    '''set header fungsi digunakan ketika edit code 
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_No.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Efficiency.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Remarks.Index).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.buffer = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWorkUnit.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvWorkUnit.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjWorkUnit.ApakahSudahAdaOnlyActive(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Primary = Me.ObjWorkUnit.GetPrimary(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemWorkUnit(Primary, e.RowIndex)
                            If index < 0 Then
                                Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetPrimary(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value)
                                Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterWorkUnit_Cari
                            ij.txtCari.Text = Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.Status = "Only Active"
                            ij.ShowDialog()
                            If ij.PrimaryWorkUnit <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                index = Me.GetIndexItemWorkUnit(ij.PrimaryWorkUnit, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                    Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                        Else
                                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                                        Exit Sub
                                    Else
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_Code.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetCode(ij.PrimaryWorkUnit)
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_Desc.Index, e.RowIndex).Value = Me.ObjWorkUnit.GetDesc(ij.PrimaryWorkUnit)
                                        Me.DgvWorkUnit.Item(Me.ColWorkUnit_PrimaryWorkUnit.Index, e.RowIndex).Value = ij.PrimaryWorkUnit
                                    End If
                                End If
                            ElseIf ij.PrimaryWorkUnit = -1 Then
                                If Me.DgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                                    Else
                                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    ElseIf e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index Then
                        'digit grouping
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        'maximal nilainya 100 %
                        If CDec(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                    ElseIf e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                        'digit grouping
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value, Me.TxtDefaultPrecision.Text)
                        'harus bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColWorkUnit_Efficiency.Index Or e.ColumnIndex = Me.ColWorkUnit_Rate.Index Then
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    ElseIf e.ColumnIndex = Me.ColWorkUnit_Code.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            If Me.DgvWorkUnit.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex = e.RowIndex
                                Me.BgwDelDgvWorkUnit.RunWorkerAsync()
                            End If
                        Else
                            Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvWorkUnit.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_No.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Code.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Desc.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Efficiency.Index).SortMode = DataGridViewColumnSortMode.Automatic
            'Me.DgvWorkUnit.Columns(Me.ColWorkUnit_Remarks.Index).SortMode = DataGridViewColumnSortMode.Automatic
        End Try
    End Sub

    Private Sub DgvWorkUnit_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvWorkUnit.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.btnOK.Enabled = True And Me.DgvWorkUnit.AllowUserToAddRows = True Then
                rowDGV = Me.DgvWorkUnit.Rows(Me.DgvWorkUnit.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColWorkUnit_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColWorkUnit_Code.Index).Value = ""
                    .Cells(Me.ColWorkUnit_Efficiency.Index).Value = "100.00"
                    .Cells(Me.ColWorkUnit_Rate.Index).Value = "1"
                    .Cells(Me.ColWorkUnit_Remarks.Index).Value = ""
                    .Cells(Me.ColWorkUnit_PrimaryWorkUnit.Index).Value = "-1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvWorkUnit.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWorkUnit_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvWorkUnit.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelDgvWorkUnit_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDelDgvWorkUnit.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDelDgvWorkUnit_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDelDgvWorkUnit.RunWorkerCompleted
        Try
            Me.DgvWorkUnit.Rows.RemoveAt(Me.EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvWorkUnit)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCopy_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCopy.LinkClicked
        Dim hj As New frmCariMasterProduct
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryProduct = -1 Then
                Me.KosongkanField()

                If Me.objProduct.CariDetail(Me.DGVKompetensi, hj.PrimaryProduct, Me.dgvPricing, Me.pbxProduct, Me.dgvItemBOM, _
                Me, Me.dgvItem_MultiUnit, Me.DGVAttanchment, Me.DgvWorkUnit) = True Then
                    Me.SalinKeField()

                    Me.baru = True
                    Me.AktifkanDetail()
                    'Me.KosongkanField()
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCategoryCode.Focus()
                    Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
                    Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
                    Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

                    'cari apakah kode terakhir yang disimpan menggunakan barcode
                    If Me.objProduct.ApakahKodeTerakhirMenggunakanBarcode = True Then
                        Me.chkBarcode.Checked = False
                        Me.chkBarcode.Checked = True
                    Else
                        Me.chkBarcode.Checked = True
                        Me.chkBarcode.Checked = False
                    End If

                    'secara default product
                    'Me.rboGoods.Checked = True

                    'barang baru, tidak ada verifikasi akun product
                    Me.TxtIsVerified.Text = 0
                    Me.TxtVerifiedBy.Text = ""
                    Me.TxtVerifiedTime.Text = Format(#1/1/1950#, "dd MMM yyyy HH:mm:ss")
                    Me.TxtAccountVerificationStatus.Text = "Not verified"

                    'hapuskan data txt category, code, spec dan desc
                    'Me.txtPrimaryProduct.Text = "-1"
                    'Me.txtPrimaryCategory.Text = "-1"
                    'Me.txtCategoryCode.Text = ""
                    Me.txtProductCode.Text = "COPY OF " + Me.txtProductCode.Text
                    If Me.txtSpecification.Text <> "" Then
                        Me.txtSpecification.Text = "COPY OF " + Me.txtSpecification.Text
                    End If
                    Me.txtDescription.Text = "COPY OF " + Me.txtDescription.Text
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

    Private Sub dgvItemBOM_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItemBOM.SelectionChanged
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.dgvItemBOM.SelectedRows.Count > 0 Then
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

    Private Sub DgvWorkUnit_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DgvWorkUnit.SelectionChanged
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.DgvWorkUnit.SelectedRows.Count > 0 Then
                    Me.btnUp1.Enabled = True
                    Me.btnDown1.Enabled = True
                Else
                    Me.btnUP1.Enabled = False
                    Me.btnDown1.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.dgvItemBOM.AllowUserToAddRows = False
            If Me.dgvItemBOM.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemBOM.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemBOM.Rows.Remove(row)
                    Me.dgvItemBOM.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemBOM.Focus()
                    Me.dgvItemBOM.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItemBOM)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True, Me.dgvItemBOM)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
                Me.Update_TreeView()
            End If
            Me.dgvItemBOM.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean, ByVal vdgv As DataGridView)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < vdgv.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    vdgv.FirstDisplayedScrollingRowIndex = vdgv.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > vdgv.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While vdgv.Rows(row.Index).Displayed = False
                    vdgv.FirstDisplayedScrollingRowIndex = vdgv.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If vdgv.DisplayedRowCount(True) <> vdgv.DisplayedRowCount(False) Then
                    vdgv.FirstDisplayedScrollingRowIndex = vdgv.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvItemBOM.AllowUserToAddRows = False
            If Me.dgvItemBOM.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItemBOM.SelectedRows(0)
                If row.Index <> Me.dgvItemBOM.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemBOM.Rows.Remove(row)
                    Me.dgvItemBOM.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemBOM.Focus()
                    Me.dgvItemBOM.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItemBOM)
                    Me.ShowRowProcess(row, False, Me.dgvItemBOM)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
                Me.Update_TreeView()
            End If
            Me.dgvItemBOM.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'penambahan untuk perbaikan master product
    Private Sub dgvItem_MultiUnit_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem_MultiUnit.UserDeletingRow

        Try
            If Me.objProduct.ApakahUnitSudahPernahDigunakan(Me.txtPrimaryProduct.Text, _
            Me.dgvItem_MultiUnit(Me.ColUnit_PrimaryUnit.Index, e.Row.Index).Value, True) = True Then
                'unit sudah digunakan
                MsgBox("This unit is already used in transaction data. You cannot delete this product", _
                MsgBoxStyle.Information, "Information")
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp1.Click
        Try
            Me.DgvWorkUnit.AllowUserToAddRows = False
            If Me.DgvWorkUnit.SelectedRows.Count > 0 Then
                Me.btnUp1.Enabled = False
                Me.btnDown1.Enabled = False
                Dim row As DataGridViewRow = Me.DgvWorkUnit.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DgvWorkUnit.Rows.Remove(row)
                    Me.DgvWorkUnit.Rows.InsertRange(indexbaru, row)
                    Me.DgvWorkUnit.Focus()
                    Me.DgvWorkUnit.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True, Me.DgvWorkUnit)
                    End If
                End If

                Me.btnUp1.Enabled = True
                Me.btnDown1.Enabled = True
            End If
            Me.DgvWorkUnit.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown1.Click
        Try
            Me.DgvWorkUnit.AllowUserToAddRows = False
            If Me.DgvWorkUnit.SelectedRows.Count > 0 Then
                Me.btnUp1.Enabled = False
                Me.btnDown1.Enabled = False

                Dim row As DataGridViewRow = Me.DgvWorkUnit.SelectedRows(0)
                If row.Index <> Me.DgvWorkUnit.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DgvWorkUnit.Rows.Remove(row)
                    Me.DgvWorkUnit.Rows.InsertRange(indexbaru, row)
                    Me.DgvWorkUnit.Focus()
                    Me.DgvWorkUnit.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvWorkUnit)
                    Me.ShowRowProcess(row, False, Me.DgvWorkUnit)
                End If

                Me.btnUp1.Enabled = True
                Me.btnDown1.Enabled = True
            End If
            Me.DgvWorkUnit.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub cboDependency_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDependency.SelectedIndexChanged
    '    Try
    '        If Me.cboDependency.SelectedIndex = clsMasterProduct.EnumTipeDependency.End_Start1 Then
    '            Me.txtDependency.Text = Me.ObjCD.Convert_ToString(0)
    '            Me.txtDependency.ReadOnly = True
    '            Me.txtDependency.BackColor = Me.TxtReadOnly.BackColor
    '        Else
    '            If Me.txtDependency.Enabled = True Then
    '            Else
    '                Me.txtDependency.BackColor = Me.TxtReadOnly.BackColor
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub txtDependency_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDependency.Enter
        Try
            Me.buffer = Me.txtDependency.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.GotFocus
        Try
            If Me.txtDependency.Enabled = True Then
                If Me.txtDependency.ReadOnly = True Then
                    Me.txtDependency.BackColor = Me.txtDependency.BackColor
                Else
                    Me.txtDependency.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.Leave
        Try
            If Me.ObjCD.Is_Time_String(Me.txtDependency.Text) = False Then
                MsgBox("Please input valid format (HH.mm).", MsgBoxStyle.Information, "Information")
                Me.txtDependency.Text = Me.buffer
                Me.txtDependency.Focus()
                Exit Sub
            End If
            If Me.ObjCD.Convert_ToMinutes(Me.txtDependency.Text) < 0 Then
                MsgBox("Delay dependency duration cannot negative value.", MsgBoxStyle.Information, "Information")
                Me.txtDependency.Text = Me.buffer
                Me.txtDependency.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDependency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDependency.LostFocus
        Try
            If Me.txtDependency.Enabled = True Then
                If Me.txtDependency.ReadOnly = True Then
                    Me.txtDependency.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtDependency.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTolerancePerc(ByVal vGoodPercentage As Decimal) As Decimal
        Try
            If vGoodPercentage > 0 Then
                Return ((100 - vGoodPercentage) / vGoodPercentage) * 100
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetGoodPercentagePerc(ByVal vTolerance As Decimal) As Decimal
        Try
            Return (100 / (100 + vTolerance)) * 100
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'add weight & volume
    Private Sub TxtNettWeight_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettWeight.Enter
        Try
            Me.buffer = Me.TxtNettWeight.Text
            Me.ObjCD.Number_Enter(Me.TxtNettWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettWeight.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtNettWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettWeight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettWeight.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtNettWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettWeight_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettWeight.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtNettWeight.Text) = False Then
                MsgBox("Nett weight be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtNettWeight.Focus()
                Exit Sub
            End If

            Me.ObjCD.Number_Leave(Me.TxtNettWeight, sender, e, CInt(Me.TxtDefaultPrecisionWeight.Text), Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossWeight_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGrossWeight.Enter
        Try
            Me.buffer = Me.TxtGrossWeight.Text
            Me.ObjCD.Number_Enter(Me.TxtGrossWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGrossWeight.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtGrossWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossWeight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGrossWeight.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtGrossWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossWeight_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGrossWeight.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtGrossWeight.Text) = False Then
                MsgBox("Gross weight be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtGrossWeight.Focus()
                Exit Sub
            End If
            Me.ObjCD.Number_Leave(Me.TxtGrossWeight, sender, e, CInt(Me.TxtDefaultPrecisionWeight.Text), Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionWeight_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecisionWeight.Enter
        Try
            Me.buffer = Me.TxtDefaultPrecisionWeight.Text
            Me.ObjCD.Number_Enter(Me.TxtDefaultPrecisionWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecisionWeight.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDefaultPrecisionWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionWeight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecisionWeight.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDefaultPrecisionWeight, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionWeight_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecisionWeight.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtDefaultPrecisionWeight.Text) = False Then
                MsgBox("Default precision be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtDefaultPrecisionWeight.Text = Me.buffer
                Exit Sub
            End If

            Me.TxtDefaultPrecisionWeight.Text = Me.ObjCD.DigitGrouping(Me.TxtDefaultPrecisionWeight.Text, 0)
            If CDbl(Me.TxtDefaultPrecisionWeight.Text) > 6 Then
                MsgBox("Please input default precision less than 7", MsgBoxStyle.Information, "Information")
                Me.TxtDefaultPrecisionWeight.Text = Me.buffer
                Exit Sub
            End If

            'add weight & volume
            If Me.objProduct.ApakahBolehEditDecimalWeight(Me.txtPrimaryProduct.Text, Me.txtPrimaryUnit.Text, Me.TxtPrimaryUnitWeight.Text, Me.TxtDefaultPrecisionWeight.Text) = False Then
                Me.TxtDefaultPrecisionWeight.Text = Me.buffer
                Exit Sub
            End If

            Me.TxtNettWeight.Text = Me.ObjCD.DigitGrouping(Me.TxtNettWeight.Text, Me.TxtDefaultPrecisionWeight.Text)
            Me.TxtGrossWeight.Text = Me.ObjCD.DigitGrouping(Me.TxtGrossWeight.Text, Me.TxtDefaultPrecisionWeight.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitWeight.KeyDown
        Dim i As New FrmCariMasterUnit
        Dim Obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtUnitWeight.Text) = True Then
                    Me.TxtPrimaryUnitWeight.Text = Obj.GetPrimary(Me.TxtUnitWeight.Text)
                    Me.TxtGrossWeight.Focus()
                Else
                    If Me.txtUnit.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnitWeight.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.TxtPrimaryUnitWeight.Text = i.PrimaryUnit
                        Me.TxtUnitWeight.Text = Obj.GetUnitCode(Me.TxtPrimaryUnitWeight.Text)
                        Me.TxtGrossWeight.Focus()
                    Else
                        Me.TxtUnitWeight.Clear()
                        Me.TxtPrimaryUnitWeight.Text = "-1"
                        Me.TxtUnitWeight.Focus()
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

    Private Sub TxtUnitWeight_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitWeight.Leave
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnitWeight.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtUnitWeight.Text) = True Then
                    Me.TxtPrimaryUnitWeight.Text = obj.GetPrimary(Me.TxtUnitWeight.Text)
                Else
                    If Me.TxtUnitWeight.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnitWeight.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.TxtPrimaryUnitWeight.Text = i.PrimaryUnit
                        Me.TxtUnitWeight.Text = obj.GetUnitCode(Me.TxtPrimaryUnitWeight.Text)
                    Else
                        Me.TxtUnitWeight.Clear()
                        Me.TxtPrimaryUnitWeight.Text = "-1"
                        Me.TxtUnitWeight.Focus()
                    End If
                End If
            Else
                Me.TxtUnitWeight.Clear()
                Me.TxtPrimaryUnitWeight.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnFindUnitWeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindUnitWeight.Click
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtUnitWeight.Text <> "" Then
                i.txtCari.Text = Me.TxtUnitWeight.Text
            End If

            i.ShowDialog()
            If i.PrimaryUnit > 0 Then
                Me.TxtPrimaryUnitWeight.Text = i.PrimaryUnit
                Me.TxtUnitWeight.Text = obj.GetUnitCode(Me.TxtPrimaryUnitWeight.Text)
                Me.TxtGrossWeight.Focus()
            Else
                Me.TxtPrimaryUnitWeight.Text = -1
                Me.TxtUnitWeight.Clear()
                Me.TxtUnitWeight.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtNettVolume_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettVolume.Enter
        Try
            Me.buffer = Me.TxtNettVolume.Text
            Me.ObjCD.Number_Enter(Me.TxtNettVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettVolume.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtNettVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettVolume_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettVolume.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtNettVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettVolume_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettVolume.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtNettVolume.Text) = False Then
                MsgBox("Nett volume cannot be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtNettVolume.Focus()
                Exit Sub
            End If

            Me.ObjCD.Number_Leave(Me.TxtNettVolume, sender, e, CInt(Me.TxtDefaultPrecisionVolume.Text), Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossVolume_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGrossVolume.Enter
        Try
            Me.buffer = Me.TxtGrossVolume.Text
            Me.ObjCD.Number_Enter(Me.TxtGrossVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGrossVolume.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtGrossVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossVolume_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGrossVolume.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtGrossVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGrossVolume_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGrossVolume.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtGrossVolume.Text) = False Then
                MsgBox("Gross volume cannot be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtGrossVolume.Focus()
                Exit Sub
            End If

            Me.ObjCD.Number_Leave(Me.TxtGrossVolume, sender, e, CInt(Me.TxtDefaultPrecisionVolume.Text), Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionVolume_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecisionVolume.Enter
        Try
            Me.buffer = Me.TxtDefaultPrecisionVolume.Text
            Me.ObjCD.Number_Enter(Me.TxtDefaultPrecisionVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecisionVolume.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDefaultPrecisionVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionVolume_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecisionVolume.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDefaultPrecisionVolume, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecisionVolume_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecisionVolume.Leave
        Try
            If Me.ObjCekData.CekNonNegatif(Me.TxtDefaultPrecisionVolume.Text) = False Then
                MsgBox("Default precision be a non negatif number.", MsgBoxStyle.Information, "Information")
                Me.TxtDefaultPrecisionVolume.Text = Me.buffer
                Exit Sub
            End If

            Me.TxtDefaultPrecisionVolume.Text = Me.ObjCD.DigitGrouping(Me.TxtDefaultPrecisionVolume.Text, 0)
            If CDbl(Me.TxtDefaultPrecisionVolume.Text) > 6 Then
                MsgBox("Please input default precision less than 7", MsgBoxStyle.Information, "Information")
                Me.TxtDefaultPrecisionVolume.Text = Me.buffer
                Exit Sub
            End If

            'add weight & volume
            If Me.objProduct.ApakahBolehEditDecimalVolume(Me.txtPrimaryProduct.Text, Me.txtPrimaryUnit.Text, Me.TxtPrimaryUnitVolume.Text, Me.TxtDefaultPrecisionVolume.Text) = False Then
                Me.TxtDefaultPrecisionVolume.Text = Me.buffer
                Exit Sub
            End If

            Me.TxtNettVolume.Text = Me.ObjCD.DigitGrouping(Me.TxtNettVolume.Text, Me.TxtDefaultPrecisionVolume.Text)
            Me.TxtGrossVolume.Text = Me.ObjCD.DigitGrouping(Me.TxtGrossVolume.Text, Me.TxtDefaultPrecisionVolume.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitVolume_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitVolume.KeyDown
        Dim i As New FrmCariMasterUnit
        Dim Obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtUnitVolume.Text) = True Then
                    Me.TxtPrimaryUnitVolume.Text = Obj.GetPrimary(Me.TxtUnitVolume.Text)
                    Me.TxtGrossVolume.Focus()
                Else
                    If Me.txtUnit.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnitVolume.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.TxtPrimaryUnitVolume.Text = i.PrimaryUnit
                        Me.TxtUnitVolume.Text = Obj.GetUnitCode(Me.TxtPrimaryUnitVolume.Text)
                        Me.TxtGrossVolume.Focus()
                    Else
                        Me.TxtUnitVolume.Clear()
                        Me.TxtPrimaryUnitVolume.Text = "-1"
                        Me.TxtUnitVolume.Focus()
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

    Private Sub TxtUnitVolume_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitVolume.Leave
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtUnitVolume.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtUnitVolume.Text) = True Then
                    Me.TxtPrimaryUnitVolume.Text = obj.GetPrimary(Me.TxtUnitVolume.Text)
                Else
                    If Me.TxtUnitVolume.Text <> "" Then
                        i.txtCari.Text = Me.TxtUnitVolume.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryUnit > 0 Then
                        Me.TxtPrimaryUnitVolume.Text = i.PrimaryUnit
                        Me.TxtUnitVolume.Text = obj.GetUnitCode(Me.TxtPrimaryUnitVolume.Text)
                    Else
                        Me.TxtUnitVolume.Clear()
                        Me.TxtPrimaryUnitVolume.Text = "-1"
                        Me.TxtUnitVolume.Focus()
                    End If
                End If
            Else
                Me.TxtUnitVolume.Clear()
                Me.TxtPrimaryUnitVolume.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnFindUnitVolume_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindUnitVolume.Click
        Dim i As New FrmCariMasterUnit
        Dim obj As New ClsMasterUnit

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtUnitVolume.Text <> "" Then
                i.txtCari.Text = Me.TxtUnitVolume.Text
            End If

            i.ShowDialog()
            If i.PrimaryUnit > 0 Then
                Me.TxtPrimaryUnitVolume.Text = i.PrimaryUnit
                Me.TxtUnitVolume.Text = obj.GetUnitCode(Me.TxtPrimaryUnitVolume.Text)
                Me.TxtGrossVolume.Focus()
            Else
                Me.TxtPrimaryUnitVolume.Text = -1
                Me.TxtUnitVolume.Clear()
                Me.TxtUnitVolume.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtUnitWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitWeight.GotFocus, TxtNettWeight.GotFocus, TxtGrossWeight.GotFocus, TxtDefaultPrecisionWeight.GotFocus, TxtUnitVolume.GotFocus, TxtNettVolume.GotFocus, TxtGrossVolume.GotFocus, TxtDefaultPrecisionVolume.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitWeight.LostFocus, TxtNettWeight.LostFocus, TxtGrossWeight.LostFocus, TxtDefaultPrecisionWeight.LostFocus, TxtUnitVolume.LostFocus, TxtNettVolume.LostFocus, TxtGrossVolume.LostFocus, TxtDefaultPrecisionVolume.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryUnit.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindTaxProductRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindTaxProductRef.Click
        Dim obj As New ClsMasterTaxProductReference
        Dim hj As New FrmMasterTaxProductReference_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtTaxProductRef.Text <> "" Then
                hj.txtCari.Text = Me.TxtTaxProductRef.Text
            End If

            hj.ShowDialog()
            If hj.Primary > 0 Then
                Me.TxtPrimaryTaxProductRef.Text = hj.Primary
                Me.TxtTaxProductRef.Text = obj.GetCode(Me.TxtPrimaryTaxProductRef.Text)
                Me.TxtTaxProductDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxProductRef.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryTaxProductRef.Text = "-1"
                Me.TxtTaxProductDesc.Clear()
                Me.TxtTaxProductRef.Clear()
                Me.TxtTaxProductRef.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtTaxProductRef_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxProductRef.KeyDown
        Dim obj As New ClsMasterTaxProductReference
        Try
            If e.KeyValue = 13 Then
                If Me.TxtTaxProductRef.Text <> "" Then
                    If obj.ApakahSudahAda(Me.TxtTaxProductRef.Text) = True Then
                        Me.TxtPrimaryTaxProductRef.Text = obj.GetPrimary(Me.TxtTaxProductRef.Text)
                        Me.TxtTaxProductRef.Text = obj.GetCode(Me.TxtPrimaryTaxProductRef.Text)
                        Me.TxtTaxProductDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxProductRef.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.BtnFindTaxProductRef.PerformClick()
                    End If
                Else
                    Me.BtnFindTaxProductRef.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxProductRef_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxProductRef.Leave
        Dim obj As New ClsMasterTaxProductReference
        Dim hj As New FrmMasterTaxProductReference_Cari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtTaxProductRef.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtTaxProductRef.Text) = True Then
                    Me.TxtPrimaryTaxProductRef.Text = obj.GetPrimary(Me.TxtTaxProductRef.Text)
                    Me.TxtTaxProductRef.Text = obj.GetCode(Me.TxtPrimaryTaxProductRef.Text)
                    Me.TxtTaxProductDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxProductRef.Text)
                Else
                    If Me.TxtTaxProductRef.Text <> "" Then
                        hj.txtCari.Text = Me.TxtTaxProductRef.Text
                    End If
                    hj.ShowDialog()

                    If hj.Primary > 0 Then
                        Me.TxtPrimaryTaxProductRef.Text = hj.Primary
                        Me.TxtTaxProductRef.Text = obj.GetCode(Me.TxtPrimaryTaxProductRef.Text)
                        Me.TxtTaxProductDesc.Text = obj.GetDesc(Me.TxtPrimaryTaxProductRef.Text)
                    Else
                        Me.TxtPrimaryTaxProductRef.Text = "-1"
                        Me.TxtTaxProductDesc.Clear()
                        Me.TxtTaxProductRef.Clear()
                        Me.TxtTaxProductRef.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryTaxProductRef.Text = "-1"
                Me.TxtTaxProductDesc.Clear()
                Me.TxtTaxProductRef.Clear()
                'Me.TxtRemarks.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtTaxProductRef_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxProductRef.GotFocus
        TxtTaxProductRef.BackColor = BackColorAll
    End Sub

    Private Sub TxtTaxProductRef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxProductRef.LostFocus
        TxtTaxProductRef.BackColor = Me.TxtPrimaryTaxProductRef.BackColor
    End Sub
End Class