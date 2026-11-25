Option Explicit On
Option Strict Off

Imports system.io

Public Class frmMasterProduct
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim baru As Boolean
    Dim objProduct As clsMasterProduct
    Dim ObjCekData As New ClsCekData
    Dim ObjCategory As New ClsMasterCategory
    Dim buffer As String = ""
    Dim ObjCD As New ClsCekData

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.txtPrimaryProduct.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimaryProduct.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.chkUseSecondaryUnit.Enabled = True
            Me.chkLocation.Enabled = True
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
            'jika edit dan product sudah pernah digunakan maka check box expiry date, batch no, location, 
            'use secondary unit
            If Me.baru = False Then
                If Me.objProduct.ApakahSudahPernahDigunakan(Me.txtPrimaryProduct.Text) = True Then
                    Me.chkBatchNo.Enabled = False
                    Me.chkExpiryDate.Enabled = False
                    Me.chkLocation.Enabled = False
                    Me.chkPDOControlled.Enabled = False
                    Me.chkUseSecondaryUnit.Enabled = False
                End If
            End If
            Me.chkSizing.Enabled = True
            If Me.chkSizing.Checked = True Then
                Me.cboSizing.Enabled = True
            Else
                Me.cboSizing.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.chkUseSecondaryUnit.Enabled = False
            Me.chkLocation.Enabled = False
            Me.chkBatchNo.Enabled = False
            Me.chkPDOControlled.Enabled = False
            Me.chkCategoryBindCode.Enabled = False
            Me.chkExpiryDate.Enabled = False
            Me.chkCategoryBindDecs.Enabled = False
            Me.cboSizing.Enabled = False
            Me.chkSizing.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.TabPage2.Controls, True, False, False)

            Me.txtPrimaryProduct.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryMasterUnit.Text = "-1"
            Me.txtPrimarySecondaryUnit.Text = "-1"
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
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

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'data salah bila:
            '1.tidak ada data pricing
            If Me.dgvPricing.RowCount = 0 Then
                MsgBox("Please input pricing data.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
                Me.txtCurrencyCode.Focus()
                Return False
            End If

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
                        Me.objProduct.CariDetail(Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. qty per master unit negatif
            'kalo master unit tidak kosong maka qty per master unit harus positif
            If Me.txtPrimaryMasterUnit.Text <> "-1" Then
                If Me.ObjCekData.CekBilanganPositif(Me.txtQtyPerMasterUnit.Text) = False Then
                    MsgBox("Please enter a positif number for quantity per master unit data.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtQtyPerMasterUnit.Focus()
                    Return False
                End If
            Else
                If Me.ObjCekData.CekNonNegatif(Me.txtQtyPerMasterUnit.Text) = False Then
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

            'default precision maximal bernilai 3
            If CInt(Me.TxtDefaultPrecision.Text) > 3 Then
                MsgBox("Default precision maximum value is 3.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtDefaultPrecision.Text = "3"
                Me.TxtDefaultPrecision.Focus()
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
                .PrimaryMasterUnit = Me.txtPrimaryMasterUnit.Text
                .PrimaryUnit = Me.txtPrimaryUnit.Text
                If Me.chkCategoryBindCode.Checked = True Then
                    .ProductCode = Trim(Me.txtCategoryCode.Text) + Trim(Me.txtProductCode.Text)
                Else
                    .ProductCode = Trim(Me.txtProductCode.Text)
                End If
                .QtyPerMasterUnit = Me.txtQtyPerMasterUnit.Text
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
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim ObjCat As New ClsMasterCategory
        Dim ObjUnit As New ClsMasterUnit

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCat = Nothing
            ObjUnit = Nothing
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvPricing.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.objProduct = New clsMasterProduct
            Me.baru = True
            Me.NonAktifkanDetail()
            Dim Obj1 As New clsMasterListedSoftware
            If Obj1.IsInventoryInstalled = True Then
                Me.chkBatchNo.Visible = True
                Me.chkExpiryDate.Visible = True
                Me.chkLocation.Visible = True
                If Obj1.IsProductionInstalled = False Then
                    Me.chkLocation.Left = 423
                    Me.chkExpiryDate.Left = 199
                    Me.chkPDOControlled.Visible = False
                    Me.chkPDOControlled.TabStop = False
                Else
                    Me.chkLocation.Left = 295
                    Me.chkExpiryDate.Left = 135
                    Me.chkPDOControlled.Visible = True
                    Me.chkPDOControlled.TabStop = True
                End If
            Else
                Me.chkBatchNo.Visible = False
                Me.chkExpiryDate.Visible = False
                Me.chkLocation.Visible = False
                Me.chkPDOControlled.Visible = False
            End If

            Me.cboSizing.SelectedIndex = 0
            Obj1 = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                        If Me.objProduct.SimpanBaru(Me.UserCode, Me.pbxProduct, Me.dgvPricing) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objProduct.SimpanPerubahan(Me.UserCode, Me.pbxProduct, Me.dgvPricing) Then
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
            If Not Me.txtProductCode.Text = "" And Me.txtDescription.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objProduct.CariDetail(Me.objProduct.Primary, Me.dgvPricing, Me.pbxProduct)
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
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TabControl1.SelectedIndex = 0
            Me.txtCategoryCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
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
                If Me.objProduct.CariDetail(hj.PrimaryProduct, Me.dgvPricing, Me.pbxProduct) = True Then
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
                    MsgBox("You cannot delete this data, there are other people using this data.", _
                                        MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objProduct.Hapus(Me.txtPrimaryProduct.Text) Then
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
                            If Me.objProduct.SimpanBaru(Me.UserCode, Me.pbxProduct, Me.dgvPricing) = True Then
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
                            If Me.objProduct.SimpanPerubahan(Me.UserCode, Me.pbxProduct, Me.dgvPricing) = True Then
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
            Else
                Me.txtPrimaryUnit.Text = "-1"
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
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                    End If
                End If
            Else
                Me.LblProduct.Text = Me.txtProductCode.Text + " - " + Me.txtDescription.Text
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
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
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
                    Else
                        Me.LblProduct.Text = "Product Code - Description"
                    End If
                End If
            Else
                Me.LblProduct.Text = Me.txtProductCode.Text + " - " + Me.txtDescription.Text
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
                Me.txtProductCode.Left = 264
                Me.txtProductCode.Width = Me.txtDescription.Width - (Me.txtProductCode.Left - Me.txtDescription.Left)
                If Me.txtPrimaryCategory.Text <> "-1" Then
                    If Me.txtProductCode.Text.StartsWith(Me.txtCategoryCode2.Text) = True Then
                        Me.txtProductCode.Text = Me.txtProductCode.Text.Substring(Me.txtCategoryCode2.Text.Length)
                    End If
                End If
                Me.txtProductCode_TextChanged(sender, e)
            Else
                Me.txtCategoryCode2.Visible = False
                Me.txtProductCode.Left = 141
                Me.txtProductCode.Width = Me.txtDescription.Width
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

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.txtCategoryCode.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 0 Then
                    Me.txtCategoryCode.Focus()
                ElseIf Me.TabControl1.SelectedIndex = 1 Then
                    Me.txtCurrencyCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefaultPrecision_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecision.KeyDown
        Me.ObjCekData.Number_KeyDown(Me.TxtDefaultPrecision, sender, e)
    End Sub

    Private Sub TxtDefaultPrecision_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefaultPrecision.KeyUp
        Me.ObjCekData.Number_KeyUp(Me.TxtDefaultPrecision, sender, e)
    End Sub

    Private Sub TxtDefaultPrecision_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.GotFocus
        TxtDefaultPrecision.BackColor = BackColorAll
    End Sub

    Private Sub TxtDefaultPrecision_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.LostFocus
        TxtDefaultPrecision.BackColor = Me.txtPrimaryProduct.BackColor
    End Sub

    Private Sub TxtDefaultPrecision_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefaultPrecision.Leave
        Try
            If Me.txtCategoryCode.Enabled = True Then
                Me.TxtDefaultPrecision.Text = Math.Floor(CDbl(Me.TxtDefaultPrecision.Text))
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
End Class