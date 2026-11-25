Option Explicit On
Option Strict Off
Public Class FrmMasterProductAccountGroup
   
    Public UserCode As String = "USER"
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim vDef_Prec As Integer
    Dim baru As Boolean = True
    Dim objMPA As New ClsMasterProductAccountGroup
    Dim objAkun As New ClsMasterAkun1
    Dim objSalesAkun As New ClsMasterSalesPerson
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objGroup As New ClsMasterProductGroup
    Dim objCategory As New ClsMasterCategory
    Dim st As Boolean = False
   
    Private Sub AktifkanDetail()
        Try
            ObjCD.SetBackColor(Me.GboDetail1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.GboDetail2.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.ScAccount.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.ScAccount.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.GboDetail1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.GboDetail2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.ScAccount.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.TxtGroupDesc.BackColor = Me.TxtReadOnly.BackColor
            Me.RboGoods.Enabled = True
            Me.RboService.Enabled = True
            Me.ScAccount.TabStop = False
            If baru = True Then
                RboGoods.Checked = True
                RboService.Checked = False
                Me.TxtCode.Focus()
            End If
            '>>>jika goods maka yang aktif semua
            If RboGoods.Checked = True Then
                ModCekData.KendalikanControl(Me.ScAccount.Panel1.Controls, True, True, True, True, True, True, True, True)
                ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, True, True, True, True, True, True, True)
            Else
                '>>> jika service yang aktif hanya sebelah kiri
                ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, False, True, False, True, False, True, False)
            End If

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ObjCD.SetBackColor(Me.GboDetail1.Controls, Me.TxtPrimary.BackColor, True)
            ObjCD.SetBackColor(Me.GboDetail2.Controls, Me.TxtPrimary.BackColor, True)
            ObjCD.SetBackColor(Me.ScAccount.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            ObjCD.SetBackColor(Me.ScAccount.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.Controls, True, False, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.GboDetail1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.GboDetail2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.ScAccount.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ScAccount.TabStop = True
            Me.RboGoods.Enabled = False
            Me.RboService.Enabled = False
            Me.CancelButton = Me.btnExit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.GboDetail1.Controls, True, False, True)
            ModCekData.KosongkanControl(Me.ScAccount.Panel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.ScAccount.Panel2.Controls, True, False, False)
            '>>>utk Primary
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryAccountAdjustment.Text = "-1"
            Me.TxtPrimaryAccountCOGS.Text = "-1"
            Me.TxtPrimaryAccountInventory.Text = "-1"
            Me.TxtPrimaryAccountSales.Text = "-1"
            Me.TxtPrimaryAccountSalesReturn.Text = "-1"
            Me.TxtPrimaryAccountStockTaking.Text = "-1"
            Me.TxtPrimaryGroup.Text = "-1"
            '--
            Me.CboGroupLevel.SelectedIndex = -1
            Me.RboGoods.Checked = True
            Me.RboService.Checked = False
            Me.TslCreated_By.Text = "Created By: "
            Me.TslCreated_Date.Text = "Created Date: "
            Me.TslModified_By.Text = "Modified By: "
            Me.TslModified_Date.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TipeAkun As ClsMasterAkun1.EnumTipeAkun2 = ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim TipeAkun1 As ClsMasterAkun1.EnumTipeAkun = ClsMasterAkun1.EnumTipeAkun.Asset
        Dim PrimaryLama As Double = -1

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtCode)
            KoleksiNamaData.Add("Desc")
            KoleksiTextBox.Add(Me.TxtDesc)
            KoleksiNamaData.Add("Group code")
            KoleksiTextBox.Add(Me.TxtGroupCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data yang salah, double dalam kode
            If Me.baru = True Then
                If Me.objMPA.ApakahSudahAda(Me.TxtCode.Text) = True Then
                    MsgBox("Duplication in code, please change your Code", _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtCode.Focus()
                    Return False
                End If
            Else
                Dim ProductGroupLama As String = Me.objMPA.GetCode(Me.TxtPrimary.Text)
                If ProductGroupLama <> Me.TxtCode.Text Then
                    If Me.objMPA.ApakahSudahAda(Me.TxtCode.Text) = True Then
                        MsgBox("Duplication in code, please change your Code", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.objMPA.GetTimeStamp(Me.TxtPrimary.Text)
                If Me.objMPA.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objMPA.CariDetail(Me.TxtPrimary.Text)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If
            '>>>kalo goods hrus isi semua akun
            If Me.RboGoods.Checked = True Then
                KoleksiNamaData.Add("Inventory Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountInventory)
                KoleksiNamaData.Add("Sales Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountSales)
                KoleksiNamaData.Add("Sales Return Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountSalesReturn)
                KoleksiNamaData.Add("Cost of Sales Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountCOGS)
                KoleksiNamaData.Add("Adjustment Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountAdjustment)
                KoleksiNamaData.Add("Stock Taking Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountStockTaking)
                If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                    Return False
                End If
            Else
                '>>>kalo bkn hanya sebelah kiri saja
                KoleksiNamaData.Add("Inventory Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountInventory)
                KoleksiNamaData.Add("Sales Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountSales)
                KoleksiNamaData.Add("Sales Return Account")
                KoleksiTextBox.Add(Me.TxtCodeAccountSalesReturn)
                If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                    Return False
                End If
                Me.Reset_Object()
            End If
            ''

            'cek tipe akun
            TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountInventory.Text)
            If Me.RboGoods.Checked = True Then
                'tipe akun inventory haruslah inventory --> goods
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Inventory Then
                    MsgBox("Please specify an inventory type account for inventory account", MsgBoxStyle.Information, "Information")
                    Me.TxtCodeAccountInventory.Focus()
                    Return False
                End If
            Else
                'tipe akun inventory haruslah expense/other expense --> service
                TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountInventory.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    'boleh juga akun asset
                    TipeAkun1 = Me.objAkun.GetTipeAkun(Me.TxtPrimaryAccountInventory.Text)
                    If TipeAkun1 <> ClsMasterAkun1.EnumTipeAkun.Asset Then
                        MsgBox("Please specify an expense/asset type account for expense account", MsgBoxStyle.Information, "Information")
                        Me.TxtCodeAccountInventory.Focus()
                        Return False
                    End If
                End If
            End If
            'tipe akun sales dan sales return haruslah sales return
            TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountSales.Text)
            If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                MsgBox("Please specify an income/expense type account for sales account", MsgBoxStyle.Information, "Information")
                Me.TxtCodeAccountSales.Focus()
                Return False
            End If
            TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountSalesReturn.Text)
            If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Income And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                MsgBox("Please specify an income/expense type account for sales return account", MsgBoxStyle.Information, "Information")
                Me.TxtCodeAccountSalesReturn.Focus()
                Return False
            End If
            If Me.RboGoods.Checked = True Then
                'tipe akun cogs haruslah cogs
                TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountCOGS.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an cost of sales type account for cost of sales account", MsgBoxStyle.Information, "Information")
                    Me.TxtCodeAccountCOGS.Focus()
                    Return False
                End If
                'tipe akun adjustment dan stock taking account harus lah expense, other expense
                TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountAdjustment.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an expense type account for adjustment account", MsgBoxStyle.Information, "Information")
                    Me.TxtCodeAccountAdjustment.Focus()
                    Return False
                End If
                TipeAkun = Me.objAkun.GetTipeAccount2(Me.TxtPrimaryAccountStockTaking.Text)
                If TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Other_Expense And _
                 TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Expense And TipeAkun <> ClsMasterAkun1.EnumTipeAkun2.Cost_Of_Sales Then
                    MsgBox("Please specify an expense type account for stock taking account", MsgBoxStyle.Information, "Information")
                    Me.TxtCodeAccountStockTaking.Focus()
                    Return False
                End If
            End If
            If String.IsNullOrEmpty(Trim(Me.TxtCode.Text)) = True Then
                MsgBox("Please input code", MsgBoxStyle.Information, "Information")
                Me.TxtCode.Focus()
                Return False
            End If
            '--
            Dim i As Integer
            If Me.RboGoods.Checked = True Then
                i = 0
            Else
                i = 1
            End If


            If baru = True Then
                If Me.objMPA.GetPrimary(Me.GetGroupLevel, Me.TxtPrimaryGroup.Text) <> -1 Then
                    MsgBox("Duplication in group code and type, please change your group code and type", MsgBoxStyle.Information, "Information")
                    Me.TxtGroupCode.Focus()
                    Return False
                End If
            Else
                'If objMPA.ApakahAdaPrimary(Me.TxtPrimary.Text, i, Me.TxtPrimaryGroup.Text) = True Then
                '    MsgBox("Duplication in group code and type, please change your group code and type", MsgBoxStyle.Information, "Information")
                '    Me.TxtGroupCode.Focus()
                '    Return False
                'End If
                PrimaryLama = Me.objMPA.GetPrimary(Me.GetGroupLevel, Me.TxtPrimaryGroup.Text)
                If PrimaryLama <> -1 Then
                    If Me.TxtPrimary.Text <> PrimaryLama Then
                        MsgBox("Duplication in group code and type, please change your group code and type", _
                        MsgBoxStyle.Information, "Information")
                        Me.TxtGroupCode.Focus()
                        Return False
                    End If
                End If
            End If

            If Me.TxtPrimaryGroup.Text <> "" Then
                If CariygSama() = True Then
                    MsgBox("Duplication in group code and type, please change your group code and type", _
                            MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


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
            With Me.objMPA

                '>>>utk Primary
                .Primary = Me.TxtPrimary.Text
                .Primary_Adjustment_Accont = Me.TxtPrimaryAccountAdjustment.Text
                .Primary_Cost_Of_Sales_Account = Me.TxtPrimaryAccountCOGS.Text
                .Primary_Group = Me.TxtPrimaryGroup.Text
                .Primary_Inventory_Account = Me.TxtPrimaryAccountInventory.Text
                .Primary_Sales_Account = Me.TxtPrimaryAccountSales.Text
                .Primary_Sales_Return_Account = Me.TxtPrimaryAccountSalesReturn.Text
                .Primary_Stock_taking_Account = Me.TxtPrimaryAccountStockTaking.Text
                '--
                .Code = Me.TxtCode.Text
                .Desc = Me.TxtDesc.Text
                .Remarks = Me.TxtRemarks.Text

                '>>>utk group level
                If Me.CboGroupLevel.SelectedIndex = 0 Then
                    .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
                ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                    .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.Brand
                ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                    .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
                ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                    .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.Category
                End If
                '--
                '>>>utk type
                If RboGoods.Checked = True Then
                    .Type = ClsMasterProductAccountGroup.EnumTipe.Goods
                Else
                    .Type = ClsMasterProductAccountGroup.EnumTipe.Service
                End If
                '--
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetGroupLevel() As ClsMasterProductAccountGroup.EnumGroupLevel
        Try
            If Me.CboGroupLevel.SelectedIndex = 0 Then
                Return ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
            ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                Return ClsMasterProductAccountGroup.EnumGroupLevel.Brand
            ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                Return ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand
            ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                Return ClsMasterProductAccountGroup.EnumGroupLevel.Category
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup
        End Try
    End Function

    Private Sub SalinKeField()
        Try
            With Me.objMPA
                '>>>For Primary--
                Me.TxtPrimary.Text = .Primary
                Me.TxtPrimaryAccountAdjustment.Text = .Primary_Adjustment_Accont
                Me.TxtPrimaryAccountCOGS.Text = .Primary_Cost_Of_Sales_Account
                Me.TxtPrimaryAccountInventory.Text = .Primary_Inventory_Account
                Me.TxtPrimaryAccountSales.Text = .Primary_Sales_Account
                Me.TxtPrimaryAccountSalesReturn.Text = .Primary_Sales_Return_Account
                Me.TxtPrimaryAccountStockTaking.Text = .Primary_Stock_taking_Account
                Me.TxtPrimaryGroup.Text = .Primary_Group
                '--
                '>>>for Cbo group
                If .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.ProductGroup Then
                    CboGroupLevel.SelectedIndex = 0
                    Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(.Primary_Group)
                    Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(.Primary_Group)

                ElseIf .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.Brand Then
                    CboGroupLevel.SelectedIndex = 1
                    Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(.Primary_Group)
                    Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(.Primary_Group)

                ElseIf .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.SubBrand Then
                    CboGroupLevel.SelectedIndex = 2
                    Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(.Primary_Group)
                    Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(.Primary_Group)

                ElseIf .Group_Level = ClsMasterProductAccountGroup.EnumGroupLevel.Category Then
                    CboGroupLevel.SelectedIndex = 3
                    Me.TxtGroupCode.Text = Me.objCategory.GetCategoryCode(.Primary_Group)
                    Me.TxtGroupDesc.Text = Me.objCategory.GetDesc(.Primary_Group)

                End If
                '--
                '>>>For RboType
                If .Type = ClsMasterProductAccountGroup.EnumTipe.Goods Then
                    RboGoods.Checked = True
                Else
                    RboService.Checked = True
                End If
                '--
                '>>>For Code
                Me.TxtCode.Text = .Code
                Me.TxtCodeAccountAdjustment.Text = Me.objAkun.GetAccountCode(.Primary_Adjustment_Accont)
                Me.TxtCodeAccountCOGS.Text = Me.objAkun.GetAccountCode(.Primary_Cost_Of_Sales_Account)
                Me.TxtCodeAccountInventory.Text = Me.objAkun.GetAccountCode(.Primary_Inventory_Account)
                Me.TxtCodeAccountSales.Text = Me.objAkun.GetAccountCode(.Primary_Sales_Account)
                Me.TxtCodeAccountSalesReturn.Text = Me.objAkun.GetAccountCode(.Primary_Sales_Return_Account)
                Me.TxtCodeAccountStockTaking.Text = Me.objAkun.GetAccountCode(.Primary_Stock_taking_Account)
                '--
                '>>>For Desc
                Me.TxtDesc.Text = .Desc
                Me.TxtDescAccountAdjustment.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Adjustment_Accont)
                Me.TxtDescAccountCOGS.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Cost_Of_Sales_Account)
                Me.TxtDescAccountInventory.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Inventory_Account)
                Me.TxtDescAccountSales.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Sales_Account)
                Me.TxtDescAccountSalesReturn.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Sales_Return_Account)
                Me.TxtDescAccountStockTaking.Text = Me.objAkun.GetDescription_ByPrimaryAcc(.Primary_Stock_taking_Account)
                '--
                Me.TxtRemarks.Text = .Remarks

                Me.TslCreated_By.Text = "Created By: " + .CreatedBy
                Me.TslCreated_Date.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.TslModified_By.Text = "Modified By: " + .ModifiedBy
                Me.TslModified_Date.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'Tipe disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objMPA.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'Tipe disimpan atau tidak
                    If Me.objMPA.ApakahBolehEdit(Me.TxtPrimary.Text, UserCode, True, True) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.objMPA.SimpanPerubahan(Me.UserCode) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
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
            If Not Me.TxtCode.Text = "" Then
                'Tipe batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
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
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TxtCode.Focus()
            Me.CboGroupLevel.SelectedIndex = 0
            Me.TslCreated_By.Text = "Created By: " + Me.UserCode
            Me.TslCreated_Date.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TslModified_By.Text = "Modified By: " + Me.UserCode
            Me.TslModified_Date.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterProductAccountGroup_Cari
        Try
            hj.BackColorAll = Me.BAckColorAll
            hj.ShowDialog()

            If Not hj.Primary = -1 Then
                If Me.objMPA.CariDetail(hj.Primary) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()

                End If
                Me.btnFind.Focus()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                If Me.objMPA.ApakahBolehEdit(Me.objMPA.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.objMPA.CariDetail(Me.objMPA.Primary)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtCode.Text <> "" Then
                If Me.objMPA.ApakahBolehHapus(Me.TxtPrimary.Text, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.objMPA.Hapus(Me.TxtPrimary.Text, UserCode) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
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
            If Me.TxtCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objMPA.SimpanBaru(Me.UserCode) = True Then
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
                            If Me.objMPA.SimpanPerubahan(Me.UserCode) = True Then
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

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.GotFocus, TxtGroupCode.GotFocus, TxtDesc.GotFocus, _
    TxtRemarks.GotFocus, TxtCodeAccountInventory.GotFocus, TxtCodeAccountSales.GotFocus, TxtCodeAccountSalesReturn.GotFocus, TxtCodeAccountAdjustment.GotFocus _
    , TxtCodeAccountCOGS.GotFocus, TxtCodeAccountStockTaking.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtcode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.Leave
        Try
            Me.TxtCode.Text = Me.TxtCode.Text.ToCharArray
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCode.LostFocus, TxtGroupCode.LostFocus, TxtDesc.LostFocus, _
    TxtRemarks.LostFocus, TxtCodeAccountInventory.LostFocus, TxtCodeAccountSales.LostFocus, TxtCodeAccountSalesReturn.LostFocus, TxtCodeAccountAdjustment.LostFocus _
    , TxtCodeAccountCOGS.LostFocus, TxtCodeAccountStockTaking.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '>>>Group 
    Private Sub TxtGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtGroupCode.KeyDown
        Try

            If Me.CboGroupLevel.SelectedIndex = 0 Or Me.CboGroupLevel.SelectedIndex = 1 Or Me.CboGroupLevel.SelectedIndex = 2 Then
                Dim hj As New FrmCariMasterProductGroup
                '>>>tentukan level group sesuai cbo
                If Me.CboGroupLevel.SelectedIndex = 0 Then
                    hj.ProductGroupLevel = 3
                ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                    hj.ProductGroupLevel = 2
                ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                    hj.ProductGroupLevel = 1
                ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                    hj.ProductGroupLevel = 0
                End If
                If e.KeyValue = 13 Then
                    If Me.TxtGroupCode.Text <> "" Then
                        If Me.objGroup.ApakahSudahAda(Me.TxtGroupCode.Text, hj.ProductGroupLevel) = True Then
                            Me.TxtPrimaryGroup.Text = Me.objGroup.GetPrimary(Me.TxtGroupCode.Text, hj.ProductGroupLevel)
                            Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(Me.TxtPrimaryGroup.Text)
                            Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(Me.TxtPrimaryGroup.Text)
                            Me.RboGoods.Focus()
                        Else
                            Me.BtnFindGroup.PerformClick()
                        End If
                    Else
                        Me.BtnFindGroup.PerformClick()
                    End If
                End If
            ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                If e.KeyValue = 13 Then
                    If Me.TxtGroupCode.Text <> "" Then
                        If Me.objCategory.ApakahSudahAda(Me.TxtGroupCode.Text) = True Then
                            Me.TxtPrimaryGroup.Text = Me.objCategory.GetPrimary(Me.TxtGroupCode.Text)
                            Me.TxtGroupCode.Text = Me.objCategory.GetCategoryCode(Me.TxtPrimaryGroup.Text)
                            Me.TxtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGroup.Text)
                            Me.RboGoods.Focus()
                        Else
                            Me.BtnFindGroup.PerformClick()
                        End If
                    Else
                        Me.BtnFindGroup.PerformClick()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGroupCode.Leave
        Try
            '--
            '>>>tentukan level group sesuai cbo
            If Me.CboGroupLevel.SelectedIndex = 0 Or Me.CboGroupLevel.SelectedIndex = 1 Or Me.CboGroupLevel.SelectedIndex = 2 Then
                Dim hj As New FrmCariMasterProductGroup
                If Me.CboGroupLevel.SelectedIndex = 0 Then
                    hj.ProductGroupLevel = 3
                ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                    hj.ProductGroupLevel = 2
                ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                    hj.ProductGroupLevel = 1
                ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                    hj.ProductGroupLevel = 0
                End If


                hj.BackColorAll = Me.BAckColorAll
                Try
                    If Me.TxtGroupCode.Text <> "" Then
                        If Me.objGroup.ApakahSudahAda(Me.TxtGroupCode.Text, hj.ProductGroupLevel) = True Then
                            Me.TxtPrimaryGroup.Text = Me.objGroup.GetPrimary(Me.TxtGroupCode.Text, hj.ProductGroupLevel)
                            Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(Me.TxtPrimaryGroup.Text)
                            Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(Me.TxtPrimaryGroup.Text)
                        Else
                            If Me.TxtGroupCode.Text <> "" Then
                                hj.txtCari.Text = Me.TxtGroupCode.Text
                            End If
                            hj.ShowDialog()

                            If hj.PrimaryProductGroup > 0 Then
                                Me.TxtPrimaryGroup.Text = hj.PrimaryProductGroup
                                Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(Me.TxtPrimaryGroup.Text)
                                Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(Me.TxtPrimaryGroup.Text)
                                Me.RboGoods.Focus()
                            Else
                                Me.TxtPrimaryGroup.Text = "-1"
                                Me.TxtGroupDesc.Clear()
                                Me.TxtGroupCode.Clear()
                                Me.TxtGroupCode.Focus()
                            End If
                        End If
                    Else
                        Me.TxtPrimaryGroup.Text = "-1"
                        Me.TxtGroupDesc.Clear()
                        Me.TxtGroupCode.Clear()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                Dim hj As New frmCariMasterCategory
                hj.BackColorAll = Me.BAckColorAll
                Try
                    If Me.TxtGroupCode.Text <> "" Then
                        If Me.objCategory.ApakahSudahAda(Me.TxtGroupCode.Text) = True Then
                            Me.TxtPrimaryGroup.Text = Me.objCategory.GetPrimary(Me.TxtGroupCode.Text)
                            Me.TxtGroupCode.Text = Me.objCategory.GetCategoryCode(Me.TxtPrimaryGroup.Text)
                            Me.TxtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGroup.Text)
                        Else
                            If Me.TxtGroupCode.Text <> "" Then
                                hj.txtCari.Text = Me.TxtGroupCode.Text
                            End If
                            hj.ShowDialog()

                            If hj.PrimaryCategory > 0 Then
                                Me.TxtPrimaryGroup.Text = hj.PrimaryCategory
                                Me.TxtGroupCode.Text = Me.objCategory.GetCategoryCode(Me.TxtPrimaryGroup.Text)
                                Me.TxtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGroup.Text)
                                Me.RboGoods.Focus()
                            Else
                                Me.TxtPrimaryGroup.Text = "-1"
                                Me.TxtGroupDesc.Clear()
                                Me.TxtGroupCode.Clear()
                                Me.TxtGroupCode.Focus()
                            End If
                        End If
                    Else
                        Me.TxtPrimaryGroup.Text = "-1"
                        Me.TxtGroupDesc.Clear()
                        Me.TxtGroupCode.Clear()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindGroup.Click
        Try
            If Me.CboGroupLevel.SelectedIndex = 0 Or Me.CboGroupLevel.SelectedIndex = 1 Or Me.CboGroupLevel.SelectedIndex = 2 Then
                Dim hj As New FrmCariMasterProductGroup
                hj.BackColorAll = Me.BAckColorAll

                '>>>tentukan level group sesuai cbo
                If Me.CboGroupLevel.SelectedIndex = 0 Then
                    hj.ProductGroupLevel = 3
                ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                    hj.ProductGroupLevel = 2
                ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                    hj.ProductGroupLevel = 1
                End If
                '--
                Try

                    If Me.TxtGroupCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtGroupCode.Text
                    End If
                    hj.ShowDialog()
                    If hj.PrimaryProductGroup > 0 Then
                        Me.TxtPrimaryGroup.Text = hj.PrimaryProductGroup
                        Me.TxtGroupCode.Text = Me.objGroup.GetGroupCode(Me.TxtPrimaryGroup.Text)
                        Me.TxtGroupDesc.Text = Me.objGroup.GetGroupDesc(Me.TxtPrimaryGroup.Text)
                        Me.RboGoods.Focus()
                    Else
                        Me.TxtPrimaryGroup.Text = "-1"
                        Me.TxtGroupDesc.Clear()
                        Me.TxtGroupCode.Clear()
                        Me.TxtGroupCode.Focus()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            ElseIf Me.CboGroupLevel.SelectedIndex = 3 Then
                Dim hj As New frmCariMasterCategory
                hj.BackColorAll = Me.BAckColorAll
                Try
                    If Me.TxtGroupCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtGroupCode.Text
                    End If

                    hj.ShowDialog()
                    If hj.PrimaryCategory > 0 Then
                        Me.TxtPrimaryGroup.Text = hj.PrimaryCategory
                        Me.TxtGroupCode.Text = Me.objCategory.GetCategoryCode(Me.TxtPrimaryGroup.Text)
                        Me.TxtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGroup.Text)
                        Me.TxtCodeAccountInventory.Focus()
                    Else
                        Me.TxtPrimaryGroup.Text = "-1"
                        Me.TxtGroupDesc.Clear()
                        Me.TxtGroupCode.Clear()
                        Me.TxtGroupCode.Focus()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                Finally
                    hj.Dispose()
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '--
    '>>>Inventory Account Inventory
    Private Sub TxtCodeAccountInventory_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountInventory.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountInventory.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountInventory.Text) = True Then
                        Me.TxtPrimaryAccountInventory.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountInventory.Text)
                        Me.TxtCodeAccountInventory.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountInventory.Text)
                        Me.TxtDescAccountInventory.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountInventory.Text)
                        Me.TxtCodeAccountSales.Focus()
                    Else
                        Me.BtnFindAccountInventory.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountInventory.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountInventory_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountInventory.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
         

            If Me.TxtCodeAccountInventory.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountInventory.Text) = True Then
                    Me.TxtPrimaryAccountInventory.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountInventory.Text)
                    Me.TxtCodeAccountInventory.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountInventory.Text)
                    Me.TxtDescAccountInventory.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountInventory.Text)
                Else
                    If Me.TxtCodeAccountInventory.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountInventory.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountInventory.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountInventory.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountInventory.Text)
                        Me.TxtDescAccountInventory.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountInventory.Text)
                        Me.TxtCodeAccountSales.Focus()
                    Else
                        Me.TxtPrimaryAccountInventory.Text = "-1"
                        Me.TxtDescAccountInventory.Clear()
                        Me.TxtCodeAccountInventory.Clear()
                        Me.TxtCodeAccountInventory.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountInventory.Text = "-1"
                Me.TxtDescAccountInventory.Clear()
                Me.TxtCodeAccountInventory.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountInventory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountInventory.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountInventory.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountInventory.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountInventory.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountInventory.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountInventory.Text)
                Me.TxtDescAccountInventory.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountInventory.Text)
                Me.TxtCodeAccountSales.Focus()
            Else
                Me.TxtPrimaryAccountInventory.Text = "-1"
                Me.TxtDescAccountInventory.Clear()
                Me.TxtCodeAccountInventory.Clear()
                Me.TxtCodeAccountInventory.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    '>>>Sales Account
    Private Sub TxtCodeAccountSales_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountSales.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountSales.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountSales.Text) = True Then
                        Me.TxtPrimaryAccountSales.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountSales.Text)
                        Me.TxtCodeAccountSales.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSales.Text)
                        Me.TxtDescAccountSales.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSales.Text)
                        Me.TxtCodeAccountSalesReturn.Focus()
                    Else
                        Me.BtnFindAccountSales.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountSales.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountSales_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountSales.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeAccountSales.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountSales.Text) = True Then
                    Me.TxtPrimaryAccountSales.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountSales.Text)
                    Me.TxtCodeAccountSales.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSales.Text)
                    Me.TxtDescAccountSales.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSales.Text)
                Else
                    If Me.TxtCodeAccountSales.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountSales.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountSales.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountSales.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSales.Text)
                        Me.TxtDescAccountSales.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSales.Text)
                        Me.TxtCodeAccountSalesReturn.Focus()
                    Else
                        Me.TxtPrimaryAccountSales.Text = "-1"
                        Me.TxtDescAccountSales.Clear()
                        Me.TxtCodeAccountSales.Clear()
                        Me.TxtCodeAccountSales.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountSales.Text = "-1"
                Me.TxtDescAccountSales.Clear()
                Me.TxtCodeAccountSales.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountSales.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountSales.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountSales.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountSales.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountSales.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSales.Text)
                Me.TxtDescAccountSales.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSales.Text)
                Me.TxtCodeAccountSalesReturn.Focus()
            Else
                Me.TxtPrimaryAccountSales.Text = "-1"
                Me.TxtDescAccountSales.Clear()
                Me.TxtCodeAccountSales.Clear()
                Me.TxtCodeAccountSales.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    '>>>Sales Account Return
    Private Sub TxtCodeAccountSalesReturn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountSalesReturn.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountSalesReturn.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountSalesReturn.Text) = True Then
                        Me.TxtPrimaryAccountSalesReturn.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountSalesReturn.Text)
                        Me.TxtCodeAccountSalesReturn.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSalesReturn.Text)
                        Me.TxtDescAccountSalesReturn.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSalesReturn.Text)
                        If Me.TxtDescAccountCOGS.Enabled = False Then
                        Else
                            Me.TxtCodeAccountCOGS.Focus()
                        End If
                    Else
                        Me.BtnFindAccountSalesReturn.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountSalesReturn.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountSalesReturn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountSalesReturn.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeAccountSalesReturn.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountSalesReturn.Text) = True Then
                    Me.TxtPrimaryAccountSalesReturn.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountSalesReturn.Text)
                    Me.TxtCodeAccountSalesReturn.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSalesReturn.Text)
                    Me.TxtDescAccountSalesReturn.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSalesReturn.Text)
                Else
                    If Me.TxtCodeAccountSalesReturn.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountSalesReturn.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountSalesReturn.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountSalesReturn.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSalesReturn.Text)
                        Me.TxtDescAccountSalesReturn.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSalesReturn.Text)
                        If Me.TxtDescAccountCOGS.Enabled = False Then
                        Else
                            Me.TxtCodeAccountCOGS.Focus()
                        End If
                    Else
                        Me.TxtPrimaryAccountSalesReturn.Text = "-1"
                        Me.TxtDescAccountSalesReturn.Clear()
                        Me.TxtCodeAccountSalesReturn.Clear()
                        Me.TxtCodeAccountSalesReturn.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountSalesReturn.Text = "-1"
                Me.TxtDescAccountSalesReturn.Clear()
                Me.TxtCodeAccountSalesReturn.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountSalesReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountSalesReturn.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountSalesReturn.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountSalesReturn.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountSalesReturn.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountSalesReturn.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountSalesReturn.Text)
                Me.TxtDescAccountSalesReturn.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountSalesReturn.Text)
                If Me.TxtDescAccountCOGS.Enabled = False Then
                Else
                    Me.TxtCodeAccountCOGS.Focus()
                End If
            Else
                Me.TxtPrimaryAccountSalesReturn.Text = "-1"
                Me.TxtDescAccountSalesReturn.Clear()
                Me.TxtCodeAccountSalesReturn.Clear()
                Me.TxtCodeAccountSalesReturn.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    '>>>Account COGS
    Private Sub TxtCodeAccountCOGS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountCOGS.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountCOGS.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountCOGS.Text) = True Then
                        Me.TxtPrimaryAccountCOGS.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountCOGS.Text)
                        Me.TxtCodeAccountCOGS.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountCOGS.Text)
                        Me.TxtDescAccountCOGS.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountCOGS.Text)
                        Me.TxtCodeAccountAdjustment.Focus()
                    Else
                        Me.BtnFindAccountCOGS.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountCOGS.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountCOGS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountCOGS.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeAccountCOGS.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountCOGS.Text) = True Then
                    Me.TxtPrimaryAccountCOGS.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountCOGS.Text)
                    Me.TxtCodeAccountCOGS.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountCOGS.Text)
                    Me.TxtDescAccountCOGS.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountCOGS.Text)
                Else
                    If Me.TxtCodeAccountCOGS.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountCOGS.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountCOGS.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountCOGS.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountCOGS.Text)
                        Me.TxtDescAccountCOGS.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountCOGS.Text)
                        Me.TxtCodeAccountAdjustment.Focus()
                    Else
                        Me.TxtPrimaryAccountCOGS.Text = "-1"
                        Me.TxtDescAccountCOGS.Clear()
                        Me.TxtCodeAccountCOGS.Clear()
                        Me.TxtCodeAccountCOGS.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountCOGS.Text = "-1"
                Me.TxtDescAccountCOGS.Clear()
                Me.TxtCodeAccountCOGS.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountCOGS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountCOGS.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountCOGS.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountCOGS.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountCOGS.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountCOGS.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountCOGS.Text)
                Me.TxtDescAccountCOGS.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountCOGS.Text)
                Me.TxtCodeAccountAdjustment.Focus()
            Else
                Me.TxtPrimaryAccountCOGS.Text = "-1"
                Me.TxtCodeAccountCOGS.Clear()
                Me.TxtDescAccountCOGS.Clear()
                Me.TxtCodeAccountCOGS.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    '>>>Account Adjustment
    Private Sub TxtCodeAccountAdjustment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountAdjustment.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountAdjustment.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountAdjustment.Text) = True Then
                        Me.TxtPrimaryAccountAdjustment.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountAdjustment.Text)
                        Me.TxtCodeAccountAdjustment.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountAdjustment.Text)
                        Me.TxtDescAccountAdjustment.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountAdjustment.Text)
                        Me.TxtCodeAccountStockTaking.Focus()
                    Else
                        Me.BtnFindAccountAdjustment.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountAdjustment.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountAdjustment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountAdjustment.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeAccountAdjustment.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountAdjustment.Text) = True Then
                    Me.TxtPrimaryAccountAdjustment.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountAdjustment.Text)
                    Me.TxtCodeAccountAdjustment.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountAdjustment.Text)
                    Me.TxtDescAccountAdjustment.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountAdjustment.Text)
                Else
                    If Me.TxtCodeAccountAdjustment.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountAdjustment.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountAdjustment.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountAdjustment.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountAdjustment.Text)
                        Me.TxtDescAccountAdjustment.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountAdjustment.Text)
                        Me.TxtCodeAccountStockTaking.Focus()
                    Else
                        Me.TxtPrimaryAccountAdjustment.Text = "-1"
                        Me.TxtDescAccountAdjustment.Clear()
                        Me.TxtCodeAccountAdjustment.Clear()
                        Me.TxtCodeAccountAdjustment.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountAdjustment.Text = "-1"
                Me.TxtDescAccountAdjustment.Clear()
                Me.TxtCodeAccountAdjustment.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountAdjustment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountAdjustment.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountAdjustment.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountAdjustment.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountAdjustment.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountAdjustment.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountAdjustment.Text)
                Me.TxtDescAccountAdjustment.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountAdjustment.Text)
                Me.TxtCodeAccountStockTaking.Focus()
            Else
                Me.TxtPrimaryAccountAdjustment.Text = "-1"
                Me.TxtDescAccountAdjustment.Clear()
                Me.TxtCodeAccountAdjustment.Clear()
                Me.TxtCodeAccountAdjustment.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    '>>>Account Stock Taking

    Private Sub TxtCodeAccountStockTaking_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAccountStockTaking.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeAccountStockTaking.Text <> "" Then
                    If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountStockTaking.Text) = True Then
                        Me.TxtPrimaryAccountStockTaking.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountStockTaking.Text)
                        Me.TxtCodeAccountStockTaking.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountStockTaking.Text)
                        Me.TxtDescAccountStockTaking.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountStockTaking.Text)
                    Else
                        Me.BtnFindAccountStockTaking.PerformClick()
                    End If
                Else
                    Me.BtnFindAccountStockTaking.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAccountStockTaking_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAccountStockTaking.Leave
        Dim hj As New FrmMasterAkun_Cari1
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeAccountStockTaking.Text <> "" Then
                If Me.objAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCodeAccountStockTaking.Text) = True Then
                    Me.TxtPrimaryAccountStockTaking.Text = Me.objAkun.GetPrimary(Me.TxtCodeAccountStockTaking.Text)
                    Me.TxtCodeAccountStockTaking.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountStockTaking.Text)
                    Me.TxtDescAccountStockTaking.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountStockTaking.Text)
                Else
                    If Me.TxtCodeAccountStockTaking.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeAccountStockTaking.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAccountStockTaking.Text = hj.mPrimaryAkun
                        Me.TxtCodeAccountStockTaking.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountStockTaking.Text)
                        Me.TxtDescAccountStockTaking.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountStockTaking.Text)
                    Else
                        Me.TxtPrimaryAccountStockTaking.Text = "-1"
                        Me.TxtDescAccountStockTaking.Clear()
                        Me.TxtCodeAccountStockTaking.Clear()
                        Me.TxtCodeAccountStockTaking.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAccountStockTaking.Text = "-1"
                Me.TxtDescAccountStockTaking.Clear()
                Me.TxtCodeAccountStockTaking.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAccountStockTaking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAccountStockTaking.Click
        Dim hj As New FrmMasterAkun_Cari1
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtCodeAccountStockTaking.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeAccountStockTaking.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAccountStockTaking.Text = hj.mPrimaryAkun
                Me.TxtCodeAccountStockTaking.Text = Me.objAkun.GetAccountCode(Me.TxtPrimaryAccountStockTaking.Text)
                Me.TxtDescAccountStockTaking.Text = Me.objAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAccountStockTaking.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtPrimaryAccountStockTaking.Text = "-1"
                Me.TxtDescAccountStockTaking.Clear()
                Me.TxtCodeAccountStockTaking.Clear()
                Me.TxtCodeAccountStockTaking.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub
    '--
    Private Sub Reset_Object()
        Me.TxtPrimaryAccountAdjustment.Text = "-1"
        Me.TxtPrimaryAccountCOGS.Text = "-1"
        Me.TxtPrimaryAccountStockTaking.Text = "-1"
    End Sub

    Private Sub CekGoodService()
        '>>>jika goods maka yang aktif semua
        If RboGoods.Checked = True Then
            ModCekData.KendalikanControl(Me.ScAccount.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, True, True, True, True, True, True, True)
        Else
            '>>> jika service yang aktif hanya sebelah kiri
            ModCekData.KendalikanControl(Me.ScAccount.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KosongkanControl(Me.ScAccount.Panel2.Controls, True, False, False)
        End If
    End Sub
    Private Sub pilihType()
        If Me.RboGoods.Checked = True Then
            LblInventory.Text = "Inventory Account"
        Else
            LblInventory.Text = "Expense Account"
        End If

    End Sub
    Private Sub RboGoods_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboGoods.CheckedChanged
        CekGoodService()
        pilihType()
    End Sub
   
    Private Sub CboGroupLevel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboGroupLevel.SelectedIndexChanged
        Me.TxtGroupCode.Text = ""
        Me.TxtGroupDesc.Text = ""
    End Sub
    Private Sub FrmMasterProductAccountGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.btnFind.Select()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ScAccount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ScAccount.GotFocus
        If st = False Then
            btnFind.Focus()
        Else
            btnExit.Focus()
        End If
    End Sub
    Private Sub btnFind_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Leave
        st = True
    End Sub
    Private Sub btnExit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Leave
        st = False
    End Sub

    Private Sub TxtCodeAccountInventory_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCodeAccountInventory.TextChanged

    End Sub

    Public Function CariygSama() As Boolean
        Try
            'Dim _GroupLevel As Double = Me.objMPA.GetGroupLevelByPrimaryGroup(Me.TxtPrimaryGroup.Text)
            'Dim _Type As Integer = Me.objMPA.GetTypeByPrimaryGroup(Me.TxtPrimaryGroup.Text)
            Dim _GroupLevel, _Type As Integer
            If Me.CboGroupLevel.SelectedIndex = 0 Then
                _GroupLevel = 3
            ElseIf Me.CboGroupLevel.SelectedIndex = 1 Then
                _GroupLevel = 2
            ElseIf Me.CboGroupLevel.SelectedIndex = 2 Then
                _GroupLevel = 1
            Else
                _GroupLevel = 0
            End If

            If Me.RboGoods.Checked = True Then
                _Type = 0
            Else
                _Type = 1
            End If

            If Me.objMPA.ApakahAdaPrimary(Me.TxtPrimary.Text, _GroupLevel, Me.TxtPrimaryGroup.Text, _Type) = True Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
