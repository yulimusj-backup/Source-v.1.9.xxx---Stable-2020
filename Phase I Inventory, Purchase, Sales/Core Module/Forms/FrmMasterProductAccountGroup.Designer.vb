<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterProductAccountGroup
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GboDetail1 = New System.Windows.Forms.GroupBox
        Me.TxtGroupDesc = New System.Windows.Forms.TextBox
        Me.BtnFindGroup = New System.Windows.Forms.Button
        Me.TxtPrimaryGroup = New System.Windows.Forms.TextBox
        Me.CboGroupLevel = New System.Windows.Forms.ComboBox
        Me.TxtGroupCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.RboService = New System.Windows.Forms.RadioButton
        Me.RboGoods = New System.Windows.Forms.RadioButton
        Me.GboDetail2 = New System.Windows.Forms.GroupBox
        Me.ScAccount = New System.Windows.Forms.SplitContainer
        Me.TxtDescAccountSalesReturn = New System.Windows.Forms.TextBox
        Me.BtnFindAccountSalesReturn = New System.Windows.Forms.Button
        Me.TxtCodeAccountSalesReturn = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtDescAccountSales = New System.Windows.Forms.TextBox
        Me.BtnFindAccountSales = New System.Windows.Forms.Button
        Me.TxtCodeAccountSales = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtDescAccountInventory = New System.Windows.Forms.TextBox
        Me.BtnFindAccountInventory = New System.Windows.Forms.Button
        Me.TxtCodeAccountInventory = New System.Windows.Forms.TextBox
        Me.LblInventory = New System.Windows.Forms.Label
        Me.TxtPrimaryAccountInventory = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccountSales = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccountSalesReturn = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccountCOGS = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtDescAccountStockTaking = New System.Windows.Forms.TextBox
        Me.BtnFindAccountStockTaking = New System.Windows.Forms.Button
        Me.TxtCodeAccountStockTaking = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtDescAccountAdjustment = New System.Windows.Forms.TextBox
        Me.BtnFindAccountAdjustment = New System.Windows.Forms.Button
        Me.TxtCodeAccountAdjustment = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtDescAccountCOGS = New System.Windows.Forms.TextBox
        Me.BtnFindAccountCOGS = New System.Windows.Forms.Button
        Me.TxtCodeAccountCOGS = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtPrimaryAccountAdjustment = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccountStockTaking = New System.Windows.Forms.TextBox
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.TslCreated_By = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslCreated_Date = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslModified_By = New System.Windows.Forms.ToolStripStatusLabel
        Me.TslModified_Date = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.GboDetail1.SuspendLayout()
        Me.GboDetail2.SuspendLayout()
        Me.ScAccount.Panel1.SuspendLayout()
        Me.ScAccount.Panel2.SuspendLayout()
        Me.ScAccount.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GboDetail1
        '
        Me.GboDetail1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail1.Controls.Add(Me.TxtGroupDesc)
        Me.GboDetail1.Controls.Add(Me.BtnFindGroup)
        Me.GboDetail1.Controls.Add(Me.TxtPrimaryGroup)
        Me.GboDetail1.Controls.Add(Me.CboGroupLevel)
        Me.GboDetail1.Controls.Add(Me.TxtGroupCode)
        Me.GboDetail1.Controls.Add(Me.Label3)
        Me.GboDetail1.Controls.Add(Me.TxtRemarks)
        Me.GboDetail1.Controls.Add(Me.Label11)
        Me.GboDetail1.Controls.Add(Me.Label10)
        Me.GboDetail1.Controls.Add(Me.RboService)
        Me.GboDetail1.Controls.Add(Me.RboGoods)
        Me.GboDetail1.Controls.Add(Me.GboDetail2)
        Me.GboDetail1.Controls.Add(Me.TxtDesc)
        Me.GboDetail1.Controls.Add(Me.Label2)
        Me.GboDetail1.Controls.Add(Me.TxtCode)
        Me.GboDetail1.Controls.Add(Me.Label1)
        Me.GboDetail1.Controls.Add(Me.btnCancel)
        Me.GboDetail1.Controls.Add(Me.btnOK)
        Me.GboDetail1.Controls.Add(Me.TxtPrimary)
        Me.GboDetail1.Location = New System.Drawing.Point(13, 4)
        Me.GboDetail1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GboDetail1.Name = "GboDetail1"
        Me.GboDetail1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GboDetail1.Size = New System.Drawing.Size(896, 430)
        Me.GboDetail1.TabIndex = 0
        Me.GboDetail1.TabStop = False
        '
        'TxtGroupDesc
        '
        Me.TxtGroupDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtGroupDesc.Location = New System.Drawing.Point(604, 74)
        Me.TxtGroupDesc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtGroupDesc.MaxLength = 250
        Me.TxtGroupDesc.Name = "TxtGroupDesc"
        Me.TxtGroupDesc.ReadOnly = True
        Me.TxtGroupDesc.Size = New System.Drawing.Size(279, 22)
        Me.TxtGroupDesc.TabIndex = 5
        Me.TxtGroupDesc.TabStop = False
        '
        'BtnFindGroup
        '
        Me.BtnFindGroup.Location = New System.Drawing.Point(561, 74)
        Me.BtnFindGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindGroup.Name = "BtnFindGroup"
        Me.BtnFindGroup.Size = New System.Drawing.Size(36, 25)
        Me.BtnFindGroup.TabIndex = 4
        Me.BtnFindGroup.TabStop = False
        Me.BtnFindGroup.Text = "..."
        Me.BtnFindGroup.UseVisualStyleBackColor = True
        '
        'TxtPrimaryGroup
        '
        Me.TxtPrimaryGroup.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryGroup.Location = New System.Drawing.Point(155, 75)
        Me.TxtPrimaryGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryGroup.MaxLength = 250
        Me.TxtPrimaryGroup.Name = "TxtPrimaryGroup"
        Me.TxtPrimaryGroup.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryGroup.TabIndex = 33
        Me.TxtPrimaryGroup.TabStop = False
        Me.TxtPrimaryGroup.Text = "-1"
        Me.TxtPrimaryGroup.Visible = False
        '
        'CboGroupLevel
        '
        Me.CboGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroupLevel.FormattingEnabled = True
        Me.CboGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category"})
        Me.CboGroupLevel.Location = New System.Drawing.Point(179, 74)
        Me.CboGroupLevel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboGroupLevel.Name = "CboGroupLevel"
        Me.CboGroupLevel.Size = New System.Drawing.Size(180, 24)
        Me.CboGroupLevel.TabIndex = 2
        '
        'TxtGroupCode
        '
        Me.TxtGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtGroupCode.Location = New System.Drawing.Point(372, 74)
        Me.TxtGroupCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtGroupCode.MaxLength = 250
        Me.TxtGroupCode.Name = "TxtGroupCode"
        Me.TxtGroupCode.Size = New System.Drawing.Size(183, 22)
        Me.TxtGroupCode.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Group level"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(179, 135)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtRemarks.MaxLength = 250
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(704, 22)
        Me.TxtRemarks.TabIndex = 8
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(11, 140)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 17)
        Me.Label11.TabIndex = 29
        Me.Label11.Text = "Remarks"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(11, 112)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 17)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Type"
        '
        'RboService
        '
        Me.RboService.AutoSize = True
        Me.RboService.Location = New System.Drawing.Point(285, 110)
        Me.RboService.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.RboService.Name = "RboService"
        Me.RboService.Size = New System.Drawing.Size(76, 21)
        Me.RboService.TabIndex = 7
        Me.RboService.TabStop = True
        Me.RboService.Text = "Service"
        Me.RboService.UseVisualStyleBackColor = True
        '
        'RboGoods
        '
        Me.RboGoods.AutoSize = True
        Me.RboGoods.Checked = True
        Me.RboGoods.Location = New System.Drawing.Point(179, 110)
        Me.RboGoods.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.RboGoods.Name = "RboGoods"
        Me.RboGoods.Size = New System.Drawing.Size(71, 21)
        Me.RboGoods.TabIndex = 6
        Me.RboGoods.TabStop = True
        Me.RboGoods.Text = "Goods"
        Me.RboGoods.UseVisualStyleBackColor = True
        '
        'GboDetail2
        '
        Me.GboDetail2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GboDetail2.Controls.Add(Me.ScAccount)
        Me.GboDetail2.Location = New System.Drawing.Point(8, 162)
        Me.GboDetail2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GboDetail2.Name = "GboDetail2"
        Me.GboDetail2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GboDetail2.Size = New System.Drawing.Size(885, 227)
        Me.GboDetail2.TabIndex = 9
        Me.GboDetail2.TabStop = False
        Me.GboDetail2.Text = "Accounts"
        '
        'ScAccount
        '
        Me.ScAccount.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ScAccount.Location = New System.Drawing.Point(4, 19)
        Me.ScAccount.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ScAccount.Name = "ScAccount"
        '
        'ScAccount.Panel1
        '
        Me.ScAccount.Panel1.Controls.Add(Me.TxtDescAccountSalesReturn)
        Me.ScAccount.Panel1.Controls.Add(Me.BtnFindAccountSalesReturn)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtCodeAccountSalesReturn)
        Me.ScAccount.Panel1.Controls.Add(Me.Label6)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtDescAccountSales)
        Me.ScAccount.Panel1.Controls.Add(Me.BtnFindAccountSales)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtCodeAccountSales)
        Me.ScAccount.Panel1.Controls.Add(Me.Label5)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtDescAccountInventory)
        Me.ScAccount.Panel1.Controls.Add(Me.BtnFindAccountInventory)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtCodeAccountInventory)
        Me.ScAccount.Panel1.Controls.Add(Me.LblInventory)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtPrimaryAccountInventory)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtPrimaryAccountSales)
        Me.ScAccount.Panel1.Controls.Add(Me.TxtPrimaryAccountSalesReturn)
        '
        'ScAccount.Panel2
        '
        Me.ScAccount.Panel2.Controls.Add(Me.TxtPrimaryAccountCOGS)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtReadOnly)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtDescAccountStockTaking)
        Me.ScAccount.Panel2.Controls.Add(Me.BtnFindAccountStockTaking)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtCodeAccountStockTaking)
        Me.ScAccount.Panel2.Controls.Add(Me.Label7)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtDescAccountAdjustment)
        Me.ScAccount.Panel2.Controls.Add(Me.BtnFindAccountAdjustment)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtCodeAccountAdjustment)
        Me.ScAccount.Panel2.Controls.Add(Me.Label8)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtDescAccountCOGS)
        Me.ScAccount.Panel2.Controls.Add(Me.BtnFindAccountCOGS)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtCodeAccountCOGS)
        Me.ScAccount.Panel2.Controls.Add(Me.Label9)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtPrimaryAccountAdjustment)
        Me.ScAccount.Panel2.Controls.Add(Me.TxtPrimaryAccountStockTaking)
        Me.ScAccount.Size = New System.Drawing.Size(877, 204)
        Me.ScAccount.SplitterDistance = 435
        Me.ScAccount.SplitterWidth = 5
        Me.ScAccount.TabIndex = 5
        '
        'TxtDescAccountSalesReturn
        '
        Me.TxtDescAccountSalesReturn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountSalesReturn.Location = New System.Drawing.Point(167, 171)
        Me.TxtDescAccountSalesReturn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountSalesReturn.Name = "TxtDescAccountSalesReturn"
        Me.TxtDescAccountSalesReturn.ReadOnly = True
        Me.TxtDescAccountSalesReturn.Size = New System.Drawing.Size(263, 22)
        Me.TxtDescAccountSalesReturn.TabIndex = 8
        Me.TxtDescAccountSalesReturn.TabStop = False
        '
        'BtnFindAccountSalesReturn
        '
        Me.BtnFindAccountSalesReturn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountSalesReturn.Location = New System.Drawing.Point(375, 139)
        Me.BtnFindAccountSalesReturn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountSalesReturn.Name = "BtnFindAccountSalesReturn"
        Me.BtnFindAccountSalesReturn.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountSalesReturn.TabIndex = 7
        Me.BtnFindAccountSalesReturn.TabStop = False
        Me.BtnFindAccountSalesReturn.Text = "Find"
        Me.BtnFindAccountSalesReturn.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountSalesReturn
        '
        Me.TxtCodeAccountSalesReturn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountSalesReturn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountSalesReturn.Location = New System.Drawing.Point(167, 139)
        Me.TxtCodeAccountSalesReturn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountSalesReturn.MaxLength = 250
        Me.TxtCodeAccountSalesReturn.Name = "TxtCodeAccountSalesReturn"
        Me.TxtCodeAccountSalesReturn.Size = New System.Drawing.Size(203, 22)
        Me.TxtCodeAccountSalesReturn.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 144)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(145, 17)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Sales Return Account"
        '
        'TxtDescAccountSales
        '
        Me.TxtDescAccountSales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountSales.Location = New System.Drawing.Point(167, 101)
        Me.TxtDescAccountSales.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountSales.Name = "TxtDescAccountSales"
        Me.TxtDescAccountSales.ReadOnly = True
        Me.TxtDescAccountSales.Size = New System.Drawing.Size(263, 22)
        Me.TxtDescAccountSales.TabIndex = 5
        Me.TxtDescAccountSales.TabStop = False
        '
        'BtnFindAccountSales
        '
        Me.BtnFindAccountSales.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountSales.Location = New System.Drawing.Point(375, 69)
        Me.BtnFindAccountSales.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountSales.Name = "BtnFindAccountSales"
        Me.BtnFindAccountSales.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountSales.TabIndex = 4
        Me.BtnFindAccountSales.TabStop = False
        Me.BtnFindAccountSales.Text = "Find"
        Me.BtnFindAccountSales.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountSales
        '
        Me.TxtCodeAccountSales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountSales.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountSales.Location = New System.Drawing.Point(167, 69)
        Me.TxtCodeAccountSales.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountSales.MaxLength = 250
        Me.TxtCodeAccountSales.Name = "TxtCodeAccountSales"
        Me.TxtCodeAccountSales.Size = New System.Drawing.Size(203, 22)
        Me.TxtCodeAccountSales.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 74)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 17)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Sales Account"
        '
        'TxtDescAccountInventory
        '
        Me.TxtDescAccountInventory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountInventory.Location = New System.Drawing.Point(167, 32)
        Me.TxtDescAccountInventory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountInventory.Name = "TxtDescAccountInventory"
        Me.TxtDescAccountInventory.ReadOnly = True
        Me.TxtDescAccountInventory.Size = New System.Drawing.Size(263, 22)
        Me.TxtDescAccountInventory.TabIndex = 2
        Me.TxtDescAccountInventory.TabStop = False
        '
        'BtnFindAccountInventory
        '
        Me.BtnFindAccountInventory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountInventory.Location = New System.Drawing.Point(375, 4)
        Me.BtnFindAccountInventory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountInventory.Name = "BtnFindAccountInventory"
        Me.BtnFindAccountInventory.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountInventory.TabIndex = 1
        Me.BtnFindAccountInventory.TabStop = False
        Me.BtnFindAccountInventory.Text = "Find"
        Me.BtnFindAccountInventory.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountInventory
        '
        Me.TxtCodeAccountInventory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountInventory.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountInventory.Location = New System.Drawing.Point(167, 4)
        Me.TxtCodeAccountInventory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountInventory.MaxLength = 250
        Me.TxtCodeAccountInventory.Name = "TxtCodeAccountInventory"
        Me.TxtCodeAccountInventory.Size = New System.Drawing.Size(203, 22)
        Me.TxtCodeAccountInventory.TabIndex = 0
        '
        'LblInventory
        '
        Me.LblInventory.AutoSize = True
        Me.LblInventory.Location = New System.Drawing.Point(13, 9)
        Me.LblInventory.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblInventory.Name = "LblInventory"
        Me.LblInventory.Size = New System.Drawing.Size(121, 17)
        Me.LblInventory.TabIndex = 12
        Me.LblInventory.Text = "Inventory Account"
        '
        'TxtPrimaryAccountInventory
        '
        Me.TxtPrimaryAccountInventory.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountInventory.Location = New System.Drawing.Point(120, 4)
        Me.TxtPrimaryAccountInventory.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountInventory.MaxLength = 250
        Me.TxtPrimaryAccountInventory.Name = "TxtPrimaryAccountInventory"
        Me.TxtPrimaryAccountInventory.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountInventory.TabIndex = 25
        Me.TxtPrimaryAccountInventory.TabStop = False
        Me.TxtPrimaryAccountInventory.Text = "-1"
        Me.TxtPrimaryAccountInventory.Visible = False
        '
        'TxtPrimaryAccountSales
        '
        Me.TxtPrimaryAccountSales.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountSales.Location = New System.Drawing.Point(120, 69)
        Me.TxtPrimaryAccountSales.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountSales.MaxLength = 250
        Me.TxtPrimaryAccountSales.Name = "TxtPrimaryAccountSales"
        Me.TxtPrimaryAccountSales.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountSales.TabIndex = 26
        Me.TxtPrimaryAccountSales.TabStop = False
        Me.TxtPrimaryAccountSales.Text = "-1"
        Me.TxtPrimaryAccountSales.Visible = False
        '
        'TxtPrimaryAccountSalesReturn
        '
        Me.TxtPrimaryAccountSalesReturn.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountSalesReturn.Location = New System.Drawing.Point(120, 139)
        Me.TxtPrimaryAccountSalesReturn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountSalesReturn.MaxLength = 250
        Me.TxtPrimaryAccountSalesReturn.Name = "TxtPrimaryAccountSalesReturn"
        Me.TxtPrimaryAccountSalesReturn.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountSalesReturn.TabIndex = 27
        Me.TxtPrimaryAccountSalesReturn.TabStop = False
        Me.TxtPrimaryAccountSalesReturn.Text = "-1"
        Me.TxtPrimaryAccountSalesReturn.Visible = False
        '
        'TxtPrimaryAccountCOGS
        '
        Me.TxtPrimaryAccountCOGS.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountCOGS.Location = New System.Drawing.Point(113, 4)
        Me.TxtPrimaryAccountCOGS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountCOGS.MaxLength = 250
        Me.TxtPrimaryAccountCOGS.Name = "TxtPrimaryAccountCOGS"
        Me.TxtPrimaryAccountCOGS.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountCOGS.TabIndex = 28
        Me.TxtPrimaryAccountCOGS.TabStop = False
        Me.TxtPrimaryAccountCOGS.Text = "-1"
        Me.TxtPrimaryAccountCOGS.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(411, 220)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(18, 22)
        Me.TxtReadOnly.TabIndex = 2
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'TxtDescAccountStockTaking
        '
        Me.TxtDescAccountStockTaking.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountStockTaking.Location = New System.Drawing.Point(156, 171)
        Me.TxtDescAccountStockTaking.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountStockTaking.Name = "TxtDescAccountStockTaking"
        Me.TxtDescAccountStockTaking.ReadOnly = True
        Me.TxtDescAccountStockTaking.Size = New System.Drawing.Size(265, 22)
        Me.TxtDescAccountStockTaking.TabIndex = 8
        Me.TxtDescAccountStockTaking.TabStop = False
        '
        'BtnFindAccountStockTaking
        '
        Me.BtnFindAccountStockTaking.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountStockTaking.Location = New System.Drawing.Point(366, 139)
        Me.BtnFindAccountStockTaking.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountStockTaking.Name = "BtnFindAccountStockTaking"
        Me.BtnFindAccountStockTaking.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountStockTaking.TabIndex = 7
        Me.BtnFindAccountStockTaking.TabStop = False
        Me.BtnFindAccountStockTaking.Text = "Find"
        Me.BtnFindAccountStockTaking.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountStockTaking
        '
        Me.TxtCodeAccountStockTaking.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountStockTaking.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountStockTaking.Location = New System.Drawing.Point(156, 139)
        Me.TxtCodeAccountStockTaking.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountStockTaking.MaxLength = 250
        Me.TxtCodeAccountStockTaking.Name = "TxtCodeAccountStockTaking"
        Me.TxtCodeAccountStockTaking.Size = New System.Drawing.Size(205, 22)
        Me.TxtCodeAccountStockTaking.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 144)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(145, 17)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "Stock Taking Account"
        '
        'TxtDescAccountAdjustment
        '
        Me.TxtDescAccountAdjustment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountAdjustment.Location = New System.Drawing.Point(156, 101)
        Me.TxtDescAccountAdjustment.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountAdjustment.Name = "TxtDescAccountAdjustment"
        Me.TxtDescAccountAdjustment.ReadOnly = True
        Me.TxtDescAccountAdjustment.Size = New System.Drawing.Size(265, 22)
        Me.TxtDescAccountAdjustment.TabIndex = 5
        Me.TxtDescAccountAdjustment.TabStop = False
        '
        'BtnFindAccountAdjustment
        '
        Me.BtnFindAccountAdjustment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountAdjustment.Location = New System.Drawing.Point(366, 69)
        Me.BtnFindAccountAdjustment.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountAdjustment.Name = "BtnFindAccountAdjustment"
        Me.BtnFindAccountAdjustment.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountAdjustment.TabIndex = 4
        Me.BtnFindAccountAdjustment.TabStop = False
        Me.BtnFindAccountAdjustment.Text = "Find"
        Me.BtnFindAccountAdjustment.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountAdjustment
        '
        Me.TxtCodeAccountAdjustment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountAdjustment.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountAdjustment.Location = New System.Drawing.Point(156, 69)
        Me.TxtCodeAccountAdjustment.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountAdjustment.MaxLength = 250
        Me.TxtCodeAccountAdjustment.Name = "TxtCodeAccountAdjustment"
        Me.TxtCodeAccountAdjustment.Size = New System.Drawing.Size(205, 22)
        Me.TxtCodeAccountAdjustment.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 74)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(133, 17)
        Me.Label8.TabIndex = 28
        Me.Label8.Text = "Adjustment Account"
        '
        'TxtDescAccountCOGS
        '
        Me.TxtDescAccountCOGS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescAccountCOGS.Location = New System.Drawing.Point(156, 32)
        Me.TxtDescAccountCOGS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDescAccountCOGS.Name = "TxtDescAccountCOGS"
        Me.TxtDescAccountCOGS.ReadOnly = True
        Me.TxtDescAccountCOGS.Size = New System.Drawing.Size(265, 22)
        Me.TxtDescAccountCOGS.TabIndex = 2
        Me.TxtDescAccountCOGS.TabStop = False
        '
        'BtnFindAccountCOGS
        '
        Me.BtnFindAccountCOGS.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindAccountCOGS.Location = New System.Drawing.Point(366, 4)
        Me.BtnFindAccountCOGS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFindAccountCOGS.Name = "BtnFindAccountCOGS"
        Me.BtnFindAccountCOGS.Size = New System.Drawing.Size(56, 25)
        Me.BtnFindAccountCOGS.TabIndex = 1
        Me.BtnFindAccountCOGS.TabStop = False
        Me.BtnFindAccountCOGS.Text = "Find"
        Me.BtnFindAccountCOGS.UseVisualStyleBackColor = True
        '
        'TxtCodeAccountCOGS
        '
        Me.TxtCodeAccountCOGS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCodeAccountCOGS.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeAccountCOGS.Location = New System.Drawing.Point(156, 4)
        Me.TxtCodeAccountCOGS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCodeAccountCOGS.MaxLength = 250
        Me.TxtCodeAccountCOGS.Name = "TxtCodeAccountCOGS"
        Me.TxtCodeAccountCOGS.Size = New System.Drawing.Size(205, 22)
        Me.TxtCodeAccountCOGS.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 9)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(146, 17)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Cost of Sales Account"
        '
        'TxtPrimaryAccountAdjustment
        '
        Me.TxtPrimaryAccountAdjustment.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountAdjustment.Location = New System.Drawing.Point(113, 69)
        Me.TxtPrimaryAccountAdjustment.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountAdjustment.MaxLength = 250
        Me.TxtPrimaryAccountAdjustment.Name = "TxtPrimaryAccountAdjustment"
        Me.TxtPrimaryAccountAdjustment.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountAdjustment.TabIndex = 36
        Me.TxtPrimaryAccountAdjustment.TabStop = False
        Me.TxtPrimaryAccountAdjustment.Text = "-1"
        Me.TxtPrimaryAccountAdjustment.Visible = False
        '
        'TxtPrimaryAccountStockTaking
        '
        Me.TxtPrimaryAccountStockTaking.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryAccountStockTaking.Location = New System.Drawing.Point(113, 139)
        Me.TxtPrimaryAccountStockTaking.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimaryAccountStockTaking.MaxLength = 250
        Me.TxtPrimaryAccountStockTaking.Name = "TxtPrimaryAccountStockTaking"
        Me.TxtPrimaryAccountStockTaking.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimaryAccountStockTaking.TabIndex = 37
        Me.TxtPrimaryAccountStockTaking.TabStop = False
        Me.TxtPrimaryAccountStockTaking.Text = "-1"
        Me.TxtPrimaryAccountStockTaking.Visible = False
        '
        'TxtDesc
        '
        Me.TxtDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDesc.Location = New System.Drawing.Point(179, 44)
        Me.TxtDesc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtDesc.MaxLength = 250
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(704, 22)
        Me.TxtDesc.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 49)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Group Desc"
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(179, 16)
        Me.TxtCode.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtCode.MaxLength = 250
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.Size = New System.Drawing.Size(704, 22)
        Me.TxtCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 21)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Group Code"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(451, 391)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 32)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(337, 391)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 32)
        Me.btnOK.TabIndex = 10
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrimary
        '
        Me.TxtPrimary.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimary.Location = New System.Drawing.Point(155, 16)
        Me.TxtPrimary.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtPrimary.MaxLength = 250
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(37, 22)
        Me.TxtPrimary.TabIndex = 24
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TslCreated_By, Me.TslCreated_Date, Me.TslModified_By, Me.TslModified_Date})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 483)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(923, 25)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TslCreated_By
        '
        Me.TslCreated_By.Name = "TslCreated_By"
        Me.TslCreated_By.Size = New System.Drawing.Size(225, 20)
        Me.TslCreated_By.Spring = True
        Me.TslCreated_By.Text = "Created By: "
        '
        'TslCreated_Date
        '
        Me.TslCreated_Date.Name = "TslCreated_Date"
        Me.TslCreated_Date.Size = New System.Drawing.Size(225, 20)
        Me.TslCreated_Date.Spring = True
        Me.TslCreated_Date.Text = "Created Date:"
        '
        'TslModified_By
        '
        Me.TslModified_By.Name = "TslModified_By"
        Me.TslModified_By.Size = New System.Drawing.Size(225, 20)
        Me.TslModified_By.Spring = True
        Me.TslModified_By.Text = "Modified By: "
        '
        'TslModified_Date
        '
        Me.TslModified_Date.Name = "TslModified_Date"
        Me.TslModified_Date.Size = New System.Drawing.Size(225, 20)
        Me.TslModified_Date.Spring = True
        Me.TslModified_Date.Text = "Modified Date: "
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(522, 441)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(637, 441)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(293, 441)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(407, 441)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 3
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(178, 441)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'FrmMasterProductAccountGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(923, 508)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GboDetail1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(931, 545)
        Me.Name = "FrmMasterProductAccountGroup"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Account Group"
        Me.GboDetail1.ResumeLayout(False)
        Me.GboDetail1.PerformLayout()
        Me.GboDetail2.ResumeLayout(False)
        Me.ScAccount.Panel1.ResumeLayout(False)
        Me.ScAccount.Panel1.PerformLayout()
        Me.ScAccount.Panel2.ResumeLayout(False)
        Me.ScAccount.Panel2.PerformLayout()
        Me.ScAccount.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GboDetail1 As System.Windows.Forms.GroupBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TslCreated_By As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslCreated_Date As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslModified_By As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TslModified_Date As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GboDetail2 As System.Windows.Forms.GroupBox
    Friend WithEvents ScAccount As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtDescAccountInventory As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountInventory As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountInventory As System.Windows.Forms.TextBox
    Friend WithEvents LblInventory As System.Windows.Forms.Label
    Friend WithEvents TxtDescAccountSalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountSalesReturn As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountSalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtDescAccountSales As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountSales As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountSales As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtDescAccountStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountStockTaking As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtDescAccountAdjustment As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountAdjustment As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountAdjustment As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtDescAccountCOGS As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindAccountCOGS As System.Windows.Forms.Button
    Friend WithEvents TxtCodeAccountCOGS As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryAccountInventory As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccountSales As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccountSalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccountCOGS As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccountAdjustment As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccountStockTaking As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents RboService As System.Windows.Forms.RadioButton
    Friend WithEvents RboGoods As System.Windows.Forms.RadioButton
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPrimaryGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtGroupDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindGroup As System.Windows.Forms.Button
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button

End Class
