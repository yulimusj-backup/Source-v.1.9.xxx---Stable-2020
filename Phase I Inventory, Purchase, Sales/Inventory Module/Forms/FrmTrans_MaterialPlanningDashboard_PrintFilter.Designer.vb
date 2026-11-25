<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_PrintFilter
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
        Me.chkRequiredOnly = New System.Windows.Forms.CheckBox
        Me.dtpMP_Filter = New System.Windows.Forms.DateTimePicker
        Me.chkMP_RequiredOnly = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.TxtFilter_PrimaryProductGroup = New System.Windows.Forms.TextBox
        Me.btnFilter_FindProductGroupLevel = New System.Windows.Forms.Button
        Me.TxtFilter_ProductGroupCode = New System.Windows.Forms.TextBox
        Me.LblFilter_PoductGroupDesc = New System.Windows.Forms.Label
        Me.cboFilter_ProductGroupLevel = New System.Windows.Forms.ComboBox
        Me.LblFilter_ProductGroup = New System.Windows.Forms.Label
        Me.TxtFilter_ProductGroupLevelDesc = New System.Windows.Forms.TextBox
        Me.chkFollowUpOnly = New System.Windows.Forms.CheckBox
        Me.chkSOBOnly = New System.Windows.Forms.CheckBox
        Me.chkPrOBOnly = New System.Windows.Forms.CheckBox
        Me.cboProductDetail = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboGrouping = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkShowLegend = New System.Windows.Forms.CheckBox
        Me.chkFilter_Service = New System.Windows.Forms.CheckBox
        Me.chkFilter_Goods = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'chkRequiredOnly
        '
        Me.chkRequiredOnly.AutoSize = True
        Me.chkRequiredOnly.Location = New System.Drawing.Point(13, 162)
        Me.chkRequiredOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkRequiredOnly.Name = "chkRequiredOnly"
        Me.chkRequiredOnly.Size = New System.Drawing.Size(121, 21)
        Me.chkRequiredOnly.TabIndex = 0
        Me.chkRequiredOnly.Text = "Required Only"
        Me.chkRequiredOnly.UseVisualStyleBackColor = True
        '
        'dtpMP_Filter
        '
        Me.dtpMP_Filter.CustomFormat = "dd MMM yyyy"
        Me.dtpMP_Filter.Enabled = False
        Me.dtpMP_Filter.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP_Filter.Location = New System.Drawing.Point(160, 135)
        Me.dtpMP_Filter.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpMP_Filter.Name = "dtpMP_Filter"
        Me.dtpMP_Filter.Size = New System.Drawing.Size(132, 22)
        Me.dtpMP_Filter.TabIndex = 79
        '
        'chkMP_RequiredOnly
        '
        Me.chkMP_RequiredOnly.AutoSize = True
        Me.chkMP_RequiredOnly.Location = New System.Drawing.Point(13, 135)
        Me.chkMP_RequiredOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.chkMP_RequiredOnly.Name = "chkMP_RequiredOnly"
        Me.chkMP_RequiredOnly.Size = New System.Drawing.Size(103, 21)
        Me.chkMP_RequiredOnly.TabIndex = 78
        Me.chkMP_RequiredOnly.Text = "Order Point"
        Me.chkMP_RequiredOnly.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(12, 311)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 80
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(93, 311)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 81
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'TxtFilter_PrimaryProductGroup
        '
        Me.TxtFilter_PrimaryProductGroup.Location = New System.Drawing.Point(100, 251)
        Me.TxtFilter_PrimaryProductGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_PrimaryProductGroup.Name = "TxtFilter_PrimaryProductGroup"
        Me.TxtFilter_PrimaryProductGroup.Size = New System.Drawing.Size(28, 22)
        Me.TxtFilter_PrimaryProductGroup.TabIndex = 93
        Me.TxtFilter_PrimaryProductGroup.TabStop = False
        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
        Me.TxtFilter_PrimaryProductGroup.Visible = False
        '
        'btnFilter_FindProductGroupLevel
        '
        Me.btnFilter_FindProductGroupLevel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindProductGroupLevel.Location = New System.Drawing.Point(339, 246)
        Me.btnFilter_FindProductGroupLevel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFilter_FindProductGroupLevel.Name = "btnFilter_FindProductGroupLevel"
        Me.btnFilter_FindProductGroupLevel.Size = New System.Drawing.Size(52, 28)
        Me.btnFilter_FindProductGroupLevel.TabIndex = 92
        Me.btnFilter_FindProductGroupLevel.Text = "Find"
        Me.btnFilter_FindProductGroupLevel.UseVisualStyleBackColor = True
        '
        'TxtFilter_ProductGroupCode
        '
        Me.TxtFilter_ProductGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_ProductGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_ProductGroupCode.Location = New System.Drawing.Point(160, 249)
        Me.TxtFilter_ProductGroupCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_ProductGroupCode.Name = "TxtFilter_ProductGroupCode"
        Me.TxtFilter_ProductGroupCode.Size = New System.Drawing.Size(170, 22)
        Me.TxtFilter_ProductGroupCode.TabIndex = 91
        '
        'LblFilter_PoductGroupDesc
        '
        Me.LblFilter_PoductGroupDesc.AutoSize = True
        Me.LblFilter_PoductGroupDesc.Location = New System.Drawing.Point(13, 253)
        Me.LblFilter_PoductGroupDesc.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblFilter_PoductGroupDesc.Name = "LblFilter_PoductGroupDesc"
        Me.LblFilter_PoductGroupDesc.Size = New System.Drawing.Size(101, 17)
        Me.LblFilter_PoductGroupDesc.TabIndex = 90
        Me.LblFilter_PoductGroupDesc.Text = "Product Group"
        '
        'cboFilter_ProductGroupLevel
        '
        Me.cboFilter_ProductGroupLevel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFilter_ProductGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilter_ProductGroupLevel.FormattingEnabled = True
        Me.cboFilter_ProductGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category"})
        Me.cboFilter_ProductGroupLevel.Location = New System.Drawing.Point(160, 217)
        Me.cboFilter_ProductGroupLevel.Margin = New System.Windows.Forms.Padding(4)
        Me.cboFilter_ProductGroupLevel.Name = "cboFilter_ProductGroupLevel"
        Me.cboFilter_ProductGroupLevel.Size = New System.Drawing.Size(231, 24)
        Me.cboFilter_ProductGroupLevel.TabIndex = 89
        '
        'LblFilter_ProductGroup
        '
        Me.LblFilter_ProductGroup.AutoSize = True
        Me.LblFilter_ProductGroup.Location = New System.Drawing.Point(13, 221)
        Me.LblFilter_ProductGroup.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblFilter_ProductGroup.Name = "LblFilter_ProductGroup"
        Me.LblFilter_ProductGroup.Size = New System.Drawing.Size(139, 17)
        Me.LblFilter_ProductGroup.TabIndex = 88
        Me.LblFilter_ProductGroup.Text = "Product Group Level"
        '
        'TxtFilter_ProductGroupLevelDesc
        '
        Me.TxtFilter_ProductGroupLevelDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_ProductGroupLevelDesc.Location = New System.Drawing.Point(160, 283)
        Me.TxtFilter_ProductGroupLevelDesc.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFilter_ProductGroupLevelDesc.Name = "TxtFilter_ProductGroupLevelDesc"
        Me.TxtFilter_ProductGroupLevelDesc.ReadOnly = True
        Me.TxtFilter_ProductGroupLevelDesc.Size = New System.Drawing.Size(230, 22)
        Me.TxtFilter_ProductGroupLevelDesc.TabIndex = 94
        '
        'chkFollowUpOnly
        '
        Me.chkFollowUpOnly.AutoSize = True
        Me.chkFollowUpOnly.Location = New System.Drawing.Point(13, 190)
        Me.chkFollowUpOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkFollowUpOnly.Name = "chkFollowUpOnly"
        Me.chkFollowUpOnly.Size = New System.Drawing.Size(124, 21)
        Me.chkFollowUpOnly.TabIndex = 95
        Me.chkFollowUpOnly.Text = "Follow Up Only"
        Me.chkFollowUpOnly.UseVisualStyleBackColor = True
        '
        'chkSOBOnly
        '
        Me.chkSOBOnly.AutoSize = True
        Me.chkSOBOnly.Location = New System.Drawing.Point(160, 162)
        Me.chkSOBOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkSOBOnly.Name = "chkSOBOnly"
        Me.chkSOBOnly.Size = New System.Drawing.Size(208, 21)
        Me.chkSOBOnly.TabIndex = 96
        Me.chkSOBOnly.Text = "Item w/ Sales Order Balance"
        Me.chkSOBOnly.UseVisualStyleBackColor = True
        '
        'chkPrOBOnly
        '
        Me.chkPrOBOnly.AutoSize = True
        Me.chkPrOBOnly.Location = New System.Drawing.Point(160, 190)
        Me.chkPrOBOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkPrOBOnly.Name = "chkPrOBOnly"
        Me.chkPrOBOnly.Size = New System.Drawing.Size(230, 21)
        Me.chkPrOBOnly.TabIndex = 97
        Me.chkPrOBOnly.Text = "Item w/ Purch. Request Balance"
        Me.chkPrOBOnly.UseVisualStyleBackColor = True
        '
        'cboProductDetail
        '
        Me.cboProductDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboProductDetail.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProductDetail.FormattingEnabled = True
        Me.cboProductDetail.Items.AddRange(New Object() {"Product Code", "Product Description", "Product Code & Description"})
        Me.cboProductDetail.Location = New System.Drawing.Point(160, 13)
        Me.cboProductDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.cboProductDetail.Name = "cboProductDetail"
        Me.cboProductDetail.Size = New System.Drawing.Size(231, 24)
        Me.cboProductDetail.TabIndex = 99
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 17)
        Me.Label1.TabIndex = 98
        Me.Label1.Text = "Product Detail"
        '
        'cboGrouping
        '
        Me.cboGrouping.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboGrouping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrouping.FormattingEnabled = True
        Me.cboGrouping.Items.AddRange(New Object() {"No Grouping", "Group By Product Group", "Group By Brand", "Group By Sub Brand", "Group By Category"})
        Me.cboGrouping.Location = New System.Drawing.Point(160, 45)
        Me.cboGrouping.Margin = New System.Windows.Forms.Padding(4)
        Me.cboGrouping.Name = "cboGrouping"
        Me.cboGrouping.Size = New System.Drawing.Size(231, 24)
        Me.cboGrouping.TabIndex = 101
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 49)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 17)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Grouping"
        '
        'chkShowLegend
        '
        Me.chkShowLegend.AutoSize = True
        Me.chkShowLegend.Checked = True
        Me.chkShowLegend.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLegend.Location = New System.Drawing.Point(13, 106)
        Me.chkShowLegend.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowLegend.Name = "chkShowLegend"
        Me.chkShowLegend.Size = New System.Drawing.Size(116, 21)
        Me.chkShowLegend.TabIndex = 102
        Me.chkShowLegend.Text = "Show Legend"
        Me.chkShowLegend.UseVisualStyleBackColor = True
        '
        'chkFilter_Service
        '
        Me.chkFilter_Service.AutoSize = True
        Me.chkFilter_Service.Checked = True
        Me.chkFilter_Service.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_Service.Location = New System.Drawing.Point(160, 77)
        Me.chkFilter_Service.Margin = New System.Windows.Forms.Padding(4)
        Me.chkFilter_Service.Name = "chkFilter_Service"
        Me.chkFilter_Service.Size = New System.Drawing.Size(77, 21)
        Me.chkFilter_Service.TabIndex = 104
        Me.chkFilter_Service.Text = "Service"
        Me.chkFilter_Service.UseVisualStyleBackColor = True
        '
        'chkFilter_Goods
        '
        Me.chkFilter_Goods.AutoSize = True
        Me.chkFilter_Goods.Checked = True
        Me.chkFilter_Goods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_Goods.Location = New System.Drawing.Point(13, 77)
        Me.chkFilter_Goods.Margin = New System.Windows.Forms.Padding(4)
        Me.chkFilter_Goods.Name = "chkFilter_Goods"
        Me.chkFilter_Goods.Size = New System.Drawing.Size(72, 21)
        Me.chkFilter_Goods.TabIndex = 103
        Me.chkFilter_Goods.Text = "Goods"
        Me.chkFilter_Goods.UseVisualStyleBackColor = True
        '
        'FrmTrans_MaterialPlanningDashboard_PrintFilter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(399, 349)
        Me.Controls.Add(Me.chkFilter_Service)
        Me.Controls.Add(Me.chkFilter_Goods)
        Me.Controls.Add(Me.chkShowLegend)
        Me.Controls.Add(Me.cboGrouping)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboProductDetail)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnFilter_FindProductGroupLevel)
        Me.Controls.Add(Me.chkPrOBOnly)
        Me.Controls.Add(Me.chkSOBOnly)
        Me.Controls.Add(Me.chkFollowUpOnly)
        Me.Controls.Add(Me.TxtFilter_ProductGroupLevelDesc)
        Me.Controls.Add(Me.TxtFilter_PrimaryProductGroup)
        Me.Controls.Add(Me.TxtFilter_ProductGroupCode)
        Me.Controls.Add(Me.LblFilter_PoductGroupDesc)
        Me.Controls.Add(Me.cboFilter_ProductGroupLevel)
        Me.Controls.Add(Me.LblFilter_ProductGroup)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dtpMP_Filter)
        Me.Controls.Add(Me.chkMP_RequiredOnly)
        Me.Controls.Add(Me.chkRequiredOnly)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(405, 384)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_PrintFilter"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Filter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkRequiredOnly As System.Windows.Forms.CheckBox
    Friend WithEvents dtpMP_Filter As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkMP_RequiredOnly As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_PrimaryProductGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnFilter_FindProductGroupLevel As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_ProductGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_PoductGroupDesc As System.Windows.Forms.Label
    Friend WithEvents cboFilter_ProductGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents LblFilter_ProductGroup As System.Windows.Forms.Label
    Friend WithEvents TxtFilter_ProductGroupLevelDesc As System.Windows.Forms.TextBox
    Friend WithEvents chkFollowUpOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkSOBOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrOBOnly As System.Windows.Forms.CheckBox
    Friend WithEvents cboProductDetail As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboGrouping As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkShowLegend As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_Service As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_Goods As System.Windows.Forms.CheckBox
End Class
