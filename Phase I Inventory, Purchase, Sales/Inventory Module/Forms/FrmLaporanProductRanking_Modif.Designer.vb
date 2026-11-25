<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProductRanking_Modif
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
        Me.Label8 = New System.Windows.Forms.Label
        Me.cboPeriod = New System.Windows.Forms.ComboBox
        Me.chkCost = New System.Windows.Forms.CheckBox
        Me.chkAbsoluteData = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.txtCategoryDesc = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtBrandDesc = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboSort = New System.Windows.Forms.ComboBox
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.btnFindCategory = New System.Windows.Forms.Button
        Me.txtCategory = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtPrimaryParentGroup = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.ChkGroupByPeriod = New System.Windows.Forms.CheckBox
        Me.btnFindParentGroup = New System.Windows.Forms.Button
        Me.txtParentGroupCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.RBNoGroup = New System.Windows.Forms.RadioButton
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.RBGroup = New System.Windows.Forms.RadioButton
        Me.btnFilter = New System.Windows.Forms.Button
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(13, 49)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(62, 13)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Period Filter"
        '
        'cboPeriod
        '
        Me.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriod.FormattingEnabled = True
        Me.cboPeriod.Items.AddRange(New Object() {"All ", "6 Months", "12 Months", "18 Months", "24 Months", ">36 Months"})
        Me.cboPeriod.Location = New System.Drawing.Point(81, 46)
        Me.cboPeriod.Name = "cboPeriod"
        Me.cboPeriod.Size = New System.Drawing.Size(179, 21)
        Me.cboPeriod.TabIndex = 25
        '
        'chkCost
        '
        Me.chkCost.AutoSize = True
        Me.chkCost.Checked = True
        Me.chkCost.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCost.Location = New System.Drawing.Point(6, 90)
        Me.chkCost.Name = "chkCost"
        Me.chkCost.Size = New System.Drawing.Size(77, 17)
        Me.chkCost.TabIndex = 24
        Me.chkCost.TabStop = False
        Me.chkCost.Text = "Show Cost"
        Me.chkCost.UseVisualStyleBackColor = True
        Me.chkCost.Visible = False
        '
        'chkAbsoluteData
        '
        Me.chkAbsoluteData.AutoSize = True
        Me.chkAbsoluteData.Checked = True
        Me.chkAbsoluteData.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAbsoluteData.Location = New System.Drawing.Point(6, 67)
        Me.chkAbsoluteData.Name = "chkAbsoluteData"
        Me.chkAbsoluteData.Size = New System.Drawing.Size(115, 17)
        Me.chkAbsoluteData.TabIndex = 23
        Me.chkAbsoluteData.TabStop = False
        Me.chkAbsoluteData.Text = "Non Moving Stock"
        Me.chkAbsoluteData.UseVisualStyleBackColor = True
        Me.chkAbsoluteData.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(13, 22)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Sort By"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(775, 44)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(99, 20)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtCategoryDesc
        '
        Me.txtCategoryDesc.Location = New System.Drawing.Point(571, 97)
        Me.txtCategoryDesc.Name = "txtCategoryDesc"
        Me.txtCategoryDesc.ReadOnly = True
        Me.txtCategoryDesc.Size = New System.Drawing.Size(198, 20)
        Me.txtCategoryDesc.TabIndex = 11
        Me.txtCategoryDesc.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(463, 100)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Categori Description"
        '
        'txtBrandDesc
        '
        Me.txtBrandDesc.Location = New System.Drawing.Point(571, 45)
        Me.txtBrandDesc.Name = "txtBrandDesc"
        Me.txtBrandDesc.ReadOnly = True
        Me.txtBrandDesc.Size = New System.Drawing.Size(198, 20)
        Me.txtBrandDesc.TabIndex = 8
        Me.txtBrandDesc.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(463, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Brand Description"
        '
        'cboSort
        '
        Me.cboSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSort.FormattingEnabled = True
        Me.cboSort.Items.AddRange(New Object() {"Sort Ascending", "Sort Descending"})
        Me.cboSort.Location = New System.Drawing.Point(81, 19)
        Me.cboSort.Name = "cboSort"
        Me.cboSort.Size = New System.Drawing.Size(179, 21)
        Me.cboSort.TabIndex = 5
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(571, 70)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimaryCategory.TabIndex = 16
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.Visible = False
        '
        'btnFindCategory
        '
        Me.btnFindCategory.Location = New System.Drawing.Point(740, 70)
        Me.btnFindCategory.Name = "btnFindCategory"
        Me.btnFindCategory.Size = New System.Drawing.Size(29, 20)
        Me.btnFindCategory.TabIndex = 10
        Me.btnFindCategory.TabStop = False
        Me.btnFindCategory.Text = "...."
        Me.btnFindCategory.UseVisualStyleBackColor = True
        '
        'txtCategory
        '
        Me.txtCategory.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCategory.Location = New System.Drawing.Point(571, 71)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.Size = New System.Drawing.Size(163, 20)
        Me.txtCategory.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(463, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Category"
        '
        'txtPrimaryParentGroup
        '
        Me.txtPrimaryParentGroup.Location = New System.Drawing.Point(571, 20)
        Me.txtPrimaryParentGroup.Name = "txtPrimaryParentGroup"
        Me.txtPrimaryParentGroup.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimaryParentGroup.TabIndex = 12
        Me.txtPrimaryParentGroup.Text = "-1"
        Me.txtPrimaryParentGroup.Visible = False
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.ChkGroupByPeriod)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.cboPeriod)
        Me.gboDetail.Controls.Add(Me.chkCost)
        Me.gboDetail.Controls.Add(Me.chkAbsoluteData)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.btnClose)
        Me.gboDetail.Controls.Add(Me.txtCategoryDesc)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.txtBrandDesc)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.cboSort)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCategory)
        Me.gboDetail.Controls.Add(Me.btnFindCategory)
        Me.gboDetail.Controls.Add(Me.txtCategory)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtPrimaryParentGroup)
        Me.gboDetail.Controls.Add(Me.btnFindParentGroup)
        Me.gboDetail.Controls.Add(Me.txtParentGroupCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.RBNoGroup)
        Me.gboDetail.Controls.Add(Me.dtpEndDate)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.dtpStartDate)
        Me.gboDetail.Controls.Add(Me.RBGroup)
        Me.gboDetail.Controls.Add(Me.btnFilter)
        Me.gboDetail.Location = New System.Drawing.Point(7, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(896, 132)
        Me.gboDetail.TabIndex = 2
        Me.gboDetail.TabStop = False
        '
        'ChkGroupByPeriod
        '
        Me.ChkGroupByPeriod.AutoSize = True
        Me.ChkGroupByPeriod.Location = New System.Drawing.Point(81, 74)
        Me.ChkGroupByPeriod.Name = "ChkGroupByPeriod"
        Me.ChkGroupByPeriod.Size = New System.Drawing.Size(103, 17)
        Me.ChkGroupByPeriod.TabIndex = 27
        Me.ChkGroupByPeriod.Text = "Group By Period"
        Me.ChkGroupByPeriod.UseVisualStyleBackColor = True
        '
        'btnFindParentGroup
        '
        Me.btnFindParentGroup.Location = New System.Drawing.Point(740, 18)
        Me.btnFindParentGroup.Name = "btnFindParentGroup"
        Me.btnFindParentGroup.Size = New System.Drawing.Size(29, 20)
        Me.btnFindParentGroup.TabIndex = 7
        Me.btnFindParentGroup.TabStop = False
        Me.btnFindParentGroup.Text = "...."
        Me.btnFindParentGroup.UseVisualStyleBackColor = True
        '
        'txtParentGroupCode
        '
        Me.txtParentGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtParentGroupCode.Location = New System.Drawing.Point(571, 19)
        Me.txtParentGroupCode.Name = "txtParentGroupCode"
        Me.txtParentGroupCode.Size = New System.Drawing.Size(163, 20)
        Me.txtParentGroupCode.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(463, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Brand"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(123, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To"
        Me.Label4.Visible = False
        '
        'RBNoGroup
        '
        Me.RBNoGroup.AutoSize = True
        Me.RBNoGroup.Location = New System.Drawing.Point(6, 44)
        Me.RBNoGroup.Name = "RBNoGroup"
        Me.RBNoGroup.Size = New System.Drawing.Size(85, 17)
        Me.RBNoGroup.TabIndex = 1
        Me.RBNoGroup.Text = "No Grouping"
        Me.RBNoGroup.UseVisualStyleBackColor = True
        Me.RBNoGroup.Visible = False
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(191, 45)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(179, 20)
        Me.dtpEndDate.TabIndex = 4
        Me.dtpEndDate.TabStop = False
        Me.dtpEndDate.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(123, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        Me.Label3.Visible = False
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(191, 19)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(179, 20)
        Me.dtpStartDate.TabIndex = 3
        Me.dtpStartDate.TabStop = False
        Me.dtpStartDate.Visible = False
        '
        'RBGroup
        '
        Me.RBGroup.AutoSize = True
        Me.RBGroup.Checked = True
        Me.RBGroup.Location = New System.Drawing.Point(6, 21)
        Me.RBGroup.Name = "RBGroup"
        Me.RBGroup.Size = New System.Drawing.Size(100, 17)
        Me.RBGroup.TabIndex = 0
        Me.RBGroup.TabStop = True
        Me.RBGroup.Text = "Group By Brand"
        Me.RBGroup.UseVisualStyleBackColor = True
        Me.RBGroup.Visible = False
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(775, 16)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(99, 20)
        Me.btnFilter.TabIndex = 12
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.Location = New System.Drawing.Point(7, 142)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(896, 335)
        Me.crv.TabIndex = 3
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanProductRanking_Modif
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(910, 481)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.crv)
        Me.Name = "FrmLaporanProductRanking_Modif"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Non Moving Stock Report"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents chkCost As System.Windows.Forms.CheckBox
    Friend WithEvents chkAbsoluteData As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtCategoryDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtBrandDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboSort As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCategory As System.Windows.Forms.Button
    Friend WithEvents txtCategory As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryParentGroup As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnFindParentGroup As System.Windows.Forms.Button
    Friend WithEvents txtParentGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents RBNoGroup As System.Windows.Forms.RadioButton
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents RBGroup As System.Windows.Forms.RadioButton
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents ChkGroupByPeriod As System.Windows.Forms.CheckBox
End Class
