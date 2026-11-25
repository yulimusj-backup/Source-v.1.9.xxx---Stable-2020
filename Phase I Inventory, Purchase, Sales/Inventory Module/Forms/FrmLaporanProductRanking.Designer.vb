<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProductRanking
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
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.btnClose = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
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
        Me.btnFindParentGroup = New System.Windows.Forms.Button
        Me.txtParentGroupCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.RBNoGroup = New System.Windows.Forms.RadioButton
        Me.RBGroup = New System.Windows.Forms.RadioButton
        Me.btnFilter = New System.Windows.Forms.Button
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.gboDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(123, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(170, 45)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(179, 20)
        Me.dtpEndDate.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(123, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(170, 19)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(179, 20)
        Me.dtpStartDate.TabIndex = 2
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(738, 47)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(99, 20)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
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
        Me.gboDetail.Location = New System.Drawing.Point(12, 6)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(877, 132)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(123, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Sort By"
        '
        'txtCategoryDesc
        '
        Me.txtCategoryDesc.Location = New System.Drawing.Point(485, 97)
        Me.txtCategoryDesc.Name = "txtCategoryDesc"
        Me.txtCategoryDesc.ReadOnly = True
        Me.txtCategoryDesc.Size = New System.Drawing.Size(198, 20)
        Me.txtCategoryDesc.TabIndex = 21
        Me.txtCategoryDesc.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(377, 100)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Categori Description"
        '
        'txtBrandDesc
        '
        Me.txtBrandDesc.Location = New System.Drawing.Point(485, 45)
        Me.txtBrandDesc.Name = "txtBrandDesc"
        Me.txtBrandDesc.ReadOnly = True
        Me.txtBrandDesc.Size = New System.Drawing.Size(198, 20)
        Me.txtBrandDesc.TabIndex = 19
        Me.txtBrandDesc.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(377, 48)
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
        Me.cboSort.Location = New System.Drawing.Point(170, 71)
        Me.cboSort.Name = "cboSort"
        Me.cboSort.Size = New System.Drawing.Size(179, 21)
        Me.cboSort.TabIndex = 4
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(485, 70)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimaryCategory.TabIndex = 16
        Me.txtPrimaryCategory.Text = "-1"
        Me.txtPrimaryCategory.Visible = False
        '
        'btnFindCategory
        '
        Me.btnFindCategory.Location = New System.Drawing.Point(654, 70)
        Me.btnFindCategory.Name = "btnFindCategory"
        Me.btnFindCategory.Size = New System.Drawing.Size(29, 20)
        Me.btnFindCategory.TabIndex = 15
        Me.btnFindCategory.TabStop = False
        Me.btnFindCategory.Text = "...."
        Me.btnFindCategory.UseVisualStyleBackColor = True
        '
        'txtCategory
        '
        Me.txtCategory.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCategory.Location = New System.Drawing.Point(485, 71)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.Size = New System.Drawing.Size(163, 20)
        Me.txtCategory.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(377, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Category"
        '
        'txtPrimaryParentGroup
        '
        Me.txtPrimaryParentGroup.Location = New System.Drawing.Point(485, 20)
        Me.txtPrimaryParentGroup.Name = "txtPrimaryParentGroup"
        Me.txtPrimaryParentGroup.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimaryParentGroup.TabIndex = 12
        Me.txtPrimaryParentGroup.Text = "-1"
        Me.txtPrimaryParentGroup.Visible = False
        '
        'btnFindParentGroup
        '
        Me.btnFindParentGroup.Location = New System.Drawing.Point(654, 18)
        Me.btnFindParentGroup.Name = "btnFindParentGroup"
        Me.btnFindParentGroup.Size = New System.Drawing.Size(29, 20)
        Me.btnFindParentGroup.TabIndex = 11
        Me.btnFindParentGroup.TabStop = False
        Me.btnFindParentGroup.Text = "...."
        Me.btnFindParentGroup.UseVisualStyleBackColor = True
        '
        'txtParentGroupCode
        '
        Me.txtParentGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtParentGroupCode.Location = New System.Drawing.Point(485, 19)
        Me.txtParentGroupCode.Name = "txtParentGroupCode"
        Me.txtParentGroupCode.Size = New System.Drawing.Size(163, 20)
        Me.txtParentGroupCode.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(377, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Brand"
        '
        'RBNoGroup
        '
        Me.RBNoGroup.AutoSize = True
        Me.RBNoGroup.Checked = True
        Me.RBNoGroup.Location = New System.Drawing.Point(6, 44)
        Me.RBNoGroup.Name = "RBNoGroup"
        Me.RBNoGroup.Size = New System.Drawing.Size(85, 17)
        Me.RBNoGroup.TabIndex = 1
        Me.RBNoGroup.TabStop = True
        Me.RBNoGroup.Text = "No Grouping"
        Me.RBNoGroup.UseVisualStyleBackColor = True
        '
        'RBGroup
        '
        Me.RBGroup.AutoSize = True
        Me.RBGroup.Location = New System.Drawing.Point(6, 21)
        Me.RBGroup.Name = "RBGroup"
        Me.RBGroup.Size = New System.Drawing.Size(100, 17)
        Me.RBGroup.TabIndex = 0
        Me.RBGroup.TabStop = True
        Me.RBGroup.Text = "Group By Brand"
        Me.RBGroup.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(738, 19)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(99, 20)
        Me.btnFilter.TabIndex = 7
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
        Me.crv.Location = New System.Drawing.Point(12, 144)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(877, 344)
        Me.crv.TabIndex = 1
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'FrmLaporanProductRanking
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(901, 500)
        Me.Controls.Add(Me.crv)
        Me.Controls.Add(Me.gboDetail)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(907, 528)
        Me.Name = "FrmLaporanProductRanking"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Ranking"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents RBGroup As System.Windows.Forms.RadioButton
    Friend WithEvents RBNoGroup As System.Windows.Forms.RadioButton
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCategory As System.Windows.Forms.Button
    Friend WithEvents txtCategory As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryParentGroup As System.Windows.Forms.TextBox
    Friend WithEvents btnFindParentGroup As System.Windows.Forms.Button
    Friend WithEvents txtParentGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboSort As System.Windows.Forms.ComboBox
    Friend WithEvents txtCategoryDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtBrandDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
