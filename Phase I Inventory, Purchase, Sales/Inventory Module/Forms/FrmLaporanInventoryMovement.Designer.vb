<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanInventoryMovement
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvStockCard = New System.Windows.Forms.DataGridView
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.CboGroup2 = New System.Windows.Forms.ComboBox
        Me.CboGroup1 = New System.Windows.Forms.ComboBox
        Me.CboGrouping = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.chkService = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkGoods = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.CrvLaporan = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboPilihGroupLevel = New System.Windows.Forms.ComboBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.DgvGroup = New System.Windows.Forms.DataGridView
        Me.ColGroup_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.clbBrand = New System.Windows.Forms.CheckedListBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        CType(Me.dgvStockCard, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        Me.gboBrand.SuspendLayout()
        CType(Me.DgvGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvStockCard
        '
        Me.dgvStockCard.AllowUserToAddRows = False
        Me.dgvStockCard.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvStockCard.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvStockCard.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStockCard.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvStockCard.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStockCard.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvStockCard.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStockCard.Location = New System.Drawing.Point(325, 10)
        Me.dgvStockCard.MultiSelect = False
        Me.dgvStockCard.Name = "dgvStockCard"
        Me.dgvStockCard.ReadOnly = True
        Me.dgvStockCard.RowHeadersVisible = False
        Me.dgvStockCard.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStockCard.Size = New System.Drawing.Size(574, 466)
        Me.dgvStockCard.StandardTab = True
        Me.dgvStockCard.TabIndex = 1
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.CboGroup2)
        Me.gboDetail.Controls.Add(Me.CboGroup1)
        Me.gboDetail.Controls.Add(Me.CboGrouping)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.chkService)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.chkGoods)
        Me.gboDetail.Controls.Add(Me.btnClose)
        Me.gboDetail.Controls.Add(Me.btnFilter)
        Me.gboDetail.Controls.Add(Me.btnPrint)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.dtpEndDate)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.dtpStartDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimary)
        Me.gboDetail.Location = New System.Drawing.Point(7, 3)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(308, 179)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'CboGroup2
        '
        Me.CboGroup2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroup2.FormattingEnabled = True
        Me.CboGroup2.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category"})
        Me.CboGroup2.Location = New System.Drawing.Point(68, 121)
        Me.CboGroup2.Name = "CboGroup2"
        Me.CboGroup2.Size = New System.Drawing.Size(155, 21)
        Me.CboGroup2.TabIndex = 4
        '
        'CboGroup1
        '
        Me.CboGroup1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroup1.FormattingEnabled = True
        Me.CboGroup1.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category"})
        Me.CboGroup1.Location = New System.Drawing.Point(68, 96)
        Me.CboGroup1.Name = "CboGroup1"
        Me.CboGroup1.Size = New System.Drawing.Size(155, 21)
        Me.CboGroup1.TabIndex = 3
        '
        'CboGrouping
        '
        Me.CboGrouping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGrouping.FormattingEnabled = True
        Me.CboGrouping.Items.AddRange(New Object() {"Double Grouping", "Single Grouping", "No Grouping"})
        Me.CboGrouping.Location = New System.Drawing.Point(68, 71)
        Me.CboGrouping.Name = "CboGrouping"
        Me.CboGrouping.Size = New System.Drawing.Size(155, 21)
        Me.CboGrouping.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 124)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Group 2"
        '
        'chkService
        '
        Me.chkService.AutoSize = True
        Me.chkService.Checked = True
        Me.chkService.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkService.Location = New System.Drawing.Point(138, 152)
        Me.chkService.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkService.Name = "chkService"
        Me.chkService.Size = New System.Drawing.Size(62, 17)
        Me.chkService.TabIndex = 6
        Me.chkService.Text = "Service"
        Me.chkService.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Group 1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Grouping"
        '
        'chkGoods
        '
        Me.chkGoods.AutoSize = True
        Me.chkGoods.Checked = True
        Me.chkGoods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkGoods.Location = New System.Drawing.Point(68, 152)
        Me.chkGoods.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkGoods.Name = "chkGoods"
        Me.chkGoods.Size = New System.Drawing.Size(57, 17)
        Me.chkGoods.TabIndex = 5
        Me.chkGoods.Text = "Goods"
        Me.chkGoods.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(234, 45)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(65, 20)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(234, 19)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(65, 20)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(234, 45)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(65, 20)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        Me.btnPrint.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "To"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpEndDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(68, 45)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(155, 20)
        Me.dtpEndDate.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStartDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(68, 19)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(155, 20)
        Me.dtpStartDate.TabIndex = 0
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(68, 19)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(36, 20)
        Me.TxtPrimary.TabIndex = 17
        Me.TxtPrimary.Visible = False
        '
        'CrvLaporan
        '
        Me.CrvLaporan.ActiveViewIndex = -1
        Me.CrvLaporan.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrvLaporan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvLaporan.DisplayGroupTree = False
        Me.CrvLaporan.EnableDrillDown = False
        Me.CrvLaporan.Location = New System.Drawing.Point(325, 10)
        Me.CrvLaporan.Name = "CrvLaporan"
        Me.CrvLaporan.SelectionFormula = ""
        Me.CrvLaporan.ShowCloseButton = False
        Me.CrvLaporan.ShowGroupTreeButton = False
        Me.CrvLaporan.ShowRefreshButton = False
        Me.CrvLaporan.Size = New System.Drawing.Size(574, 466)
        Me.CrvLaporan.TabIndex = 3
        Me.CrvLaporan.TabStop = False
        Me.CrvLaporan.ViewTimeSelectionFormula = ""
        '
        'gboBrand
        '
        Me.gboBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.gboBrand.Controls.Add(Me.Label1)
        Me.gboBrand.Controls.Add(Me.CboPilihGroupLevel)
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.DgvGroup)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbBrand)
        Me.gboBrand.Location = New System.Drawing.Point(7, 188)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Size = New System.Drawing.Size(308, 288)
        Me.gboBrand.TabIndex = 1
        Me.gboBrand.TabStop = False
        Me.gboBrand.Text = "Product Group Filter"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(33, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Level"
        '
        'CboPilihGroupLevel
        '
        Me.CboPilihGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPilihGroupLevel.FormattingEnabled = True
        Me.CboPilihGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.CboPilihGroupLevel.Location = New System.Drawing.Point(68, 21)
        Me.CboPilihGroupLevel.Name = "CboPilihGroupLevel"
        Me.CboPilihGroupLevel.Size = New System.Drawing.Size(155, 21)
        Me.CboPilihGroupLevel.TabIndex = 7
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(234, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        Me.LinkLabel2.Visible = False
        '
        'DgvGroup
        '
        Me.DgvGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DgvGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColGroup_Code, Me.ColGroup_Desc, Me.ColGroup_Primary})
        Me.DgvGroup.Location = New System.Drawing.Point(3, 48)
        Me.DgvGroup.MultiSelect = False
        Me.DgvGroup.Name = "DgvGroup"
        Me.DgvGroup.RowTemplate.Height = 24
        Me.DgvGroup.Size = New System.Drawing.Size(302, 233)
        Me.DgvGroup.TabIndex = 8
        '
        'ColGroup_Code
        '
        Me.ColGroup_Code.HeaderText = "Code"
        Me.ColGroup_Code.MaxInputLength = 250
        Me.ColGroup_Code.Name = "ColGroup_Code"
        Me.ColGroup_Code.Width = 57
        '
        'ColGroup_Desc
        '
        Me.ColGroup_Desc.HeaderText = "Description"
        Me.ColGroup_Desc.Name = "ColGroup_Desc"
        Me.ColGroup_Desc.ReadOnly = True
        Me.ColGroup_Desc.Width = 85
        '
        'ColGroup_Primary
        '
        Me.ColGroup_Primary.HeaderText = "Primary"
        Me.ColGroup_Primary.Name = "ColGroup_Primary"
        Me.ColGroup_Primary.ReadOnly = True
        Me.ColGroup_Primary.Visible = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(179, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        Me.LinkLabel1.Visible = False
        '
        'clbBrand
        '
        Me.clbBrand.CheckOnClick = True
        Me.clbBrand.ColumnWidth = 300
        Me.clbBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbBrand.FormattingEnabled = True
        Me.clbBrand.Location = New System.Drawing.Point(3, 16)
        Me.clbBrand.MultiColumn = True
        Me.clbBrand.Name = "clbBrand"
        Me.clbBrand.Size = New System.Drawing.Size(302, 259)
        Me.clbBrand.TabIndex = 0
        Me.clbBrand.ThreeDCheckBoxes = True
        Me.clbBrand.Visible = False
        '
        'BackgroundWorker2
        '
        '
        'FrmLaporanInventoryMovement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(911, 490)
        Me.Controls.Add(Me.CrvLaporan)
        Me.Controls.Add(Me.dgvStockCard)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.gboBrand)
        Me.MinimumSize = New System.Drawing.Size(909, 528)
        Me.Name = "FrmLaporanInventoryMovement"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Movement Reports"
        CType(Me.dgvStockCard, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        CType(Me.DgvGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvStockCard As System.Windows.Forms.DataGridView
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CrvLaporan As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents chkService As System.Windows.Forms.CheckBox
    Friend WithEvents chkGoods As System.Windows.Forms.CheckBox
    Friend WithEvents CboPilihGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents DgvGroup As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColGroup_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents clbBrand As System.Windows.Forms.CheckedListBox
    Friend WithEvents CboGroup2 As System.Windows.Forms.ComboBox
    Friend WithEvents CboGroup1 As System.Windows.Forms.ComboBox
    Friend WithEvents CboGrouping As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
End Class
