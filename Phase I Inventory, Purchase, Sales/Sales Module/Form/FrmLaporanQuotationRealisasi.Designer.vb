<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanQuotationRealisasi
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
        Me.Crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtDescCustomer = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BtnFindCustomerCode = New System.Windows.Forms.Button
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BtnFindSales = New System.Windows.Forms.Button
        Me.TxtKodeSales = New System.Windows.Forms.TextBox
        Me.TxtDescSalesPerson = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TxtPrimarySales = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Crv
        '
        Me.Crv.ActiveViewIndex = -1
        Me.Crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Crv.DisplayGroupTree = False
        Me.Crv.EnableDrillDown = False
        Me.Crv.Location = New System.Drawing.Point(12, 173)
        Me.Crv.Margin = New System.Windows.Forms.Padding(4)
        Me.Crv.Name = "Crv"
        Me.Crv.SelectionFormula = ""
        Me.Crv.ShowCloseButton = False
        Me.Crv.ShowGroupTreeButton = False
        Me.Crv.ShowRefreshButton = False
        Me.Crv.Size = New System.Drawing.Size(970, 447)
        Me.Crv.TabIndex = 3
        Me.Crv.TabStop = False
        Me.Crv.ViewTimeSelectionFormula = ""
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(147, 64)
        Me.TxtCustomerCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(99, 22)
        Me.TxtCustomerCode.TabIndex = 8
        '
        'TxtDescCustomer
        '
        Me.TxtDescCustomer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescCustomer.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescCustomer.Location = New System.Drawing.Point(319, 64)
        Me.TxtDescCustomer.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDescCustomer.Name = "TxtDescCustomer"
        Me.TxtDescCustomer.ReadOnly = True
        Me.TxtDescCustomer.Size = New System.Drawing.Size(235, 22)
        Me.TxtDescCustomer.TabIndex = 10
        Me.TxtDescCustomer.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(290, 6)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(25, 17)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "To"
        '
        'BtnFindCustomerCode
        '
        Me.BtnFindCustomerCode.Location = New System.Drawing.Point(254, 64)
        Me.BtnFindCustomerCode.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFindCustomerCode.Name = "BtnFindCustomerCode"
        Me.BtnFindCustomerCode.Size = New System.Drawing.Size(57, 25)
        Me.BtnFindCustomerCode.TabIndex = 9
        Me.BtnFindCustomerCode.TabStop = False
        Me.BtnFindCustomerCode.Text = "Find"
        Me.BtnFindCustomerCode.UseVisualStyleBackColor = True
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(323, 4)
        Me.DtpTo.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(135, 22)
        Me.DtpTo.TabIndex = 4
        '
        'DtpFrom
        '
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(147, 4)
        Me.DtpFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(135, 22)
        Me.DtpFrom.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Period"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 66)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 17)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Sales Person Group"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(11, 13)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboFormat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDescCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindSales)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtKodeSales)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDescSalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySales)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(971, 152)
        Me.SplitContainer1.SplitterDistance = 558
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 2
        Me.SplitContainer1.TabStop = False
        '
        'CboFormat
        '
        Me.CboFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Group By Sales Person", "Group By Sales Person Group", "Group By Product Group", "Group By Brand", "Group By Sub Brand", "Group By Category", "Group By Product"})
        Me.CboFormat.Location = New System.Drawing.Point(147, 94)
        Me.CboFormat.Margin = New System.Windows.Forms.Padding(4)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(230, 24)
        Me.CboFormat.TabIndex = 25
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 98)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 17)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Format"
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(465, 4)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(89, 28)
        Me.BtnFilter.TabIndex = 3
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(465, 4)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(89, 28)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnFindSales
        '
        Me.BtnFindSales.Location = New System.Drawing.Point(254, 32)
        Me.BtnFindSales.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFindSales.Name = "BtnFindSales"
        Me.BtnFindSales.Size = New System.Drawing.Size(57, 26)
        Me.BtnFindSales.TabIndex = 6
        Me.BtnFindSales.TabStop = False
        Me.BtnFindSales.Text = "Find"
        Me.BtnFindSales.UseVisualStyleBackColor = True
        '
        'TxtKodeSales
        '
        Me.TxtKodeSales.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeSales.Location = New System.Drawing.Point(147, 34)
        Me.TxtKodeSales.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtKodeSales.Name = "TxtKodeSales"
        Me.TxtKodeSales.Size = New System.Drawing.Size(99, 22)
        Me.TxtKodeSales.TabIndex = 5
        '
        'TxtDescSalesPerson
        '
        Me.TxtDescSalesPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDescSalesPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDescSalesPerson.Location = New System.Drawing.Point(319, 34)
        Me.TxtDescSalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtDescSalesPerson.Name = "TxtDescSalesPerson"
        Me.TxtDescSalesPerson.ReadOnly = True
        Me.TxtDescSalesPerson.Size = New System.Drawing.Size(235, 22)
        Me.TxtDescSalesPerson.TabIndex = 7
        Me.TxtDescSalesPerson.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 37)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(92, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Sales Person"
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(3, 65)
        Me.TxtPrimaryCurrency.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(39, 22)
        Me.TxtPrimaryCurrency.TabIndex = 24
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Text = "-1"
        Me.TxtPrimaryCurrency.Visible = False
        '
        'TxtPrimarySales
        '
        Me.TxtPrimarySales.Location = New System.Drawing.Point(3, 34)
        Me.TxtPrimarySales.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySales.Name = "TxtPrimarySales"
        Me.TxtPrimarySales.Size = New System.Drawing.Size(39, 22)
        Me.TxtPrimarySales.TabIndex = 13
        Me.TxtPrimarySales.TabStop = False
        Me.TxtPrimarySales.Text = "-1"
        Me.TxtPrimarySales.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvItemProductGroup)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(408, 152)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Product Group"
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeight = 25
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(3, 18)
        Me.dgvItemProductGroup.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(402, 131)
        Me.dgvItemProductGroup.TabIndex = 0
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.Frozen = True
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_No.Width = 32
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.Frozen = True
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 48
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.Frozen = True
        Me.ColProductGroup_CodeNo.HeaderText = "Group Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_CodeNo.Width = 91
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_Remarks.Width = 70
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 106
        '
        'FrmLaporanQuotationRealisasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.Crv)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanQuotationRealisasi"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quotation Realization"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCustomerCode As System.Windows.Forms.Button
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFindSales As System.Windows.Forms.Button
    Friend WithEvents TxtKodeSales As System.Windows.Forms.TextBox
    Friend WithEvents TxtDescSalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimarySales As System.Windows.Forms.TextBox
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
