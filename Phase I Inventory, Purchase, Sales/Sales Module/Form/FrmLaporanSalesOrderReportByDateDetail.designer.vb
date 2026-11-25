<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalesOrderReportByDateDetail
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtProductCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtProductName = New System.Windows.Forms.TextBox
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.txtCustomerName = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.txtPrimarySalesPerson = New System.Windows.Forms.TextBox
        Me.btnFindSalesPerson = New System.Windows.Forms.Button
        Me.txtSalesPerson = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtSalesPersonName = New System.Windows.Forms.TextBox
        Me.ChkAllProductSO = New System.Windows.Forms.CheckBox
        Me.ChkOS = New System.Windows.Forms.CheckBox
        Me.ChkShowRatio = New System.Windows.Forms.CheckBox
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(257, 58)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(57, 23)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(257, 28)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(57, 23)
        Me.btnFilter.TabIndex = 12
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(240, 9)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(233, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(86, 9)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(110, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'crv
        '
        Me.crv.ActiveViewIndex = -1
        Me.crv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv.DisplayGroupTree = False
        Me.crv.EnableDrillDown = False
        Me.crv.Location = New System.Drawing.Point(2, 185)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(795, 319)
        Me.crv.TabIndex = 24
        Me.crv.TabStop = False
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'txtProductCode
        '
        Me.txtProductCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtProductCode.Location = New System.Drawing.Point(106, 362)
        Me.txtProductCode.Name = "txtProductCode"
        Me.txtProductCode.Size = New System.Drawing.Size(142, 20)
        Me.txtProductCode.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 365)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Product"
        '
        'txtProductName
        '
        Me.txtProductName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProductName.Location = New System.Drawing.Point(290, 362)
        Me.txtProductName.Name = "txtProductName"
        Me.txtProductName.ReadOnly = True
        Me.txtProductName.Size = New System.Drawing.Size(411, 20)
        Me.txtProductName.TabIndex = 8
        Me.txtProductName.TabStop = False
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(254, 362)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(31, 18)
        Me.btnFindProduct.TabIndex = 7
        Me.btnFindProduct.TabStop = False
        Me.btnFindProduct.Text = "..."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Customer"
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(202, 33)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(31, 18)
        Me.btnFindCustomer.TabIndex = 9
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "..."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'txtCustomerName
        '
        Me.txtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomerName.Location = New System.Drawing.Point(240, 33)
        Me.txtCustomerName.Name = "txtCustomerName"
        Me.txtCustomerName.ReadOnly = True
        Me.txtCustomerName.Size = New System.Drawing.Size(233, 20)
        Me.txtCustomerName.TabIndex = 16
        Me.txtCustomerName.TabStop = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(86, 33)
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(110, 20)
        Me.txtCustomerCode.TabIndex = 8
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Location = New System.Drawing.Point(86, 82)
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.Size = New System.Drawing.Size(110, 20)
        Me.txtCurrencyCode.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 82)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(77, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Currency Code"
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(202, 82)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(31, 18)
        Me.btnFindCurrency.TabIndex = 5
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'txtPrimaryProduct
        '
        Me.txtPrimaryProduct.Location = New System.Drawing.Point(106, 362)
        Me.txtPrimaryProduct.Name = "txtPrimaryProduct"
        Me.txtPrimaryProduct.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryProduct.TabIndex = 5
        Me.txtPrimaryProduct.Text = "-1"
        Me.txtPrimaryProduct.Visible = False
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(10, 106)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryCustomer.TabIndex = 10
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(46, 106)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimaryCurrency.TabIndex = 15
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'txtPrimarySalesPerson
        '
        Me.txtPrimarySalesPerson.Location = New System.Drawing.Point(86, 106)
        Me.txtPrimarySalesPerson.Name = "txtPrimarySalesPerson"
        Me.txtPrimarySalesPerson.Size = New System.Drawing.Size(34, 20)
        Me.txtPrimarySalesPerson.TabIndex = 19
        Me.txtPrimarySalesPerson.Text = "-1"
        Me.txtPrimarySalesPerson.Visible = False
        '
        'btnFindSalesPerson
        '
        Me.btnFindSalesPerson.Location = New System.Drawing.Point(202, 58)
        Me.btnFindSalesPerson.Name = "btnFindSalesPerson"
        Me.btnFindSalesPerson.Size = New System.Drawing.Size(31, 18)
        Me.btnFindSalesPerson.TabIndex = 11
        Me.btnFindSalesPerson.TabStop = False
        Me.btnFindSalesPerson.Text = "..."
        Me.btnFindSalesPerson.UseVisualStyleBackColor = True
        '
        'txtSalesPerson
        '
        Me.txtSalesPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSalesPerson.Location = New System.Drawing.Point(86, 58)
        Me.txtSalesPerson.Name = "txtSalesPerson"
        Me.txtSalesPerson.Size = New System.Drawing.Size(110, 20)
        Me.txtSalesPerson.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Sales Person"
        '
        'TxtSalesPersonName
        '
        Me.TxtSalesPersonName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesPersonName.Location = New System.Drawing.Point(240, 58)
        Me.TxtSalesPersonName.Name = "TxtSalesPersonName"
        Me.TxtSalesPersonName.ReadOnly = True
        Me.TxtSalesPersonName.Size = New System.Drawing.Size(233, 20)
        Me.TxtSalesPersonName.TabIndex = 25
        Me.TxtSalesPersonName.TabStop = False
        '
        'ChkAllProductSO
        '
        Me.ChkAllProductSO.AutoSize = True
        Me.ChkAllProductSO.Location = New System.Drawing.Point(255, 101)
        Me.ChkAllProductSO.Name = "ChkAllProductSO"
        Me.ChkAllProductSO.Size = New System.Drawing.Size(152, 17)
        Me.ChkAllProductSO.TabIndex = 2
        Me.ChkAllProductSO.Text = "Always display all SO items"
        Me.ChkAllProductSO.UseVisualStyleBackColor = True
        '
        'ChkOS
        '
        Me.ChkOS.AutoSize = True
        Me.ChkOS.Location = New System.Drawing.Point(255, 84)
        Me.ChkOS.Name = "ChkOS"
        Me.ChkOS.Size = New System.Drawing.Size(215, 17)
        Me.ChkOS.TabIndex = 3
        Me.ChkOS.Text = "Only show SO with sales orders balance"
        Me.ChkOS.UseVisualStyleBackColor = True
        '
        'ChkShowRatio
        '
        Me.ChkShowRatio.AutoSize = True
        Me.ChkShowRatio.Location = New System.Drawing.Point(255, 118)
        Me.ChkShowRatio.Name = "ChkShowRatio"
        Me.ChkShowRatio.Size = New System.Drawing.Size(142, 17)
        Me.ChkShowRatio.TabIndex = 26
        Me.ChkShowRatio.Text = "Show Ratio (If Available)"
        Me.ChkShowRatio.UseVisualStyleBackColor = True
        '
        'BackgroundWorker3
        '
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        Me.dgvItemProductGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemProductGroup.Location = New System.Drawing.Point(2, 15)
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidth = 24
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        Me.dgvItemProductGroup.Size = New System.Drawing.Size(245, 156)
        Me.dgvItemProductGroup.TabIndex = 7
        '
        'ColProductGroup_No
        '
        Me.ColProductGroup_No.HeaderText = "No"
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        Me.ColProductGroup_No.Width = 46
        '
        'ColProductGroup_Level
        '
        Me.ColProductGroup_Level.HeaderText = "Level"
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        Me.ColProductGroup_Level.Width = 39
        '
        'ColProductGroup_CodeNo
        '
        Me.ColProductGroup_CodeNo.HeaderText = "Code"
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        Me.ColProductGroup_CodeNo.Width = 57
        '
        'ColProductGroup_Desc
        '
        Me.ColProductGroup_Desc.HeaderText = "Description"
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        Me.ColProductGroup_Desc.Width = 85
        '
        'ColProductGroup_Remarks
        '
        Me.ColProductGroup_Remarks.HeaderText = "Remarks"
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.ReadOnly = True
        Me.ColProductGroup_Remarks.Visible = False
        Me.ColProductGroup_Remarks.Width = 89
        '
        'ColProductGroup_PrimaryAcuan
        '
        Me.ColProductGroup_PrimaryAcuan.HeaderText = "Primary"
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        Me.ColProductGroup_PrimaryAcuan.Visible = False
        Me.ColProductGroup_PrimaryAcuan.Width = 81
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtSalesPersonName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpTo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowRatio)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkOS)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimarySalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkAllProductSO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindSalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCurrencyCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtSalesPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(798, 178)
        Me.SplitContainer1.SplitterDistance = 478
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 28
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvItemProductGroup)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(249, 173)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Product Filter"
        '
        'FrmLaporanSalesOrderReportByDateDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(801, 514)
        Me.Controls.Add(Me.crv)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.txtPrimaryProduct)
        Me.Controls.Add(Me.btnFindProduct)
        Me.Controls.Add(Me.txtProductName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtProductCode)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanSalesOrderReportByDateDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Order Report By Date (Detail)"
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtProductName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents txtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimarySalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSalesPerson As System.Windows.Forms.Button
    Friend WithEvents txtSalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSalesPersonName As System.Windows.Forms.TextBox
    Friend WithEvents ChkAllProductSO As System.Windows.Forms.CheckBox
    Friend WithEvents ChkOS As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowRatio As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
