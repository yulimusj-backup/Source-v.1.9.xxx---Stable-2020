<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_RawMaterialUsage
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
        Me.dtpToBundlingDate = New System.Windows.Forms.DateTimePicker
        Me.chkIncludeWIP = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.crv = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.ChkDescRM = New System.Windows.Forms.CheckBox
        Me.ChkCodeRM = New System.Windows.Forms.CheckBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.ChkDescFG = New System.Windows.Forms.CheckBox
        Me.ChkCodeFG = New System.Windows.Forms.CheckBox
        Me.ChkShowZeroUsed = New System.Windows.Forms.CheckBox
        Me.btnShow = New System.Windows.Forms.Button
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkLogDateFrom = New System.Windows.Forms.CheckBox
        Me.dtpFromBundlingDate = New System.Windows.Forms.DateTimePicker
        Me.colMO_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvMO = New System.Windows.Forms.DataGridView
        Me.colMO_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colMO_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colMO_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ChkMODate = New System.Windows.Forms.CheckBox
        Me.dtpToMODate = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtpFromMODate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dgvProduct = New System.Windows.Forms.DataGridView
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.colProduct_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_GroupLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colProduct_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colProduct_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvMO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpToBundlingDate
        '
        Me.dtpToBundlingDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpToBundlingDate.CustomFormat = "dd MMM yyyy"
        Me.dtpToBundlingDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToBundlingDate.Location = New System.Drawing.Point(65, 45)
        Me.dtpToBundlingDate.Name = "dtpToBundlingDate"
        Me.dtpToBundlingDate.Size = New System.Drawing.Size(125, 20)
        Me.dtpToBundlingDate.TabIndex = 2
        '
        'chkIncludeWIP
        '
        Me.chkIncludeWIP.AutoSize = True
        Me.chkIncludeWIP.Location = New System.Drawing.Point(51, 61)
        Me.chkIncludeWIP.Name = "chkIncludeWIP"
        Me.chkIncludeWIP.Size = New System.Drawing.Size(85, 17)
        Me.chkIncludeWIP.TabIndex = 2
        Me.chkIncludeWIP.Text = "Include WIP"
        Me.chkIncludeWIP.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
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
        Me.crv.Location = New System.Drawing.Point(3, 89)
        Me.crv.Name = "crv"
        Me.crv.SelectionFormula = ""
        Me.crv.ShowCloseButton = False
        Me.crv.ShowGroupTreeButton = False
        Me.crv.ShowRefreshButton = False
        Me.crv.Size = New System.Drawing.Size(672, 469)
        Me.crv.TabIndex = 1
        Me.crv.ViewTimeSelectionFormula = ""
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.GroupBox7)
        Me.GroupBox5.Controls.Add(Me.GroupBox6)
        Me.GroupBox5.Controls.Add(Me.ChkShowZeroUsed)
        Me.GroupBox5.Controls.Add(Me.chkIncludeWIP)
        Me.GroupBox5.Controls.Add(Me.btnShow)
        Me.GroupBox5.Controls.Add(Me.CboFormat)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(671, 80)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.ChkDescRM)
        Me.GroupBox7.Controls.Add(Me.ChkCodeRM)
        Me.GroupBox7.Enabled = False
        Me.GroupBox7.Location = New System.Drawing.Point(330, 11)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(127, 61)
        Me.GroupBox7.TabIndex = 4
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Raw Material"
        '
        'ChkDescRM
        '
        Me.ChkDescRM.AutoSize = True
        Me.ChkDescRM.Checked = True
        Me.ChkDescRM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkDescRM.Location = New System.Drawing.Point(8, 39)
        Me.ChkDescRM.Name = "ChkDescRM"
        Me.ChkDescRM.Size = New System.Drawing.Size(109, 17)
        Me.ChkDescRM.TabIndex = 1
        Me.ChkDescRM.Text = "Show Description"
        Me.ChkDescRM.UseVisualStyleBackColor = True
        '
        'ChkCodeRM
        '
        Me.ChkCodeRM.AutoSize = True
        Me.ChkCodeRM.Checked = True
        Me.ChkCodeRM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkCodeRM.Location = New System.Drawing.Point(8, 16)
        Me.ChkCodeRM.Name = "ChkCodeRM"
        Me.ChkCodeRM.Size = New System.Drawing.Size(81, 17)
        Me.ChkCodeRM.TabIndex = 0
        Me.ChkCodeRM.Text = "Show Code"
        Me.ChkCodeRM.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ChkDescFG)
        Me.GroupBox6.Controls.Add(Me.ChkCodeFG)
        Me.GroupBox6.Enabled = False
        Me.GroupBox6.Location = New System.Drawing.Point(197, 11)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(127, 61)
        Me.GroupBox6.TabIndex = 3
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Finish Good"
        '
        'ChkDescFG
        '
        Me.ChkDescFG.AutoSize = True
        Me.ChkDescFG.Checked = True
        Me.ChkDescFG.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkDescFG.Location = New System.Drawing.Point(8, 39)
        Me.ChkDescFG.Name = "ChkDescFG"
        Me.ChkDescFG.Size = New System.Drawing.Size(109, 17)
        Me.ChkDescFG.TabIndex = 1
        Me.ChkDescFG.Text = "Show Description"
        Me.ChkDescFG.UseVisualStyleBackColor = True
        '
        'ChkCodeFG
        '
        Me.ChkCodeFG.AutoSize = True
        Me.ChkCodeFG.Checked = True
        Me.ChkCodeFG.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkCodeFG.Location = New System.Drawing.Point(8, 16)
        Me.ChkCodeFG.Name = "ChkCodeFG"
        Me.ChkCodeFG.Size = New System.Drawing.Size(81, 17)
        Me.ChkCodeFG.TabIndex = 0
        Me.ChkCodeFG.Text = "Show Code"
        Me.ChkCodeFG.UseVisualStyleBackColor = True
        '
        'ChkShowZeroUsed
        '
        Me.ChkShowZeroUsed.AutoSize = True
        Me.ChkShowZeroUsed.Location = New System.Drawing.Point(51, 38)
        Me.ChkShowZeroUsed.Name = "ChkShowZeroUsed"
        Me.ChkShowZeroUsed.Size = New System.Drawing.Size(109, 17)
        Me.ChkShowZeroUsed.TabIndex = 1
        Me.ChkShowZeroUsed.Text = "Show 0 Used Qty"
        Me.ChkShowZeroUsed.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(463, 11)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(64, 21)
        Me.btnShow.TabIndex = 5
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Product Recap", "Product Detail", "Job Order Recap", "Job Order Detail", "Work Order Recap", "Work Order Detail"})
        Me.CboFormat.Location = New System.Drawing.Point(51, 11)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(140, 21)
        Me.CboFormat.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Format"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkLogDateFrom)
        Me.GroupBox1.Controls.Add(Me.dtpToBundlingDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpFromBundlingDate)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(197, 73)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Log Date"
        '
        'ChkLogDateFrom
        '
        Me.ChkLogDateFrom.AutoSize = True
        Me.ChkLogDateFrom.Checked = True
        Me.ChkLogDateFrom.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkLogDateFrom.Location = New System.Drawing.Point(8, 21)
        Me.ChkLogDateFrom.Name = "ChkLogDateFrom"
        Me.ChkLogDateFrom.Size = New System.Drawing.Size(49, 17)
        Me.ChkLogDateFrom.TabIndex = 0
        Me.ChkLogDateFrom.Text = "From"
        Me.ChkLogDateFrom.UseVisualStyleBackColor = True
        '
        'dtpFromBundlingDate
        '
        Me.dtpFromBundlingDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFromBundlingDate.CustomFormat = "dd MMM yyyy"
        Me.dtpFromBundlingDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromBundlingDate.Location = New System.Drawing.Point(66, 19)
        Me.dtpFromBundlingDate.Name = "dtpFromBundlingDate"
        Me.dtpFromBundlingDate.Size = New System.Drawing.Size(125, 20)
        Me.dtpFromBundlingDate.TabIndex = 1
        '
        'colMO_No
        '
        Me.colMO_No.HeaderText = "No"
        Me.colMO_No.Name = "colMO_No"
        Me.colMO_No.Width = 46
        '
        'dgvMO
        '
        Me.dgvMO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMO.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvMO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMO.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colMO_No, Me.colMO_DocNo, Me.colMO_Date, Me.colMO_Primary})
        Me.dgvMO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvMO.Location = New System.Drawing.Point(3, 16)
        Me.dgvMO.Name = "dgvMO"
        Me.dgvMO.RowHeadersWidth = 24
        Me.dgvMO.RowTemplate.Height = 24
        Me.dgvMO.Size = New System.Drawing.Size(191, 182)
        Me.dgvMO.TabIndex = 0
        '
        'colMO_DocNo
        '
        Me.colMO_DocNo.HeaderText = "MO No"
        Me.colMO_DocNo.Name = "colMO_DocNo"
        Me.colMO_DocNo.Width = 66
        '
        'colMO_Date
        '
        Me.colMO_Date.HeaderText = "Date"
        Me.colMO_Date.Name = "colMO_Date"
        Me.colMO_Date.Width = 55
        '
        'colMO_Primary
        '
        Me.colMO_Primary.HeaderText = "Primary"
        Me.colMO_Primary.Name = "colMO_Primary"
        Me.colMO_Primary.Visible = False
        Me.colMO_Primary.Width = 66
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.dgvMO)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 357)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(197, 201)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Filter MO"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crv)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox5)
        Me.SplitContainer1.Size = New System.Drawing.Size(884, 561)
        Me.SplitContainer1.SplitterDistance = 203
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.ChkMODate)
        Me.GroupBox3.Controls.Add(Me.dtpToMODate)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.dtpFromMODate)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Location = New System.Drawing.Point(3, 85)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(197, 73)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'ChkMODate
        '
        Me.ChkMODate.AutoSize = True
        Me.ChkMODate.Location = New System.Drawing.Point(6, 0)
        Me.ChkMODate.Name = "ChkMODate"
        Me.ChkMODate.Size = New System.Drawing.Size(69, 17)
        Me.ChkMODate.TabIndex = 0
        Me.ChkMODate.Text = "MO Date"
        Me.ChkMODate.UseVisualStyleBackColor = True
        '
        'dtpToMODate
        '
        Me.dtpToMODate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpToMODate.CustomFormat = "dd MMM yyyy"
        Me.dtpToMODate.Enabled = False
        Me.dtpToMODate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToMODate.Location = New System.Drawing.Point(65, 45)
        Me.dtpToMODate.Name = "dtpToMODate"
        Me.dtpToMODate.Size = New System.Drawing.Size(125, 20)
        Me.dtpToMODate.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "To"
        '
        'dtpFromMODate
        '
        Me.dtpFromMODate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFromMODate.CustomFormat = "dd MMM yyyy"
        Me.dtpFromMODate.Enabled = False
        Me.dtpFromMODate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromMODate.Location = New System.Drawing.Point(66, 19)
        Me.dtpFromMODate.Name = "dtpFromMODate"
        Me.dtpFromMODate.Size = New System.Drawing.Size(125, 20)
        Me.dtpFromMODate.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "From"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgvProduct)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 164)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(197, 190)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter Product FG"
        '
        'dgvProduct
        '
        Me.dgvProduct.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvProduct.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProduct.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colProduct_No, Me.colProduct_GroupLevel, Me.colProduct_Code, Me.colProduct_Desc, Me.colProduct_Primary})
        Me.dgvProduct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProduct.Location = New System.Drawing.Point(3, 16)
        Me.dgvProduct.Name = "dgvProduct"
        Me.dgvProduct.RowHeadersWidth = 24
        Me.dgvProduct.RowTemplate.Height = 24
        Me.dgvProduct.Size = New System.Drawing.Size(191, 171)
        Me.dgvProduct.TabIndex = 0
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker4
        '
        '
        'colProduct_No
        '
        Me.colProduct_No.HeaderText = "No"
        Me.colProduct_No.Name = "colProduct_No"
        Me.colProduct_No.Width = 46
        '
        'colProduct_GroupLevel
        '
        Me.colProduct_GroupLevel.HeaderText = "Level"
        Me.colProduct_GroupLevel.Items.AddRange(New Object() {"Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.colProduct_GroupLevel.Name = "colProduct_GroupLevel"
        Me.colProduct_GroupLevel.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colProduct_GroupLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colProduct_GroupLevel.Width = 58
        '
        'colProduct_Code
        '
        Me.colProduct_Code.HeaderText = "Code"
        Me.colProduct_Code.Name = "colProduct_Code"
        Me.colProduct_Code.Width = 57
        '
        'colProduct_Desc
        '
        Me.colProduct_Desc.HeaderText = "Desc"
        Me.colProduct_Desc.Name = "colProduct_Desc"
        Me.colProduct_Desc.Width = 57
        '
        'colProduct_Primary
        '
        Me.colProduct_Primary.HeaderText = "Primary"
        Me.colProduct_Primary.Name = "colProduct_Primary"
        Me.colProduct_Primary.Visible = False
        Me.colProduct_Primary.Width = 66
        '
        'FrmLaporanProd_RawMaterialUsage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 561)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "FrmLaporanProd_RawMaterialUsage"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Raw Material Usage Report"
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvMO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgvProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dtpToBundlingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkIncludeWIP As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents crv As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents colMO_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvMO As System.Windows.Forms.DataGridView
    Friend WithEvents colMO_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colMO_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colMO_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvProduct As System.Windows.Forms.DataGridView
    Friend WithEvents dtpFromBundlingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkMODate As System.Windows.Forms.CheckBox
    Friend WithEvents dtpToMODate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dtpFromMODate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkLogDateFrom As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowZeroUsed As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDescRM As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeRM As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDescFG As System.Windows.Forms.CheckBox
    Friend WithEvents ChkCodeFG As System.Windows.Forms.CheckBox
    Friend WithEvents colProduct_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_GroupLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colProduct_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProduct_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
