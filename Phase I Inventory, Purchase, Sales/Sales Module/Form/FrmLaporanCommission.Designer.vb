<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanCommission
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
        Me.CRVfaktur = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.llbSorting = New System.Windows.Forms.LinkLabel
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ComboBox4 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.btnSettings = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnFilter = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.ColFilter_CommisionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_CommissionDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_PrimaryCommission = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvWarehouse = New System.Windows.Forms.DataGridView
        Me.ColFilter_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColFilter_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_ProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFilter_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CRVfaktur
        '
        Me.CRVfaktur.ActiveViewIndex = -1
        Me.CRVfaktur.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CRVfaktur.DisplayGroupTree = False
        Me.CRVfaktur.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CRVfaktur.Location = New System.Drawing.Point(0, 0)
        Me.CRVfaktur.Margin = New System.Windows.Forms.Padding(4)
        Me.CRVfaktur.Name = "CRVfaktur"
        Me.CRVfaktur.SelectionFormula = ""
        Me.CRVfaktur.ShowCloseButton = False
        Me.CRVfaktur.ShowGroupTreeButton = False
        Me.CRVfaktur.ShowRefreshButton = False
        Me.CRVfaktur.Size = New System.Drawing.Size(968, 405)
        Me.CRVfaktur.TabIndex = 2
        Me.CRVfaktur.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Button1.Location = New System.Drawing.Point(476, 421)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 28)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(12, 12)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer4)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.CRVfaktur)
        Me.SplitContainer2.Size = New System.Drawing.Size(968, 609)
        Me.SplitContainer2.SplitterDistance = 200
        Me.SplitContainer2.TabIndex = 49
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.llbSorting)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ComboBox3)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ComboBox4)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ComboBox2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ComboBox1)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnSettings)
        Me.SplitContainer4.Panel1.Controls.Add(Me.dtpFrom)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer4.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.CheckBox4)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer4.Panel1.Controls.Add(Me.cboFormat)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer4.Size = New System.Drawing.Size(968, 200)
        Me.SplitContainer4.SplitterDistance = 482
        Me.SplitContainer4.TabIndex = 50
        '
        'llbSorting
        '
        Me.llbSorting.AutoSize = True
        Me.llbSorting.Location = New System.Drawing.Point(101, 152)
        Me.llbSorting.Name = "llbSorting"
        Me.llbSorting.Size = New System.Drawing.Size(367, 17)
        Me.llbSorting.TabIndex = 19
        Me.llbSorting.TabStop = True
        Me.llbSorting.Text = "Detail Sorting: Date, Document No, Code No, Description"
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Code No"})
        Me.ComboBox3.Location = New System.Drawing.Point(311, 94)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(168, 24)
        Me.ComboBox3.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 17)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Product Level"
        '
        'ComboBox4
        '
        Me.ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Code No"})
        Me.ComboBox4.Location = New System.Drawing.Point(104, 94)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(175, 24)
        Me.ComboBox4.TabIndex = 13
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"n/a", "Commission", "Sales Person", "Product"})
        Me.ComboBox2.Location = New System.Drawing.Point(311, 64)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(168, 24)
        Me.ComboBox2.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 17)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Grouping"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Commission", "Sales Person", "Product"})
        Me.ComboBox1.Location = New System.Drawing.Point(104, 64)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(175, 24)
        Me.ComboBox1.TabIndex = 10
        '
        'btnSettings
        '
        Me.btnSettings.Location = New System.Drawing.Point(172, 124)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(62, 25)
        Me.btnSettings.TabIndex = 9
        Me.btnSettings.Text = "Settings"
        Me.btnSettings.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(104, 3)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(175, 22)
        Me.dtpFrom.TabIndex = 5
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(104, 124)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(62, 25)
        Me.btnFilter.TabIndex = 8
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Sales Period"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd MMM yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(311, 3)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(168, 22)
        Me.DateTimePicker1.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(285, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "to"
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Checked = True
        Me.CheckBox4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox4.Location = New System.Drawing.Point(311, 33)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(135, 21)
        Me.CheckBox4.TabIndex = 18
        Me.CheckBox4.Text = "Calculate Return"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 34)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 17)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Simple", "With Detail Item"})
        Me.cboFormat.Location = New System.Drawing.Point(104, 31)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(175, 24)
        Me.cboFormat.TabIndex = 16
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(482, 200)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(474, 171)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Sales Person Filter"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(468, 165)
        Me.DataGridView1.TabIndex = 2
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Code"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Primary"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(474, 171)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Commission Filter"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilter_CommisionCode, Me.ColFilter_CommissionDesc, Me.ColFilter_PrimaryCommission})
        Me.DataGridView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView2.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(468, 165)
        Me.DataGridView2.TabIndex = 2
        '
        'ColFilter_CommisionCode
        '
        Me.ColFilter_CommisionCode.HeaderText = "Code"
        Me.ColFilter_CommisionCode.Name = "ColFilter_CommisionCode"
        '
        'ColFilter_CommissionDesc
        '
        Me.ColFilter_CommissionDesc.HeaderText = "Description"
        Me.ColFilter_CommissionDesc.Name = "ColFilter_CommissionDesc"
        '
        'ColFilter_PrimaryCommission
        '
        Me.ColFilter_PrimaryCommission.HeaderText = "Primary"
        Me.ColFilter_PrimaryCommission.Name = "ColFilter_PrimaryCommission"
        Me.ColFilter_PrimaryCommission.Visible = False
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvWarehouse)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(474, 171)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Product Filter"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvWarehouse
        '
        Me.dgvWarehouse.AllowUserToResizeColumns = False
        Me.dgvWarehouse.AllowUserToResizeRows = False
        Me.dgvWarehouse.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvWarehouse.ColumnHeadersHeight = 25
        Me.dgvWarehouse.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvWarehouse.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColFilter_Level, Me.ColFilter_CodeNo, Me.ColFilter_ProductDesc, Me.ColFilter_PrimaryAcuan})
        Me.dgvWarehouse.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvWarehouse.Location = New System.Drawing.Point(3, 3)
        Me.dgvWarehouse.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvWarehouse.MultiSelect = False
        Me.dgvWarehouse.Name = "dgvWarehouse"
        Me.dgvWarehouse.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvWarehouse.RowTemplate.Height = 24
        Me.dgvWarehouse.Size = New System.Drawing.Size(468, 165)
        Me.dgvWarehouse.TabIndex = 11
        '
        'ColFilter_Level
        '
        Me.ColFilter_Level.Frozen = True
        Me.ColFilter_Level.HeaderText = "Level"
        Me.ColFilter_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColFilter_Level.Name = "ColFilter_Level"
        Me.ColFilter_Level.ReadOnly = True
        Me.ColFilter_Level.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColFilter_Level.Width = 67
        '
        'ColFilter_CodeNo
        '
        Me.ColFilter_CodeNo.Frozen = True
        Me.ColFilter_CodeNo.HeaderText = "Code"
        Me.ColFilter_CodeNo.Name = "ColFilter_CodeNo"
        Me.ColFilter_CodeNo.ReadOnly = True
        Me.ColFilter_CodeNo.Width = 66
        '
        'ColFilter_ProductDesc
        '
        Me.ColFilter_ProductDesc.HeaderText = "Description"
        Me.ColFilter_ProductDesc.Name = "ColFilter_ProductDesc"
        Me.ColFilter_ProductDesc.ReadOnly = True
        Me.ColFilter_ProductDesc.Width = 104
        '
        'ColFilter_PrimaryAcuan
        '
        Me.ColFilter_PrimaryAcuan.HeaderText = "Primary Acuan"
        Me.ColFilter_PrimaryAcuan.Name = "ColFilter_PrimaryAcuan"
        Me.ColFilter_PrimaryAcuan.ReadOnly = True
        Me.ColFilter_PrimaryAcuan.Visible = False
        Me.ColFilter_PrimaryAcuan.Width = 125
        '
        'FrmLaporanCommission
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Controls.Add(Me.Button1)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanCommission"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Commission Report"
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvWarehouse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CRVfaktur As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents btnSettings As System.Windows.Forms.Button
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilter_CommisionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_CommissionDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_PrimaryCommission As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvWarehouse As System.Windows.Forms.DataGridView
    Friend WithEvents ColFilter_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColFilter_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_ProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFilter_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbSorting As System.Windows.Forms.LinkLabel
End Class
