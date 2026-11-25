<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanAdjustmentList
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
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.clbBrand = New System.Windows.Forms.CheckedListBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnExit = New System.Windows.Forms.Button
        Me.cboTipeAdjustment = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkAdjust = New System.Windows.Forms.CheckBox
        Me.chkBundling = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel
        Me.clbAdjustmentType = New System.Windows.Forms.CheckedListBox
        Me.gboBrand.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(7, 180)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(910, 336)
        Me.CrystalReportViewer1.TabIndex = 7
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'gboBrand
        '
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbBrand)
        Me.gboBrand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboBrand.Location = New System.Drawing.Point(2, 2)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Size = New System.Drawing.Size(610, 132)
        Me.gboBrand.TabIndex = 0
        Me.gboBrand.TabStop = False
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(523, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(465, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
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
        Me.clbBrand.Size = New System.Drawing.Size(604, 109)
        Me.clbBrand.TabIndex = 0
        Me.clbBrand.ThreeDCheckBoxes = True
        '
        'btnFilter
        '
        Me.btnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter.Location = New System.Drawing.Point(818, 11)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(86, 28)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(628, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 300
        Me.Label3.Text = "Period"
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(671, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(141, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(818, 45)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(86, 28)
        Me.btnExit.TabIndex = 6
        Me.btnExit.TabStop = False
        Me.btnExit.Text = "&Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'cboTipeAdjustment
        '
        Me.cboTipeAdjustment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTipeAdjustment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipeAdjustment.FormattingEnabled = True
        Me.cboTipeAdjustment.Items.AddRange(New Object() {"In Adjustment", "Out Adjustment"})
        Me.cboTipeAdjustment.Location = New System.Drawing.Point(671, 38)
        Me.cboTipeAdjustment.Name = "cboTipeAdjustment"
        Me.cboTipeAdjustment.Size = New System.Drawing.Size(141, 21)
        Me.cboTipeAdjustment.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(628, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 302
        Me.Label1.Text = "Type"
        '
        'chkAdjust
        '
        Me.chkAdjust.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAdjust.AutoSize = True
        Me.chkAdjust.Checked = True
        Me.chkAdjust.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAdjust.Location = New System.Drawing.Point(633, 65)
        Me.chkAdjust.Name = "chkAdjust"
        Me.chkAdjust.Size = New System.Drawing.Size(154, 17)
        Me.chkAdjust.TabIndex = 3
        Me.chkAdjust.Text = "Adjustment && Stock Taking"
        Me.chkAdjust.UseVisualStyleBackColor = True
        '
        'chkBundling
        '
        Me.chkBundling.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkBundling.AutoSize = True
        Me.chkBundling.Checked = True
        Me.chkBundling.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBundling.Location = New System.Drawing.Point(633, 88)
        Me.chkBundling.Name = "chkBundling"
        Me.chkBundling.Size = New System.Drawing.Size(67, 17)
        Me.chkBundling.TabIndex = 4
        Me.chkBundling.Text = "Bundling"
        Me.chkBundling.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 12)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(622, 162)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.gboBrand)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage1.Size = New System.Drawing.Size(614, 136)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Brand Filter"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage2.Size = New System.Drawing.Size(614, 136)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Adjustment Type"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LinkLabel3)
        Me.GroupBox1.Controls.Add(Me.LinkLabel4)
        Me.GroupBox1.Controls.Add(Me.clbAdjustmentType)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(610, 132)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(523, 0)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel3.TabIndex = 2
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Uncheck All"
        '
        'LinkLabel4
        '
        Me.LinkLabel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel4.AutoSize = True
        Me.LinkLabel4.Location = New System.Drawing.Point(465, 0)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel4.TabIndex = 1
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "Check All"
        '
        'clbAdjustmentType
        '
        Me.clbAdjustmentType.CheckOnClick = True
        Me.clbAdjustmentType.ColumnWidth = 300
        Me.clbAdjustmentType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbAdjustmentType.FormattingEnabled = True
        Me.clbAdjustmentType.Location = New System.Drawing.Point(3, 16)
        Me.clbAdjustmentType.MultiColumn = True
        Me.clbAdjustmentType.Name = "clbAdjustmentType"
        Me.clbAdjustmentType.Size = New System.Drawing.Size(604, 109)
        Me.clbAdjustmentType.TabIndex = 0
        Me.clbAdjustmentType.ThreeDCheckBoxes = True
        '
        'frmLaporanAdjustmentList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(926, 518)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.chkBundling)
        Me.Controls.Add(Me.chkAdjust)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboTipeAdjustment)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.MinimumSize = New System.Drawing.Size(942, 556)
        Me.Name = "frmLaporanAdjustmentList"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Laporan Adjustment List"
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbBrand As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents cboTipeAdjustment As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkAdjust As System.Windows.Forms.CheckBox
    Friend WithEvents chkBundling As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbAdjustmentType As System.Windows.Forms.CheckedListBox
End Class
