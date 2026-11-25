<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseRequisition
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
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtBrandName = New System.Windows.Forms.TextBox
        Me.btnFindBrand = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtBrandCode = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtRoundingPR = New System.Windows.Forms.TextBox
        Me.txtRoundingAvgPerMonth = New System.Windows.Forms.TextBox
        Me.txtFormatYears = New System.Windows.Forms.TextBox
        Me.txtFormatDay = New System.Windows.Forms.TextBox
        Me.txtFormatMonth = New System.Windows.Forms.TextBox
        Me.txtPrimaryBrand = New System.Windows.Forms.TextBox
        Me.txtFormatDecimal = New System.Windows.Forms.TextBox
        Me.txtFormatLeadTime = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 336
        Me.Label4.Text = "Description"
        '
        'txtBrandName
        '
        Me.txtBrandName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBrandName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBrandName.Location = New System.Drawing.Point(111, 38)
        Me.txtBrandName.MaxLength = 50
        Me.txtBrandName.Name = "txtBrandName"
        Me.txtBrandName.ReadOnly = True
        Me.txtBrandName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtBrandName.Size = New System.Drawing.Size(330, 20)
        Me.txtBrandName.TabIndex = 2
        Me.txtBrandName.TabStop = False
        Me.txtBrandName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnFindBrand
        '
        Me.btnFindBrand.Location = New System.Drawing.Point(409, 12)
        Me.btnFindBrand.Name = "btnFindBrand"
        Me.btnFindBrand.Size = New System.Drawing.Size(32, 20)
        Me.btnFindBrand.TabIndex = 1
        Me.btnFindBrand.TabStop = False
        Me.btnFindBrand.Text = "..."
        Me.btnFindBrand.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 334
        Me.Label2.Text = "Brand Code"
        '
        'txtBrandCode
        '
        Me.txtBrandCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBrandCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBrandCode.Location = New System.Drawing.Point(111, 12)
        Me.txtBrandCode.MaxLength = 50
        Me.txtBrandCode.Name = "txtBrandCode"
        Me.txtBrandCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBrandCode.Size = New System.Drawing.Size(292, 20)
        Me.txtBrandCode.TabIndex = 0
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(447, 11)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(86, 21)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(447, 63)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 21)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(263, 67)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(28, 13)
        Me.Label1.TabIndex = 340
        Me.Label1.Text = "thru "
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(293, 64)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.ShowUpDown = True
        Me.dtpTo.Size = New System.Drawing.Size(148, 20)
        Me.dtpTo.TabIndex = 4
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(111, 64)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(148, 20)
        Me.dtpFrom.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 13)
        Me.Label3.TabIndex = 338
        Me.Label3.Text = "Reference Period:"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(14, 90)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(875, 399)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(867, 373)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data Viewer"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(861, 367)
        Me.DataGridView1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crvReport)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(867, 373)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Report Viewer"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.DisplayStatusBar = False
        Me.crvReport.DisplayToolbar = False
        Me.crvReport.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReport.Location = New System.Drawing.Point(3, 3)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(861, 367)
        Me.crvReport.TabIndex = 329
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'txtRoundingPR
        '
        Me.txtRoundingPR.Location = New System.Drawing.Point(555, 38)
        Me.txtRoundingPR.Name = "txtRoundingPR"
        Me.txtRoundingPR.Size = New System.Drawing.Size(100, 20)
        Me.txtRoundingPR.TabIndex = 342
        Me.txtRoundingPR.TabStop = False
        Me.txtRoundingPR.Text = "0"
        Me.txtRoundingPR.Visible = False
        '
        'txtRoundingAvgPerMonth
        '
        Me.txtRoundingAvgPerMonth.Location = New System.Drawing.Point(555, 13)
        Me.txtRoundingAvgPerMonth.Name = "txtRoundingAvgPerMonth"
        Me.txtRoundingAvgPerMonth.Size = New System.Drawing.Size(100, 20)
        Me.txtRoundingAvgPerMonth.TabIndex = 341
        Me.txtRoundingAvgPerMonth.TabStop = False
        Me.txtRoundingAvgPerMonth.Text = "0"
        Me.txtRoundingAvgPerMonth.Visible = False
        '
        'txtFormatYears
        '
        Me.txtFormatYears.Location = New System.Drawing.Point(682, 62)
        Me.txtFormatYears.Name = "txtFormatYears"
        Me.txtFormatYears.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatYears.TabIndex = 346
        Me.txtFormatYears.TabStop = False
        Me.txtFormatYears.Text = "1"
        Me.txtFormatYears.Visible = False
        '
        'txtFormatDay
        '
        Me.txtFormatDay.Location = New System.Drawing.Point(682, 36)
        Me.txtFormatDay.Name = "txtFormatDay"
        Me.txtFormatDay.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDay.TabIndex = 345
        Me.txtFormatDay.TabStop = False
        Me.txtFormatDay.Text = "1"
        Me.txtFormatDay.Visible = False
        '
        'txtFormatMonth
        '
        Me.txtFormatMonth.Location = New System.Drawing.Point(682, 9)
        Me.txtFormatMonth.Name = "txtFormatMonth"
        Me.txtFormatMonth.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatMonth.TabIndex = 344
        Me.txtFormatMonth.TabStop = False
        Me.txtFormatMonth.Text = "3"
        Me.txtFormatMonth.Visible = False
        '
        'txtPrimaryBrand
        '
        Me.txtPrimaryBrand.Location = New System.Drawing.Point(111, 13)
        Me.txtPrimaryBrand.Name = "txtPrimaryBrand"
        Me.txtPrimaryBrand.Size = New System.Drawing.Size(58, 20)
        Me.txtPrimaryBrand.TabIndex = 347
        Me.txtPrimaryBrand.Text = "-1"
        Me.txtPrimaryBrand.Visible = False
        '
        'txtFormatDecimal
        '
        Me.txtFormatDecimal.Location = New System.Drawing.Point(555, 67)
        Me.txtFormatDecimal.Name = "txtFormatDecimal"
        Me.txtFormatDecimal.Size = New System.Drawing.Size(100, 20)
        Me.txtFormatDecimal.TabIndex = 348
        Me.txtFormatDecimal.TabStop = False
        Me.txtFormatDecimal.Text = "0"
        Me.txtFormatDecimal.Visible = False
        '
        'txtFormatLeadTime
        '
        Me.txtFormatLeadTime.Location = New System.Drawing.Point(449, 67)
        Me.txtFormatLeadTime.Name = "txtFormatLeadTime"
        Me.txtFormatLeadTime.Size = New System.Drawing.Size(100, 20)
        Me.txtFormatLeadTime.TabIndex = 349
        Me.txtFormatLeadTime.TabStop = False
        Me.txtFormatLeadTime.Text = "1"
        Me.txtFormatLeadTime.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(539, 15)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(111, 17)
        Me.CheckBox1.TabIndex = 350
        Me.CheckBox1.Text = "Negative PR Only"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Location = New System.Drawing.Point(447, 38)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(86, 21)
        Me.btnFindCodeNo.TabIndex = 351
        Me.btnFindCodeNo.Text = "Find Code No"
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'FrmLaporanPurchaseRequisition
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(902, 501)
        Me.Controls.Add(Me.btnFindCodeNo)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.txtFormatLeadTime)
        Me.Controls.Add(Me.txtFormatDecimal)
        Me.Controls.Add(Me.txtPrimaryBrand)
        Me.Controls.Add(Me.txtFormatYears)
        Me.Controls.Add(Me.txtFormatDay)
        Me.Controls.Add(Me.txtFormatMonth)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.txtRoundingPR)
        Me.Controls.Add(Me.txtRoundingAvgPerMonth)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtBrandName)
        Me.Controls.Add(Me.btnFindBrand)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtBrandCode)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.btnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(918, 537)
        Me.Name = "FrmLaporanPurchaseRequisition"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Requisition"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBrandName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindBrand As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtBrandCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents txtRoundingPR As System.Windows.Forms.TextBox
    Friend WithEvents txtRoundingAvgPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatYears As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDay As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatMonth As System.Windows.Forms.TextBox
    Public WithEvents txtPrimaryBrand As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDecimal As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
End Class
