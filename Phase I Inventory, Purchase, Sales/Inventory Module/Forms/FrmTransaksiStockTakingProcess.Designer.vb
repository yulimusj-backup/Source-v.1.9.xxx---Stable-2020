<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockTakingProcess
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
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.crvExpiryDataCheck = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crvStockTakingReportDetail = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.txtFormatYears = New System.Windows.Forms.TextBox
        Me.btnRefreshReport = New System.Windows.Forms.Button
        Me.txtFormatDay = New System.Windows.Forms.TextBox
        Me.crvStockTakingReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.txtFormatMonth = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnStartProses = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.crvQtyDataCheck = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage5.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.crvExpiryDataCheck)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(905, 638)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Expiry Date Data Check"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'crvExpiryDataCheck
        '
        Me.crvExpiryDataCheck.ActiveViewIndex = -1
        Me.crvExpiryDataCheck.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvExpiryDataCheck.DisplayGroupTree = False
        Me.crvExpiryDataCheck.DisplayStatusBar = False
        Me.crvExpiryDataCheck.DisplayToolbar = False
        Me.crvExpiryDataCheck.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvExpiryDataCheck.Location = New System.Drawing.Point(3, 3)
        Me.crvExpiryDataCheck.Name = "crvExpiryDataCheck"
        Me.crvExpiryDataCheck.SelectionFormula = ""
        Me.crvExpiryDataCheck.ShowCloseButton = False
        Me.crvExpiryDataCheck.ShowRefreshButton = False
        Me.crvExpiryDataCheck.Size = New System.Drawing.Size(899, 632)
        Me.crvExpiryDataCheck.TabIndex = 11
        Me.crvExpiryDataCheck.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crvStockTakingReportDetail)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(905, 638)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Stock Taking Report-Detail"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crvStockTakingReportDetail
        '
        Me.crvStockTakingReportDetail.ActiveViewIndex = -1
        Me.crvStockTakingReportDetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvStockTakingReportDetail.DisplayGroupTree = False
        Me.crvStockTakingReportDetail.DisplayStatusBar = False
        Me.crvStockTakingReportDetail.DisplayToolbar = False
        Me.crvStockTakingReportDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvStockTakingReportDetail.Location = New System.Drawing.Point(3, 3)
        Me.crvStockTakingReportDetail.Name = "crvStockTakingReportDetail"
        Me.crvStockTakingReportDetail.SelectionFormula = ""
        Me.crvStockTakingReportDetail.ShowCloseButton = False
        Me.crvStockTakingReportDetail.ShowRefreshButton = False
        Me.crvStockTakingReportDetail.Size = New System.Drawing.Size(899, 632)
        Me.crvStockTakingReportDetail.TabIndex = 9
        Me.crvStockTakingReportDetail.ViewTimeSelectionFormula = ""
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.txtFormatYears)
        Me.TabPage1.Controls.Add(Me.btnRefreshReport)
        Me.TabPage1.Controls.Add(Me.txtFormatDay)
        Me.TabPage1.Controls.Add(Me.crvStockTakingReport)
        Me.TabPage1.Controls.Add(Me.txtFormatMonth)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.btnStartProses)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(905, 638)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Stock Taking Report"
        '
        'txtFormatYears
        '
        Me.txtFormatYears.Location = New System.Drawing.Point(130, 149)
        Me.txtFormatYears.Name = "txtFormatYears"
        Me.txtFormatYears.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatYears.TabIndex = 308
        Me.txtFormatYears.TabStop = False
        Me.txtFormatYears.Text = "1"
        Me.txtFormatYears.Visible = False
        '
        'btnRefreshReport
        '
        Me.btnRefreshReport.Location = New System.Drawing.Point(102, 6)
        Me.btnRefreshReport.Name = "btnRefreshReport"
        Me.btnRefreshReport.Size = New System.Drawing.Size(95, 23)
        Me.btnRefreshReport.TabIndex = 13
        Me.btnRefreshReport.Text = "Refresh Report"
        Me.btnRefreshReport.UseVisualStyleBackColor = True
        '
        'txtFormatDay
        '
        Me.txtFormatDay.Location = New System.Drawing.Point(130, 123)
        Me.txtFormatDay.Name = "txtFormatDay"
        Me.txtFormatDay.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDay.TabIndex = 307
        Me.txtFormatDay.TabStop = False
        Me.txtFormatDay.Text = "1"
        Me.txtFormatDay.Visible = False
        '
        'crvStockTakingReport
        '
        Me.crvStockTakingReport.ActiveViewIndex = -1
        Me.crvStockTakingReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvStockTakingReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvStockTakingReport.DisplayGroupTree = False
        Me.crvStockTakingReport.DisplayStatusBar = False
        Me.crvStockTakingReport.DisplayToolbar = False
        Me.crvStockTakingReport.Location = New System.Drawing.Point(6, 219)
        Me.crvStockTakingReport.Name = "crvStockTakingReport"
        Me.crvStockTakingReport.SelectionFormula = ""
        Me.crvStockTakingReport.ShowCloseButton = False
        Me.crvStockTakingReport.ShowRefreshButton = False
        Me.crvStockTakingReport.Size = New System.Drawing.Size(896, 416)
        Me.crvStockTakingReport.TabIndex = 8
        Me.crvStockTakingReport.ViewTimeSelectionFormula = ""
        '
        'txtFormatMonth
        '
        Me.txtFormatMonth.Location = New System.Drawing.Point(130, 96)
        Me.txtFormatMonth.Name = "txtFormatMonth"
        Me.txtFormatMonth.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatMonth.TabIndex = 306
        Me.txtFormatMonth.TabStop = False
        Me.txtFormatMonth.Text = "3"
        Me.txtFormatMonth.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.LinkLabel2)
        Me.GroupBox2.Controls.Add(Me.LinkLabel1)
        Me.GroupBox2.Controls.Add(Me.CheckedListBox1)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Location = New System.Drawing.Point(203, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(694, 207)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Brand"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(606, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 4
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(550, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(3, 16)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(688, 184)
        Me.CheckedListBox1.TabIndex = 10
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(56, 16)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Button3"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(63, 362)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 34)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnStartProses
        '
        Me.btnStartProses.Location = New System.Drawing.Point(8, 6)
        Me.btnStartProses.Name = "btnStartProses"
        Me.btnStartProses.Size = New System.Drawing.Size(88, 23)
        Me.btnStartProses.TabIndex = 11
        Me.btnStartProses.Text = "Start Proses"
        Me.btnStartProses.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.HotTrack = True
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(913, 664)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.crvQtyDataCheck)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(905, 638)
        Me.TabPage3.TabIndex = 5
        Me.TabPage3.Text = "Location Data Check"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'crvQtyDataCheck
        '
        Me.crvQtyDataCheck.ActiveViewIndex = -1
        Me.crvQtyDataCheck.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvQtyDataCheck.DisplayGroupTree = False
        Me.crvQtyDataCheck.DisplayStatusBar = False
        Me.crvQtyDataCheck.DisplayToolbar = False
        Me.crvQtyDataCheck.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvQtyDataCheck.Location = New System.Drawing.Point(3, 3)
        Me.crvQtyDataCheck.Name = "crvQtyDataCheck"
        Me.crvQtyDataCheck.SelectionFormula = ""
        Me.crvQtyDataCheck.ShowCloseButton = False
        Me.crvQtyDataCheck.ShowRefreshButton = False
        Me.crvQtyDataCheck.Size = New System.Drawing.Size(899, 632)
        Me.crvQtyDataCheck.TabIndex = 11
        Me.crvQtyDataCheck.ViewTimeSelectionFormula = ""
        '
        'FrmTransaksiStockTakingProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(913, 664)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(929, 700)
        Me.Name = "FrmTransaksiStockTakingProcess"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "End Stock Taking"
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents crvStockTakingReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents crvStockTakingReportDetail As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents crvExpiryDataCheck As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crvQtyDataCheck As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnRefreshReport As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnStartProses As System.Windows.Forms.Button
    Friend WithEvents txtFormatYears As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDay As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatMonth As System.Windows.Forms.TextBox
End Class
