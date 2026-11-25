<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseOrderBalanceBySupplier
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSupplierName = New System.Windows.Forms.TextBox
        Me.txtFormatYears = New System.Windows.Forms.TextBox
        Me.txtFormatDay = New System.Windows.Forms.TextBox
        Me.txtFormatMonth = New System.Windows.Forms.TextBox
        Me.txtFormatDecimal = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkPeriod = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(238, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 317
        Me.Label1.Text = "to: "
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Enabled = False
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(266, 19)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(148, 20)
        Me.dtpTo.TabIndex = 1
        '
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.Location = New System.Drawing.Point(15, 116)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(878, 373)
        Me.crvReport.TabIndex = 7
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(444, 12)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(86, 21)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(84, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(148, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(444, 39)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 21)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 311
        Me.Label3.Text = "PO Date from:"
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(397, 64)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSupplier.TabIndex = 3
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "..."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(99, 64)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(27, 20)
        Me.txtPrimarySupplier.TabIndex = 319
        Me.txtPrimarySupplier.Text = "-1"
        Me.txtPrimarySupplier.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimarySupplier.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 321
        Me.Label2.Text = "Supplier Code"
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierCode.Location = New System.Drawing.Point(99, 64)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSupplierCode.Size = New System.Drawing.Size(292, 20)
        Me.txtSupplierCode.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(21, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 13)
        Me.Label4.TabIndex = 323
        Me.Label4.Text = "Supplier Name"
        '
        'txtSupplierName
        '
        Me.txtSupplierName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierName.Location = New System.Drawing.Point(99, 90)
        Me.txtSupplierName.MaxLength = 50
        Me.txtSupplierName.Name = "txtSupplierName"
        Me.txtSupplierName.ReadOnly = True
        Me.txtSupplierName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtSupplierName.Size = New System.Drawing.Size(330, 20)
        Me.txtSupplierName.TabIndex = 4
        Me.txtSupplierName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFormatYears
        '
        Me.txtFormatYears.Location = New System.Drawing.Point(649, 232)
        Me.txtFormatYears.Name = "txtFormatYears"
        Me.txtFormatYears.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatYears.TabIndex = 326
        Me.txtFormatYears.TabStop = False
        Me.txtFormatYears.Text = "1"
        Me.txtFormatYears.Visible = False
        '
        'txtFormatDay
        '
        Me.txtFormatDay.Location = New System.Drawing.Point(649, 206)
        Me.txtFormatDay.Name = "txtFormatDay"
        Me.txtFormatDay.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDay.TabIndex = 325
        Me.txtFormatDay.TabStop = False
        Me.txtFormatDay.Text = "1"
        Me.txtFormatDay.Visible = False
        '
        'txtFormatMonth
        '
        Me.txtFormatMonth.Location = New System.Drawing.Point(649, 179)
        Me.txtFormatMonth.Name = "txtFormatMonth"
        Me.txtFormatMonth.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatMonth.TabIndex = 324
        Me.txtFormatMonth.TabStop = False
        Me.txtFormatMonth.Text = "3"
        Me.txtFormatMonth.Visible = False
        '
        'txtFormatDecimal
        '
        Me.txtFormatDecimal.Location = New System.Drawing.Point(725, 179)
        Me.txtFormatDecimal.Name = "txtFormatDecimal"
        Me.txtFormatDecimal.Size = New System.Drawing.Size(58, 20)
        Me.txtFormatDecimal.TabIndex = 327
        Me.txtFormatDecimal.TabStop = False
        Me.txtFormatDecimal.Text = "2"
        Me.txtFormatDecimal.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpFrom)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpTo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(423, 47)
        Me.GroupBox1.TabIndex = 328
        Me.GroupBox1.TabStop = False
        '
        'chkPeriod
        '
        Me.chkPeriod.AutoSize = True
        Me.chkPeriod.Location = New System.Drawing.Point(27, 8)
        Me.chkPeriod.Name = "chkPeriod"
        Me.chkPeriod.Size = New System.Drawing.Size(96, 17)
        Me.chkPeriod.TabIndex = 329
        Me.chkPeriod.Text = "Filter By Period"
        Me.chkPeriod.UseVisualStyleBackColor = True
        '
        'FrmLaporanPurchaseOrderBalanceBySupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(902, 501)
        Me.Controls.Add(Me.chkPeriod)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.crvReport)
        Me.Controls.Add(Me.txtFormatDecimal)
        Me.Controls.Add(Me.txtFormatYears)
        Me.Controls.Add(Me.txtFormatDay)
        Me.Controls.Add(Me.txtFormatMonth)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSupplierName)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.txtPrimarySupplier)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSupplierCode)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.btnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(918, 537)
        Me.Name = "FrmLaporanPurchaseOrderBalanceBySupplier"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Balance By Supplier"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatYears As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDay As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatMonth As System.Windows.Forms.TextBox
    Friend WithEvents txtFormatDecimal As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkPeriod As System.Windows.Forms.CheckBox
End Class
