<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseBySupplier
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
        Me.clbSupplier = New System.Windows.Forms.CheckedListBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.gboBrand = New System.Windows.Forms.GroupBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.gboBrand.SuspendLayout()
        Me.SuspendLayout()
        '
        'clbSupplier
        '
        Me.clbSupplier.CheckOnClick = True
        Me.clbSupplier.ColumnWidth = 300
        Me.clbSupplier.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSupplier.FormattingEnabled = True
        Me.clbSupplier.Location = New System.Drawing.Point(3, 16)
        Me.clbSupplier.MultiColumn = True
        Me.clbSupplier.Name = "clbSupplier"
        Me.clbSupplier.Size = New System.Drawing.Size(869, 79)
        Me.clbSupplier.TabIndex = 0
        Me.clbSupplier.ThreeDCheckBoxes = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(688, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 21)
        Me.btnClose.TabIndex = 4
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(728, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Check All"
        '
        'gboBrand
        '
        Me.gboBrand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboBrand.Controls.Add(Me.LinkLabel2)
        Me.gboBrand.Controls.Add(Me.LinkLabel1)
        Me.gboBrand.Controls.Add(Me.clbSupplier)
        Me.gboBrand.Location = New System.Drawing.Point(15, 38)
        Me.gboBrand.Name = "gboBrand"
        Me.gboBrand.Size = New System.Drawing.Size(875, 106)
        Me.gboBrand.TabIndex = 5
        Me.gboBrand.TabStop = False
        Me.gboBrand.Text = "Supplier"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(786, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel2.TabIndex = 2
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Uncheck All"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 303
        Me.Label3.Text = "Periode from: "
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(90, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(148, 20)
        Me.dtpFrom.TabIndex = 0
        '
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.Location = New System.Drawing.Point(12, 150)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(878, 339)
        Me.crvReport.TabIndex = 6
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(596, 8)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(86, 21)
        Me.btnFilter.TabIndex = 3
        Me.btnFilter.TabStop = False
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(272, 12)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(148, 20)
        Me.dtpTo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(244, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 309
        Me.Label1.Text = "to: "
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(426, 11)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(164, 17)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Use Supplier Invoice Number"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FrmLaporanPurchaseBySupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(902, 501)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.gboBrand)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.crvReport)
        Me.Controls.Add(Me.btnFilter)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(918, 537)
        Me.Name = "FrmLaporanPurchaseBySupplier"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Puchase By Supplier"
        Me.gboBrand.ResumeLayout(False)
        Me.gboBrand.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents clbSupplier As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents gboBrand As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
