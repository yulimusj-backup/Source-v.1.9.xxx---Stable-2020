<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseStatisticBySupplier
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.rboLocalSupplier = New System.Windows.Forms.RadioButton
        Me.rboForeignSupplier = New System.Windows.Forms.RadioButton
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.rboAll = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(565, 14)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(115, 26)
        Me.btnClose.TabIndex = 305
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(443, 14)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(115, 26)
        Me.btnFilter.TabIndex = 304
        Me.btnFilter.TabStop = False
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 18)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 17)
        Me.Label3.TabIndex = 303
        Me.Label3.Text = "Periode"
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(123, 15)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(311, 22)
        Me.dtpFrom.TabIndex = 302
        '
        'rboLocalSupplier
        '
        Me.rboLocalSupplier.AutoSize = True
        Me.rboLocalSupplier.Location = New System.Drawing.Point(123, 47)
        Me.rboLocalSupplier.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rboLocalSupplier.Name = "rboLocalSupplier"
        Me.rboLocalSupplier.Size = New System.Drawing.Size(119, 21)
        Me.rboLocalSupplier.TabIndex = 306
        Me.rboLocalSupplier.Text = "Local Supplier"
        Me.rboLocalSupplier.UseVisualStyleBackColor = True
        Me.rboLocalSupplier.Visible = False
        '
        'rboForeignSupplier
        '
        Me.rboForeignSupplier.AutoSize = True
        Me.rboForeignSupplier.Location = New System.Drawing.Point(300, 47)
        Me.rboForeignSupplier.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rboForeignSupplier.Name = "rboForeignSupplier"
        Me.rboForeignSupplier.Size = New System.Drawing.Size(133, 21)
        Me.rboForeignSupplier.TabIndex = 307
        Me.rboForeignSupplier.Text = "Foreign Supplier"
        Me.rboForeignSupplier.UseVisualStyleBackColor = True
        Me.rboForeignSupplier.Visible = False
        '
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.Location = New System.Drawing.Point(16, 47)
        Me.crvReport.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowGroupTreeButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(1170, 555)
        Me.crvReport.TabIndex = 308
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'rboAll
        '
        Me.rboAll.AutoSize = True
        Me.rboAll.Checked = True
        Me.rboAll.Location = New System.Drawing.Point(487, 47)
        Me.rboAll.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rboAll.Name = "rboAll"
        Me.rboAll.Size = New System.Drawing.Size(100, 21)
        Me.rboAll.TabIndex = 309
        Me.rboAll.TabStop = True
        Me.rboAll.Text = "All Supplier"
        Me.rboAll.UseVisualStyleBackColor = True
        Me.rboAll.Visible = False
        '
        'FrmLaporanPurchaseStatisticBySupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(1203, 617)
        Me.Controls.Add(Me.crvReport)
        Me.Controls.Add(Me.rboAll)
        Me.Controls.Add(Me.rboForeignSupplier)
        Me.Controls.Add(Me.rboLocalSupplier)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpFrom)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1218, 650)
        Me.Name = "FrmLaporanPurchaseStatisticBySupplier"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Statistic by Supplier"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents rboLocalSupplier As System.Windows.Forms.RadioButton
    Friend WithEvents rboForeignSupplier As System.Windows.Forms.RadioButton
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents rboAll As System.Windows.Forms.RadioButton
End Class
