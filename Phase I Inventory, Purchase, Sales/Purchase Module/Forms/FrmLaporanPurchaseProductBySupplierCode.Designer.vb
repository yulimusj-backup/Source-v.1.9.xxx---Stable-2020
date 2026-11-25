<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanPurchaseProductBySupplierCode
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
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSupplierName = New System.Windows.Forms.TextBox
        Me.btnFindSupplier = New System.Windows.Forms.Button
        Me.txtPrimariSupplier = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.crvReport = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnFilter = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 13)
        Me.Label4.TabIndex = 336
        Me.Label4.Text = "Supplier Name"
        '
        'txtSupplierName
        '
        Me.txtSupplierName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierName.Location = New System.Drawing.Point(90, 64)
        Me.txtSupplierName.MaxLength = 50
        Me.txtSupplierName.Name = "txtSupplierName"
        Me.txtSupplierName.ReadOnly = True
        Me.txtSupplierName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtSupplierName.Size = New System.Drawing.Size(330, 20)
        Me.txtSupplierName.TabIndex = 4
        Me.txtSupplierName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnFindSupplier
        '
        Me.btnFindSupplier.Location = New System.Drawing.Point(388, 38)
        Me.btnFindSupplier.Name = "btnFindSupplier"
        Me.btnFindSupplier.Size = New System.Drawing.Size(32, 20)
        Me.btnFindSupplier.TabIndex = 3
        Me.btnFindSupplier.TabStop = False
        Me.btnFindSupplier.Text = "..."
        Me.btnFindSupplier.UseVisualStyleBackColor = True
        '
        'txtPrimariSupplier
        '
        Me.txtPrimariSupplier.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimariSupplier.Location = New System.Drawing.Point(90, 38)
        Me.txtPrimariSupplier.Name = "txtPrimariSupplier"
        Me.txtPrimariSupplier.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimariSupplier.Size = New System.Drawing.Size(27, 20)
        Me.txtPrimariSupplier.TabIndex = 332
        Me.txtPrimariSupplier.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimariSupplier.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 334
        Me.Label2.Text = "Supplier Code"
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSupplierCode.Location = New System.Drawing.Point(90, 38)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtSupplierCode.Size = New System.Drawing.Size(292, 20)
        Me.txtSupplierCode.TabIndex = 2
        Me.txtSupplierCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(244, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 330
        Me.Label1.Text = "to: "
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
        'crvReport
        '
        Me.crvReport.ActiveViewIndex = -1
        Me.crvReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReport.DisplayGroupTree = False
        Me.crvReport.Location = New System.Drawing.Point(12, 90)
        Me.crvReport.Name = "crvReport"
        Me.crvReport.SelectionFormula = ""
        Me.crvReport.ShowCloseButton = False
        Me.crvReport.ShowRefreshButton = False
        Me.crvReport.Size = New System.Drawing.Size(878, 399)
        Me.crvReport.TabIndex = 7
        Me.crvReport.ViewTimeSelectionFormula = ""
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(426, 11)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(86, 21)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.TabStop = False
        Me.btnFilter.Text = "&Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
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
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(426, 38)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 21)
        Me.btnClose.TabIndex = 6
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 13)
        Me.Label3.TabIndex = 325
        Me.Label3.Text = "AOP Date from:"
        '
        'FrmLaporanPurchaseProductBySupplierCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(902, 501)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSupplierName)
        Me.Controls.Add(Me.btnFindSupplier)
        Me.Controls.Add(Me.txtPrimariSupplier)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSupplierCode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.crvReport)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label3)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(918, 537)
        Me.Name = "FrmLaporanPurchaseProductBySupplierCode"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Puchase Product By Supplier"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindSupplier As System.Windows.Forms.Button
    Friend WithEvents txtPrimariSupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents crvReport As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
