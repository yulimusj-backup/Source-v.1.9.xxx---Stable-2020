<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanAkunProduct
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
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnClose = New System.Windows.Forms.Button
        Me.ChkGoods = New System.Windows.Forms.CheckBox
        Me.ChkService = New System.Windows.Forms.CheckBox
        Me.CboProduct_CodeDesc = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(13, 39)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(963, 577)
        Me.crv1.TabIndex = 48
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(466, 236)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(101, 28)
        Me.btnClose.TabIndex = 47
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ChkGoods
        '
        Me.ChkGoods.AutoSize = True
        Me.ChkGoods.Checked = True
        Me.ChkGoods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkGoods.Location = New System.Drawing.Point(21, 12)
        Me.ChkGoods.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkGoods.Name = "ChkGoods"
        Me.ChkGoods.Size = New System.Drawing.Size(72, 21)
        Me.ChkGoods.TabIndex = 49
        Me.ChkGoods.Text = "Goods"
        Me.ChkGoods.UseVisualStyleBackColor = True
        '
        'ChkService
        '
        Me.ChkService.AutoSize = True
        Me.ChkService.Checked = True
        Me.ChkService.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkService.Location = New System.Drawing.Point(99, 12)
        Me.ChkService.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkService.Name = "ChkService"
        Me.ChkService.Size = New System.Drawing.Size(77, 21)
        Me.ChkService.TabIndex = 50
        Me.ChkService.Text = "Service"
        Me.ChkService.UseVisualStyleBackColor = True
        '
        'CboProduct_CodeDesc
        '
        Me.CboProduct_CodeDesc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboProduct_CodeDesc.FormattingEnabled = True
        Me.CboProduct_CodeDesc.Items.AddRange(New Object() {"Show Code No Only", "Show Description Only", "Show Code No And Description"})
        Me.CboProduct_CodeDesc.Location = New System.Drawing.Point(181, 9)
        Me.CboProduct_CodeDesc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboProduct_CodeDesc.Name = "CboProduct_CodeDesc"
        Me.CboProduct_CodeDesc.Size = New System.Drawing.Size(253, 24)
        Me.CboProduct_CodeDesc.TabIndex = 51
        '
        'FrmLaporanAkunProduct
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.CboProduct_CodeDesc)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.ChkService)
        Me.Controls.Add(Me.ChkGoods)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanAkunProduct"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Account Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ChkGoods As System.Windows.Forms.CheckBox
    Friend WithEvents ChkService As System.Windows.Forms.CheckBox
    Friend WithEvents CboProduct_CodeDesc As System.Windows.Forms.ComboBox
End Class
