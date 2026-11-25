<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanHistoricalCostImport
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.TxtCodeProduct = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BtnFind = New System.Windows.Forms.Button
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbFillByProductCode = New System.Windows.Forms.RadioButton
        Me.rbFillAll = New System.Windows.Forms.RadioButton
        Me.lbProductCode = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnCetak = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
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
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(12, 82)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(775, 436)
        Me.CrystalReportViewer1.TabIndex = 1
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryProduct)
        Me.GroupBox1.Controls.Add(Me.TxtCodeProduct)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.BtnFind)
        Me.GroupBox1.Controls.Add(Me.DtpDate)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.lbProductCode)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.BtnCetak)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(775, 74)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(388, 45)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(10, 20)
        Me.TxtPrimaryProduct.TabIndex = 34
        Me.TxtPrimaryProduct.TabStop = False
        Me.TxtPrimaryProduct.Visible = False
        '
        'TxtCodeProduct
        '
        Me.TxtCodeProduct.Location = New System.Drawing.Point(388, 45)
        Me.TxtCodeProduct.Name = "TxtCodeProduct"
        Me.TxtCodeProduct.Size = New System.Drawing.Size(119, 20)
        Me.TxtCodeProduct.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(308, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Product Code"
        '
        'BtnFind
        '
        Me.BtnFind.Location = New System.Drawing.Point(513, 45)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(38, 20)
        Me.BtnFind.TabIndex = 3
        Me.BtnFind.TabStop = False
        Me.BtnFind.Text = "...."
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'DtpDate
        '
        Me.DtpDate.CustomFormat = "MMMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(351, 19)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.ShowUpDown = True
        Me.DtpDate.Size = New System.Drawing.Size(200, 20)
        Me.DtpDate.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbFillByProductCode)
        Me.GroupBox2.Controls.Add(Me.rbFillAll)
        Me.GroupBox2.Location = New System.Drawing.Point(24, 13)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 49)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'rbFillByProductCode
        '
        Me.rbFillByProductCode.AutoSize = True
        Me.rbFillByProductCode.Checked = True
        Me.rbFillByProductCode.Location = New System.Drawing.Point(76, 19)
        Me.rbFillByProductCode.Name = "rbFillByProductCode"
        Me.rbFillByProductCode.Size = New System.Drawing.Size(119, 17)
        Me.rbFillByProductCode.TabIndex = 1
        Me.rbFillByProductCode.TabStop = True
        Me.rbFillByProductCode.Text = "Fill by Product Code"
        Me.rbFillByProductCode.UseVisualStyleBackColor = True
        '
        'rbFillAll
        '
        Me.rbFillAll.AutoSize = True
        Me.rbFillAll.Location = New System.Drawing.Point(8, 19)
        Me.rbFillAll.Name = "rbFillAll"
        Me.rbFillAll.Size = New System.Drawing.Size(51, 17)
        Me.rbFillAll.TabIndex = 0
        Me.rbFillAll.TabStop = True
        Me.rbFillAll.Text = "Fill All"
        Me.rbFillAll.UseVisualStyleBackColor = True
        '
        'lbProductCode
        '
        Me.lbProductCode.AutoSize = True
        Me.lbProductCode.Location = New System.Drawing.Point(308, 21)
        Me.lbProductCode.Name = "lbProductCode"
        Me.lbProductCode.Size = New System.Drawing.Size(37, 13)
        Me.lbProductCode.TabIndex = 28
        Me.lbProductCode.Text = "Period"
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(557, 45)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 20)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "&Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BtnCetak
        '
        Me.BtnCetak.Location = New System.Drawing.Point(557, 19)
        Me.BtnCetak.Name = "BtnCetak"
        Me.BtnCetak.Size = New System.Drawing.Size(57, 20)
        Me.BtnCetak.TabIndex = 4
        Me.BtnCetak.Text = "&Preview"
        Me.BtnCetak.UseVisualStyleBackColor = True
        '
        'FrmLaporanHistoricalCostImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(789, 508)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(805, 544)
        Me.Name = "FrmLaporanHistoricalCostImport"
        Me.ShowInTaskbar = False
        Me.Text = "Laporan Historical Cost Import"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeProduct As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbFillByProductCode As System.Windows.Forms.RadioButton
    Friend WithEvents rbFillAll As System.Windows.Forms.RadioButton
    Friend WithEvents lbProductCode As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtnCetak As System.Windows.Forms.Button
End Class
