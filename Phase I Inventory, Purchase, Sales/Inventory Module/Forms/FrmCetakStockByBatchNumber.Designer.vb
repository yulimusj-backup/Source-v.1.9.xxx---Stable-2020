<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCetakStockByBatchNumber
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
        Me.BtnCetak = New System.Windows.Forms.Button
        Me.TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.TxtCodeProduct = New System.Windows.Forms.TextBox
        Me.lbProductCode = New System.Windows.Forms.Label
        Me.BtnCari = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbFillByProductCode = New System.Windows.Forms.RadioButton
        Me.rbFillAll = New System.Windows.Forms.RadioButton
        Me.BtnClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.ShowCloseButton = False
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.ShowRefreshButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(541, 471)
        Me.CrystalReportViewer1.TabIndex = 0
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnCetak)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryProduct)
        Me.GroupBox1.Controls.Add(Me.TxtCodeProduct)
        Me.GroupBox1.Controls.Add(Me.lbProductCode)
        Me.GroupBox1.Controls.Add(Me.BtnCari)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(525, 65)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'BtnCetak
        '
        Me.BtnCetak.Location = New System.Drawing.Point(215, 40)
        Me.BtnCetak.Name = "BtnCetak"
        Me.BtnCetak.Size = New System.Drawing.Size(48, 20)
        Me.BtnCetak.TabIndex = 21
        Me.BtnCetak.Text = "&Cetak"
        Me.BtnCetak.UseVisualStyleBackColor = True
        '
        'TxtPrimaryProduct
        '
        Me.TxtPrimaryProduct.Location = New System.Drawing.Point(290, 19)
        Me.TxtPrimaryProduct.Name = "TxtPrimaryProduct"
        Me.TxtPrimaryProduct.Size = New System.Drawing.Size(10, 20)
        Me.TxtPrimaryProduct.TabIndex = 17
        Me.TxtPrimaryProduct.TabStop = False
        Me.TxtPrimaryProduct.Visible = False
        '
        'TxtCodeProduct
        '
        Me.TxtCodeProduct.Location = New System.Drawing.Point(290, 19)
        Me.TxtCodeProduct.Name = "TxtCodeProduct"
        Me.TxtCodeProduct.Size = New System.Drawing.Size(96, 20)
        Me.TxtCodeProduct.TabIndex = 16
        '
        'lbProductCode
        '
        Me.lbProductCode.AutoSize = True
        Me.lbProductCode.Location = New System.Drawing.Point(212, 19)
        Me.lbProductCode.Name = "lbProductCode"
        Me.lbProductCode.Size = New System.Drawing.Size(72, 13)
        Me.lbProductCode.TabIndex = 15
        Me.lbProductCode.Text = "Product Code"
        '
        'BtnCari
        '
        Me.BtnCari.Location = New System.Drawing.Point(392, 19)
        Me.BtnCari.Name = "BtnCari"
        Me.BtnCari.Size = New System.Drawing.Size(33, 20)
        Me.BtnCari.TabIndex = 13
        Me.BtnCari.Text = "Cari"
        Me.BtnCari.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbFillByProductCode)
        Me.GroupBox2.Controls.Add(Me.rbFillAll)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 49)
        Me.GroupBox2.TabIndex = 20
        Me.GroupBox2.TabStop = False
        '
        'rbFillByProductCode
        '
        Me.rbFillByProductCode.AutoSize = True
        Me.rbFillByProductCode.Checked = True
        Me.rbFillByProductCode.Location = New System.Drawing.Point(75, 19)
        Me.rbFillByProductCode.Name = "rbFillByProductCode"
        Me.rbFillByProductCode.Size = New System.Drawing.Size(119, 17)
        Me.rbFillByProductCode.TabIndex = 19
        Me.rbFillByProductCode.TabStop = True
        Me.rbFillByProductCode.Text = "Fill by Product Code"
        Me.rbFillByProductCode.UseVisualStyleBackColor = True
        '
        'rbFillAll
        '
        Me.rbFillAll.AutoSize = True
        Me.rbFillAll.Location = New System.Drawing.Point(6, 19)
        Me.rbFillAll.Name = "rbFillAll"
        Me.rbFillAll.Size = New System.Drawing.Size(51, 17)
        Me.rbFillAll.TabIndex = 18
        Me.rbFillAll.TabStop = True
        Me.rbFillAll.Text = "Fill All"
        Me.rbFillAll.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(450, 77)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(48, 20)
        Me.BtnClose.TabIndex = 22
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmCetakStockByBatchNumber
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(541, 471)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(557, 507)
        Me.Name = "FrmCetakStockByBatchNumber"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cetak Stock By Batch Number"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbFillByProductCode As System.Windows.Forms.RadioButton
    Friend WithEvents rbFillAll As System.Windows.Forms.RadioButton
    Friend WithEvents TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents TxtCodeProduct As System.Windows.Forms.TextBox
    Friend WithEvents lbProductCode As System.Windows.Forms.Label
    Friend WithEvents BtnCari As System.Windows.Forms.Button
    Friend WithEvents BtnCetak As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
