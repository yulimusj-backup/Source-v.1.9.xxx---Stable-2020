<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing_DetailCalc2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonthlyClosing_DetailCalc2))
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Button1 = New System.Windows.Forms.Button
        Me.chkUseSupplierInvoiceNumber = New System.Windows.Forms.CheckBox
        Me.chkShowDetailBegining = New System.Windows.Forms.CheckBox
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
        Me.crv1.Location = New System.Drawing.Point(0, 34)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(992, 599)
        Me.crv1.TabIndex = 3
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(392, 142)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(172, 34)
        Me.Button1.TabIndex = 4
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'chkUseSupplierInvoiceNumber
        '
        Me.chkUseSupplierInvoiceNumber.AutoSize = True
        Me.chkUseSupplierInvoiceNumber.Location = New System.Drawing.Point(12, 12)
        Me.chkUseSupplierInvoiceNumber.Name = "chkUseSupplierInvoiceNumber"
        Me.chkUseSupplierInvoiceNumber.Size = New System.Drawing.Size(213, 21)
        Me.chkUseSupplierInvoiceNumber.TabIndex = 5
        Me.chkUseSupplierInvoiceNumber.Text = "Use Supplier Invoice Number"
        Me.chkUseSupplierInvoiceNumber.UseVisualStyleBackColor = True
        '
        'chkShowDetailBegining
        '
        Me.chkShowDetailBegining.AutoSize = True
        Me.chkShowDetailBegining.Location = New System.Drawing.Point(256, 12)
        Me.chkShowDetailBegining.Name = "chkShowDetailBegining"
        Me.chkShowDetailBegining.Size = New System.Drawing.Size(279, 21)
        Me.chkShowDetailBegining.TabIndex = 6
        Me.chkShowDetailBegining.Text = "Show Begining && Ending Balance Detail"
        Me.chkShowDetailBegining.UseVisualStyleBackColor = True
        '
        'FrmTransaksiMonthlyClosing_DetailCalc2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.chkShowDetailBegining)
        Me.Controls.Add(Me.chkUseSupplierInvoiceNumber)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmTransaksiMonthlyClosing_DetailCalc2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Cost Calculation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents chkUseSupplierInvoiceNumber As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowDetailBegining As System.Windows.Forms.CheckBox
End Class
