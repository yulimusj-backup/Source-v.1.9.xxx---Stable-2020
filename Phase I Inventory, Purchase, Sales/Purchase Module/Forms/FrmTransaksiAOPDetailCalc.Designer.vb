<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiAOPDetailCalc
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
        Me.TxtDetailCalculation = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtDetailCalculation
        '
        Me.TxtDetailCalculation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDetailCalculation.BackColor = System.Drawing.Color.White
        Me.TxtDetailCalculation.Location = New System.Drawing.Point(12, 12)
        Me.TxtDetailCalculation.Multiline = True
        Me.TxtDetailCalculation.Name = "TxtDetailCalculation"
        Me.TxtDetailCalculation.ReadOnly = True
        Me.TxtDetailCalculation.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TxtDetailCalculation.Size = New System.Drawing.Size(583, 261)
        Me.TxtDetailCalculation.TabIndex = 1
        Me.TxtDetailCalculation.WordWrap = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.Location = New System.Drawing.Point(476, 279)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(119, 28)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiAOPDetailCalc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnOK
        Me.ClientSize = New System.Drawing.Size(607, 319)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtDetailCalculation)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiAOPDetailCalc"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Unit Price Calculation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtDetailCalculation As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
