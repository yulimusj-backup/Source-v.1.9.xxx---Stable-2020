<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiLoan_Information
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiLoan_Information))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TxtAnnuity = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TxtFlat = New System.Windows.Forms.TextBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.TxtEffective = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(0, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(443, 380)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TxtAnnuity)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(435, 354)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Annuity"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TxtAnnuity
        '
        Me.TxtAnnuity.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAnnuity.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TxtAnnuity.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtAnnuity.Location = New System.Drawing.Point(8, 6)
        Me.TxtAnnuity.Multiline = True
        Me.TxtAnnuity.Name = "TxtAnnuity"
        Me.TxtAnnuity.ReadOnly = True
        Me.TxtAnnuity.Size = New System.Drawing.Size(420, 342)
        Me.TxtAnnuity.TabIndex = 0
        Me.TxtAnnuity.Text = resources.GetString("TxtAnnuity.Text")
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TxtFlat)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(435, 354)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Flat"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TxtFlat
        '
        Me.TxtFlat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFlat.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TxtFlat.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtFlat.Location = New System.Drawing.Point(6, 6)
        Me.TxtFlat.Multiline = True
        Me.TxtFlat.Name = "TxtFlat"
        Me.TxtFlat.ReadOnly = True
        Me.TxtFlat.Size = New System.Drawing.Size(423, 342)
        Me.TxtFlat.TabIndex = 2
        Me.TxtFlat.Text = resources.GetString("TxtFlat.Text")
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TxtEffective)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(435, 354)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Effective"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TxtEffective
        '
        Me.TxtEffective.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEffective.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TxtEffective.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtEffective.Location = New System.Drawing.Point(6, 6)
        Me.TxtEffective.Multiline = True
        Me.TxtEffective.Name = "TxtEffective"
        Me.TxtEffective.ReadOnly = True
        Me.TxtEffective.Size = New System.Drawing.Size(423, 342)
        Me.TxtEffective.TabIndex = 2
        Me.TxtEffective.Text = resources.GetString("TxtEffective.Text")
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(681, 131)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiLoan_Information
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(444, 382)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(460, 420)
        Me.Name = "FrmTransaksiLoan_Information"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Loan Information Data"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TxtAnnuity As System.Windows.Forms.TextBox
    Friend WithEvents TxtFlat As System.Windows.Forms.TextBox
    Friend WithEvents TxtEffective As System.Windows.Forms.TextBox
End Class
