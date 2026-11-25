<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_EditPPh
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
        Me.TxtPPhRutin = New System.Windows.Forms.TextBox
        Me.TxtPPHNonRutin = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtRef = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "PPh Rutin"
        '
        'TxtPPhRutin
        '
        Me.TxtPPhRutin.Location = New System.Drawing.Point(119, 12)
        Me.TxtPPhRutin.Name = "TxtPPhRutin"
        Me.TxtPPhRutin.Size = New System.Drawing.Size(242, 22)
        Me.TxtPPhRutin.TabIndex = 1
        Me.TxtPPhRutin.Text = "0"
        Me.TxtPPhRutin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPPHNonRutin
        '
        Me.TxtPPHNonRutin.Location = New System.Drawing.Point(119, 40)
        Me.TxtPPHNonRutin.Name = "TxtPPHNonRutin"
        Me.TxtPPHNonRutin.Size = New System.Drawing.Size(242, 22)
        Me.TxtPPHNonRutin.TabIndex = 3
        Me.TxtPPHNonRutin.Text = "0"
        Me.TxtPPHNonRutin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "PPh Non Rutin"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(367, 11)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtRef
        '
        Me.TxtRef.Location = New System.Drawing.Point(367, 40)
        Me.TxtRef.Name = "TxtRef"
        Me.TxtRef.Size = New System.Drawing.Size(51, 22)
        Me.TxtRef.TabIndex = 5
        Me.TxtRef.TabStop = False
        Me.TxtRef.Text = "0"
        Me.TxtRef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtRef.Visible = False
        '
        'FrmTrans_Gaji_EditPPh
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(454, 75)
        Me.Controls.Add(Me.TxtRef)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtPPHNonRutin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtPPhRutin)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(462, 112)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(462, 112)
        Me.Name = "FrmTrans_Gaji_EditPPh"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit PPh"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPPhRutin As System.Windows.Forms.TextBox
    Friend WithEvents TxtPPHNonRutin As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtRef As System.Windows.Forms.TextBox
End Class
