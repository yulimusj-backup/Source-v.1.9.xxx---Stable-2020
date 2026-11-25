<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPilihCetakLoan
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RbtJournal = New System.Windows.Forms.RadioButton
        Me.RbtDetail = New System.Windows.Forms.RadioButton
        Me.BtnShow = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnShow)
        Me.GroupBox1.Controls.Add(Me.RbtDetail)
        Me.GroupBox1.Controls.Add(Me.RbtJournal)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(222, 53)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'RbtJournal
        '
        Me.RbtJournal.AutoSize = True
        Me.RbtJournal.Location = New System.Drawing.Point(7, 20)
        Me.RbtJournal.Name = "RbtJournal"
        Me.RbtJournal.Size = New System.Drawing.Size(59, 17)
        Me.RbtJournal.TabIndex = 0
        Me.RbtJournal.TabStop = True
        Me.RbtJournal.Text = "Journal"
        Me.RbtJournal.UseVisualStyleBackColor = True
        '
        'RbtDetail
        '
        Me.RbtDetail.AutoSize = True
        Me.RbtDetail.Location = New System.Drawing.Point(72, 20)
        Me.RbtDetail.Name = "RbtDetail"
        Me.RbtDetail.Size = New System.Drawing.Size(52, 17)
        Me.RbtDetail.TabIndex = 1
        Me.RbtDetail.TabStop = True
        Me.RbtDetail.Text = "Detail"
        Me.RbtDetail.UseVisualStyleBackColor = True
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(131, 17)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(75, 23)
        Me.BtnShow.TabIndex = 2
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'FrmPilihCetakLoan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(243, 77)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximumSize = New System.Drawing.Size(259, 116)
        Me.MinimumSize = New System.Drawing.Size(259, 116)
        Me.Name = "FrmPilihCetakLoan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RbtDetail As System.Windows.Forms.RadioButton
    Friend WithEvents RbtJournal As System.Windows.Forms.RadioButton
    Friend WithEvents BtnShow As System.Windows.Forms.Button
End Class
