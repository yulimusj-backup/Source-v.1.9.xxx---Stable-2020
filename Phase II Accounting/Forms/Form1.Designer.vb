<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.btnFindJob = New System.Windows.Forms.Button
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.LblJob = New System.Windows.Forms.Label
        Me.TxtPrimaryDept = New System.Windows.Forms.TextBox
        Me.btnFindDept = New System.Windows.Forms.Button
        Me.TxtDeptCode = New System.Windows.Forms.TextBox
        Me.LblDept = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(50, 34)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryJob.TabIndex = 59
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'btnFindJob
        '
        Me.btnFindJob.Location = New System.Drawing.Point(202, 34)
        Me.btnFindJob.Name = "btnFindJob"
        Me.btnFindJob.Size = New System.Drawing.Size(37, 21)
        Me.btnFindJob.TabIndex = 55
        Me.btnFindJob.TabStop = False
        Me.btnFindJob.Text = "Find"
        Me.btnFindJob.UseVisualStyleBackColor = True
        '
        'TxtJobCode
        '
        Me.TxtJobCode.Location = New System.Drawing.Point(86, 35)
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(110, 20)
        Me.TxtJobCode.TabIndex = 54
        '
        'LblJob
        '
        Me.LblJob.AutoSize = True
        Me.LblJob.Location = New System.Drawing.Point(2, 38)
        Me.LblJob.Name = "LblJob"
        Me.LblJob.Size = New System.Drawing.Size(24, 13)
        Me.LblJob.TabIndex = 58
        Me.LblJob.Text = "Job"
        '
        'TxtPrimaryDept
        '
        Me.TxtPrimaryDept.Location = New System.Drawing.Point(50, 11)
        Me.TxtPrimaryDept.Name = "TxtPrimaryDept"
        Me.TxtPrimaryDept.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryDept.TabIndex = 57
        Me.TxtPrimaryDept.Text = "-1"
        Me.TxtPrimaryDept.Visible = False
        '
        'btnFindDept
        '
        Me.btnFindDept.Location = New System.Drawing.Point(202, 11)
        Me.btnFindDept.Name = "btnFindDept"
        Me.btnFindDept.Size = New System.Drawing.Size(37, 21)
        Me.btnFindDept.TabIndex = 53
        Me.btnFindDept.TabStop = False
        Me.btnFindDept.Text = "Find"
        Me.btnFindDept.UseVisualStyleBackColor = True
        '
        'TxtDeptCode
        '
        Me.TxtDeptCode.Location = New System.Drawing.Point(86, 12)
        Me.TxtDeptCode.Name = "TxtDeptCode"
        Me.TxtDeptCode.Size = New System.Drawing.Size(110, 20)
        Me.TxtDeptCode.TabIndex = 52
        '
        'LblDept
        '
        Me.LblDept.AutoSize = True
        Me.LblDept.Location = New System.Drawing.Point(2, 15)
        Me.LblDept.Name = "LblDept"
        Me.LblDept.Size = New System.Drawing.Size(68, 13)
        Me.LblDept.TabIndex = 56
        Me.LblDept.Text = "Departement"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.TxtPrimaryJob)
        Me.Controls.Add(Me.btnFindJob)
        Me.Controls.Add(Me.TxtJobCode)
        Me.Controls.Add(Me.LblJob)
        Me.Controls.Add(Me.TxtPrimaryDept)
        Me.Controls.Add(Me.btnFindDept)
        Me.Controls.Add(Me.TxtDeptCode)
        Me.Controls.Add(Me.LblDept)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents btnFindJob As System.Windows.Forms.Button
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents LblJob As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryDept As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDept As System.Windows.Forms.Button
    Friend WithEvents TxtDeptCode As System.Windows.Forms.TextBox
    Friend WithEvents LblDept As System.Windows.Forms.Label
End Class
