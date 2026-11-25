<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_Reorder_LMB
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
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnStep1 = New System.Windows.Forms.Button
        Me.PrgBarStep1 = New System.Windows.Forms.ProgressBar
        Me.lblProgressStep1 = New System.Windows.Forms.Label
        Me.LblProgressStep3 = New System.Windows.Forms.Label
        Me.PrgBarStep3 = New System.Windows.Forms.ProgressBar
        Me.btnStep3 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.LblProgressStep2 = New System.Windows.Forms.Label
        Me.PrgBarStep2 = New System.Windows.Forms.ProgressBar
        Me.btnStep2 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.dtpStep3 = New System.Windows.Forms.DateTimePicker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.LblElapseTime = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(209, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Step 1: Buat workshift allocation"
        '
        'btnStep1
        '
        Me.btnStep1.Location = New System.Drawing.Point(623, 12)
        Me.btnStep1.Name = "btnStep1"
        Me.btnStep1.Size = New System.Drawing.Size(75, 23)
        Me.btnStep1.TabIndex = 1
        Me.btnStep1.Text = "Start"
        Me.btnStep1.UseVisualStyleBackColor = True
        '
        'PrgBarStep1
        '
        Me.PrgBarStep1.Location = New System.Drawing.Point(246, 12)
        Me.PrgBarStep1.Name = "PrgBarStep1"
        Me.PrgBarStep1.Size = New System.Drawing.Size(371, 23)
        Me.PrgBarStep1.TabIndex = 2
        '
        'lblProgressStep1
        '
        Me.lblProgressStep1.AutoSize = True
        Me.lblProgressStep1.Location = New System.Drawing.Point(243, 38)
        Me.lblProgressStep1.Name = "lblProgressStep1"
        Me.lblProgressStep1.Size = New System.Drawing.Size(61, 17)
        Me.lblProgressStep1.TabIndex = 3
        Me.lblProgressStep1.Text = "Ready..."
        '
        'LblProgressStep3
        '
        Me.LblProgressStep3.AutoSize = True
        Me.LblProgressStep3.Location = New System.Drawing.Point(243, 130)
        Me.LblProgressStep3.Name = "LblProgressStep3"
        Me.LblProgressStep3.Size = New System.Drawing.Size(61, 17)
        Me.LblProgressStep3.TabIndex = 7
        Me.LblProgressStep3.Text = "Ready..."
        '
        'PrgBarStep3
        '
        Me.PrgBarStep3.Location = New System.Drawing.Point(246, 104)
        Me.PrgBarStep3.Name = "PrgBarStep3"
        Me.PrgBarStep3.Size = New System.Drawing.Size(213, 23)
        Me.PrgBarStep3.TabIndex = 6
        '
        'btnStep3
        '
        Me.btnStep3.Location = New System.Drawing.Point(623, 104)
        Me.btnStep3.Name = "btnStep3"
        Me.btnStep3.Size = New System.Drawing.Size(75, 23)
        Me.btnStep3.TabIndex = 5
        Me.btnStep3.Text = "Start"
        Me.btnStep3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(178, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Step 3: Urutkan tanpa jeda"
        '
        'LblProgressStep2
        '
        Me.LblProgressStep2.AutoSize = True
        Me.LblProgressStep2.Location = New System.Drawing.Point(243, 84)
        Me.LblProgressStep2.Name = "LblProgressStep2"
        Me.LblProgressStep2.Size = New System.Drawing.Size(61, 17)
        Me.LblProgressStep2.TabIndex = 11
        Me.LblProgressStep2.Text = "Ready..."
        '
        'PrgBarStep2
        '
        Me.PrgBarStep2.Location = New System.Drawing.Point(246, 58)
        Me.PrgBarStep2.Name = "PrgBarStep2"
        Me.PrgBarStep2.Size = New System.Drawing.Size(371, 23)
        Me.PrgBarStep2.TabIndex = 10
        '
        'btnStep2
        '
        Me.btnStep2.Location = New System.Drawing.Point(623, 58)
        Me.btnStep2.Name = "btnStep2"
        Me.btnStep2.Size = New System.Drawing.Size(75, 23)
        Me.btnStep2.TabIndex = 9
        Me.btnStep2.Text = "Start"
        Me.btnStep2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 61)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(134, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Step 2: Cek anomali"
        '
        'BackgroundWorker1
        '
        '
        'dtpStep3
        '
        Me.dtpStep3.CustomFormat = "dd MMM yyyy"
        Me.dtpStep3.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStep3.Location = New System.Drawing.Point(465, 104)
        Me.dtpStep3.Name = "dtpStep3"
        Me.dtpStep3.Size = New System.Drawing.Size(152, 22)
        Me.dtpStep3.TabIndex = 16
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker3
        '
        '
        'LblElapseTime
        '
        Me.LblElapseTime.AutoSize = True
        Me.LblElapseTime.Location = New System.Drawing.Point(12, 179)
        Me.LblElapseTime.Name = "LblElapseTime"
        Me.LblElapseTime.Size = New System.Drawing.Size(61, 17)
        Me.LblElapseTime.TabIndex = 17
        Me.LblElapseTime.Text = "Ready..."
        '
        'Timer1
        '
        Me.Timer1.Interval = 5000
        '
        'FrmTrans_ProductionSchedule_Reorder_LMB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(710, 205)
        Me.Controls.Add(Me.LblElapseTime)
        Me.Controls.Add(Me.dtpStep3)
        Me.Controls.Add(Me.LblProgressStep2)
        Me.Controls.Add(Me.PrgBarStep2)
        Me.Controls.Add(Me.btnStep2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LblProgressStep3)
        Me.Controls.Add(Me.PrgBarStep3)
        Me.Controls.Add(Me.btnStep3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblProgressStep1)
        Me.Controls.Add(Me.PrgBarStep1)
        Me.Controls.Add(Me.btnStep1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FrmTrans_ProductionSchedule_Reorder_LMB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Re-OrderSchedule LMB"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnStep1 As System.Windows.Forms.Button
    Friend WithEvents PrgBarStep1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgressStep1 As System.Windows.Forms.Label
    Friend WithEvents LblProgressStep3 As System.Windows.Forms.Label
    Friend WithEvents PrgBarStep3 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnStep3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblProgressStep2 As System.Windows.Forms.Label
    Friend WithEvents PrgBarStep2 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnStep2 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dtpStep3 As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblElapseTime As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
