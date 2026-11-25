<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiClosing
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
        Me.dtpConvertion = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpClosingDate = New System.Windows.Forms.DateTimePicker
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpConventionDate_HR = New System.Windows.Forms.DateTimePicker
        Me.btnExplain_ConvertionDate = New System.Windows.Forms.Button
        Me.btnExplain_PayrollConvertionDate = New System.Windows.Forms.Button
        Me.btnExplainClosingDate = New System.Windows.Forms.Button
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpConvertion
        '
        Me.dtpConvertion.CustomFormat = "dd MMM yyyy"
        Me.dtpConvertion.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpConvertion.Location = New System.Drawing.Point(185, 15)
        Me.dtpConvertion.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpConvertion.Name = "dtpConvertion"
        Me.dtpConvertion.Size = New System.Drawing.Size(358, 22)
        Me.dtpConvertion.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 18)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(110, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Convertion Date"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 78)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Closing Date"
        '
        'dtpClosingDate
        '
        Me.dtpClosingDate.CustomFormat = "dd MMM yyyy"
        Me.dtpClosingDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpClosingDate.Location = New System.Drawing.Point(185, 75)
        Me.dtpClosingDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpClosingDate.Name = "dtpClosingDate"
        Me.dtpClosingDate.Size = New System.Drawing.Size(358, 22)
        Me.dtpClosingDate.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(605, 57)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(101, 34)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(605, 15)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(101, 34)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 129)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(719, 25)
        Me.StatusStrip1.TabIndex = 40
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(349, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(349, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 48)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(160, 17)
        Me.Label3.TabIndex = 42
        Me.Label3.Text = "Payroll Convention Date"
        '
        'DtpConventionDate_HR
        '
        Me.DtpConventionDate_HR.CustomFormat = "dd MMM yyyy"
        Me.DtpConventionDate_HR.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpConventionDate_HR.Location = New System.Drawing.Point(185, 45)
        Me.DtpConventionDate_HR.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpConventionDate_HR.Name = "DtpConventionDate_HR"
        Me.DtpConventionDate_HR.Size = New System.Drawing.Size(358, 22)
        Me.DtpConventionDate_HR.TabIndex = 1
        '
        'btnExplain_ConvertionDate
        '
        Me.btnExplain_ConvertionDate.Location = New System.Drawing.Point(551, 15)
        Me.btnExplain_ConvertionDate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExplain_ConvertionDate.Name = "btnExplain_ConvertionDate"
        Me.btnExplain_ConvertionDate.Size = New System.Drawing.Size(46, 22)
        Me.btnExplain_ConvertionDate.TabIndex = 43
        Me.btnExplain_ConvertionDate.Text = "?"
        Me.btnExplain_ConvertionDate.UseVisualStyleBackColor = True
        '
        'btnExplain_PayrollConvertionDate
        '
        Me.btnExplain_PayrollConvertionDate.Location = New System.Drawing.Point(551, 45)
        Me.btnExplain_PayrollConvertionDate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExplain_PayrollConvertionDate.Name = "btnExplain_PayrollConvertionDate"
        Me.btnExplain_PayrollConvertionDate.Size = New System.Drawing.Size(46, 22)
        Me.btnExplain_PayrollConvertionDate.TabIndex = 44
        Me.btnExplain_PayrollConvertionDate.Text = "?"
        Me.btnExplain_PayrollConvertionDate.UseVisualStyleBackColor = True
        '
        'btnExplainClosingDate
        '
        Me.btnExplainClosingDate.Location = New System.Drawing.Point(551, 75)
        Me.btnExplainClosingDate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExplainClosingDate.Name = "btnExplainClosingDate"
        Me.btnExplainClosingDate.Size = New System.Drawing.Size(46, 22)
        Me.btnExplainClosingDate.TabIndex = 45
        Me.btnExplainClosingDate.Text = "?"
        Me.btnExplainClosingDate.UseVisualStyleBackColor = True
        '
        'FrmTransaksiClosing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(719, 154)
        Me.Controls.Add(Me.btnExplainClosingDate)
        Me.Controls.Add(Me.btnExplain_PayrollConvertionDate)
        Me.Controls.Add(Me.btnExplain_ConvertionDate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DtpConventionDate_HR)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpClosingDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpConvertion)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(727, 191)
        Me.Name = "FrmTransaksiClosing"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aplication Closing Setup"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpConvertion As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpClosingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpConventionDate_HR As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnExplain_ConvertionDate As System.Windows.Forms.Button
    Friend WithEvents btnExplain_PayrollConvertionDate As System.Windows.Forms.Button
    Friend WithEvents btnExplainClosingDate As System.Windows.Forms.Button
End Class
