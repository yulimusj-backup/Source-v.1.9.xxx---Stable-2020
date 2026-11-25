<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_ColumnsSettings
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
        Me.chkShowOT = New System.Windows.Forms.CheckBox
        Me.ChkShowHolidayOT = New System.Windows.Forms.CheckBox
        Me.ChkShowEarlyOT = New System.Windows.Forms.CheckBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkShowHolidayDelay = New System.Windows.Forms.CheckBox
        Me.ChkShowHolidayBreak = New System.Windows.Forms.CheckBox
        Me.ChkShowDelay = New System.Windows.Forms.CheckBox
        Me.ChkShowEarlyBreak = New System.Windows.Forms.CheckBox
        Me.ChkShowBreak = New System.Windows.Forms.CheckBox
        Me.ChkShowEarlyDelay = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkShowOT
        '
        Me.chkShowOT.AutoSize = True
        Me.chkShowOT.Location = New System.Drawing.Point(5, 35)
        Me.chkShowOT.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowOT.Name = "chkShowOT"
        Me.chkShowOT.Size = New System.Drawing.Size(71, 17)
        Me.chkShowOT.TabIndex = 3
        Me.chkShowOT.Text = "Show OT"
        Me.chkShowOT.UseVisualStyleBackColor = True
        '
        'ChkShowHolidayOT
        '
        Me.ChkShowHolidayOT.AutoSize = True
        Me.ChkShowHolidayOT.Location = New System.Drawing.Point(5, 56)
        Me.ChkShowHolidayOT.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowHolidayOT.Name = "ChkShowHolidayOT"
        Me.ChkShowHolidayOT.Size = New System.Drawing.Size(113, 17)
        Me.ChkShowHolidayOT.TabIndex = 6
        Me.ChkShowHolidayOT.Text = "Show Holoday OT"
        Me.ChkShowHolidayOT.UseVisualStyleBackColor = True
        '
        'ChkShowEarlyOT
        '
        Me.ChkShowEarlyOT.AutoSize = True
        Me.ChkShowEarlyOT.Location = New System.Drawing.Point(5, 14)
        Me.ChkShowEarlyOT.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowEarlyOT.Name = "ChkShowEarlyOT"
        Me.ChkShowEarlyOT.Size = New System.Drawing.Size(97, 17)
        Me.ChkShowEarlyOT.TabIndex = 0
        Me.ChkShowEarlyOT.Text = "Show Early OT"
        Me.ChkShowEarlyOT.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(330, 89)
        Me.btnOK.MaximumSize = New System.Drawing.Size(53, 21)
        Me.btnOK.MinimumSize = New System.Drawing.Size(53, 21)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(53, 21)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkShowHolidayDelay)
        Me.GroupBox1.Controls.Add(Me.chkShowOT)
        Me.GroupBox1.Controls.Add(Me.ChkShowHolidayBreak)
        Me.GroupBox1.Controls.Add(Me.ChkShowDelay)
        Me.GroupBox1.Controls.Add(Me.ChkShowHolidayOT)
        Me.GroupBox1.Controls.Add(Me.ChkShowEarlyOT)
        Me.GroupBox1.Controls.Add(Me.ChkShowEarlyBreak)
        Me.GroupBox1.Controls.Add(Me.ChkShowBreak)
        Me.GroupBox1.Controls.Add(Me.ChkShowEarlyDelay)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(377, 79)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ChkShowHolidayDelay
        '
        Me.ChkShowHolidayDelay.AutoSize = True
        Me.ChkShowHolidayDelay.Location = New System.Drawing.Point(122, 56)
        Me.ChkShowHolidayDelay.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowHolidayDelay.Name = "ChkShowHolidayDelay"
        Me.ChkShowHolidayDelay.Size = New System.Drawing.Size(121, 17)
        Me.ChkShowHolidayDelay.TabIndex = 7
        Me.ChkShowHolidayDelay.Text = "Show Holiday Delay"
        Me.ChkShowHolidayDelay.UseVisualStyleBackColor = True
        '
        'ChkShowHolidayBreak
        '
        Me.ChkShowHolidayBreak.AutoSize = True
        Me.ChkShowHolidayBreak.Location = New System.Drawing.Point(247, 56)
        Me.ChkShowHolidayBreak.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowHolidayBreak.Name = "ChkShowHolidayBreak"
        Me.ChkShowHolidayBreak.Size = New System.Drawing.Size(122, 17)
        Me.ChkShowHolidayBreak.TabIndex = 8
        Me.ChkShowHolidayBreak.Text = "Show Holiday Break"
        Me.ChkShowHolidayBreak.UseVisualStyleBackColor = True
        '
        'ChkShowDelay
        '
        Me.ChkShowDelay.AutoSize = True
        Me.ChkShowDelay.Location = New System.Drawing.Point(122, 35)
        Me.ChkShowDelay.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowDelay.Name = "ChkShowDelay"
        Me.ChkShowDelay.Size = New System.Drawing.Size(83, 17)
        Me.ChkShowDelay.TabIndex = 4
        Me.ChkShowDelay.Text = "Show Delay"
        Me.ChkShowDelay.UseVisualStyleBackColor = True
        '
        'ChkShowEarlyBreak
        '
        Me.ChkShowEarlyBreak.AutoSize = True
        Me.ChkShowEarlyBreak.Location = New System.Drawing.Point(247, 14)
        Me.ChkShowEarlyBreak.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowEarlyBreak.Name = "ChkShowEarlyBreak"
        Me.ChkShowEarlyBreak.Size = New System.Drawing.Size(110, 17)
        Me.ChkShowEarlyBreak.TabIndex = 2
        Me.ChkShowEarlyBreak.Text = "Show Early Break"
        Me.ChkShowEarlyBreak.UseVisualStyleBackColor = True
        '
        'ChkShowBreak
        '
        Me.ChkShowBreak.AutoSize = True
        Me.ChkShowBreak.Location = New System.Drawing.Point(247, 35)
        Me.ChkShowBreak.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowBreak.Name = "ChkShowBreak"
        Me.ChkShowBreak.Size = New System.Drawing.Size(84, 17)
        Me.ChkShowBreak.TabIndex = 5
        Me.ChkShowBreak.Text = "Show Break"
        Me.ChkShowBreak.UseVisualStyleBackColor = True
        '
        'ChkShowEarlyDelay
        '
        Me.ChkShowEarlyDelay.AutoSize = True
        Me.ChkShowEarlyDelay.Location = New System.Drawing.Point(122, 14)
        Me.ChkShowEarlyDelay.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowEarlyDelay.Name = "ChkShowEarlyDelay"
        Me.ChkShowEarlyDelay.Size = New System.Drawing.Size(109, 17)
        Me.ChkShowEarlyDelay.TabIndex = 1
        Me.ChkShowEarlyDelay.Text = "Show Early Delay"
        Me.ChkShowEarlyDelay.UseVisualStyleBackColor = True
        '
        'FrmTrans_SPL_ColumnsSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 122)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximumSize = New System.Drawing.Size(401, 146)
        Me.MinimumSize = New System.Drawing.Size(401, 146)
        Me.Name = "FrmTrans_SPL_ColumnsSettings"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Columns Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ChkShowHolidayOT As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowEarlyOT As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowOT As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkShowEarlyBreak As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowEarlyDelay As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowDelay As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowBreak As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowHolidayDelay As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowHolidayBreak As System.Windows.Forms.CheckBox
End Class
