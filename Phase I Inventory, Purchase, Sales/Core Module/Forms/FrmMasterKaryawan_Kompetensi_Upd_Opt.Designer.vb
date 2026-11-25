<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKaryawan_Kompetensi_Upd_Opt
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
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.DtpNextCertDate = New System.Windows.Forms.DateTimePicker
        Me.cboDefValue = New System.Windows.Forms.ComboBox
        Me.DtpDefValue = New System.Windows.Forms.DateTimePicker
        Me.TxtDefValueText = New System.Windows.Forms.TextBox
        Me.cboAlfabetic = New System.Windows.Forms.ComboBox
        Me.TxtDefValueDigit = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(321, 124)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(42, 20)
        Me.BtnOK.TabIndex = 3
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(369, 124)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(42, 20)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 15)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(34, 13)
        Me.Label6.TabIndex = 56
        Me.Label6.Text = "Value"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 83)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 54
        Me.Label1.Text = "Remarks"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 37)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 13)
        Me.Label5.TabIndex = 53
        Me.Label5.Text = "Next Certificate Date"
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(163, 83)
        Me.txtRemarks.MaxLength = 250
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(249, 36)
        Me.txtRemarks.TabIndex = 2
        '
        'DtpNextCertDate
        '
        Me.DtpNextCertDate.CustomFormat = "dd MMM yyyy"
        Me.DtpNextCertDate.Location = New System.Drawing.Point(163, 37)
        Me.DtpNextCertDate.Name = "DtpNextCertDate"
        Me.DtpNextCertDate.Size = New System.Drawing.Size(249, 20)
        Me.DtpNextCertDate.TabIndex = 1
        '
        'cboDefValue
        '
        Me.cboDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDefValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDefValue.FormattingEnabled = True
        Me.cboDefValue.Items.AddRange(New Object() {"No", "Yes"})
        Me.cboDefValue.Location = New System.Drawing.Point(163, 12)
        Me.cboDefValue.Name = "cboDefValue"
        Me.cboDefValue.Size = New System.Drawing.Size(249, 21)
        Me.cboDefValue.TabIndex = 0
        '
        'DtpDefValue
        '
        Me.DtpDefValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDefValue.Location = New System.Drawing.Point(163, 12)
        Me.DtpDefValue.Name = "DtpDefValue"
        Me.DtpDefValue.Size = New System.Drawing.Size(249, 20)
        Me.DtpDefValue.TabIndex = 60
        '
        'TxtDefValueText
        '
        Me.TxtDefValueText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefValueText.Location = New System.Drawing.Point(163, 12)
        Me.TxtDefValueText.MaxLength = 250
        Me.TxtDefValueText.Name = "TxtDefValueText"
        Me.TxtDefValueText.Size = New System.Drawing.Size(249, 20)
        Me.TxtDefValueText.TabIndex = 61
        '
        'cboAlfabetic
        '
        Me.cboAlfabetic.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAlfabetic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAlfabetic.FormattingEnabled = True
        Me.cboAlfabetic.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F"})
        Me.cboAlfabetic.Location = New System.Drawing.Point(163, 12)
        Me.cboAlfabetic.Name = "cboAlfabetic"
        Me.cboAlfabetic.Size = New System.Drawing.Size(249, 21)
        Me.cboAlfabetic.TabIndex = 59
        '
        'TxtDefValueDigit
        '
        Me.TxtDefValueDigit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDefValueDigit.Location = New System.Drawing.Point(163, 12)
        Me.TxtDefValueDigit.MaxLength = 15
        Me.TxtDefValueDigit.Name = "TxtDefValueDigit"
        Me.TxtDefValueDigit.Size = New System.Drawing.Size(249, 20)
        Me.TxtDefValueDigit.TabIndex = 62
        Me.TxtDefValueDigit.Text = "0"
        Me.TxtDefValueDigit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(2, 138)
        Me.TextBox1.MaxLength = 250
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(14, 36)
        Me.TextBox1.TabIndex = 63
        Me.TextBox1.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(163, 60)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(162, 17)
        Me.CheckBox1.TabIndex = 64
        Me.CheckBox1.Text = "Include Next Certificate Date"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FrmMasterKaryawan_Kompetensi_Upd_Opt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(423, 156)
        Me.ControlBox = False
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cboDefValue)
        Me.Controls.Add(Me.DtpDefValue)
        Me.Controls.Add(Me.TxtDefValueText)
        Me.Controls.Add(Me.cboAlfabetic)
        Me.Controls.Add(Me.TxtDefValueDigit)
        Me.Controls.Add(Me.DtpNextCertDate)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtRemarks)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnOK)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(439, 194)
        Me.Name = "FrmMasterKaryawan_Kompetensi_Upd_Opt"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Competence Value"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents DtpNextCertDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboDefValue As System.Windows.Forms.ComboBox
    Friend WithEvents DtpDefValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtDefValueText As System.Windows.Forms.TextBox
    Friend WithEvents cboAlfabetic As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDefValueDigit As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
