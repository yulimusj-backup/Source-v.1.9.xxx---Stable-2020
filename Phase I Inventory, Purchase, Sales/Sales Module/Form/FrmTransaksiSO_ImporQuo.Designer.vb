<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_ImporQuo
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.rboHeaderOnly = New System.Windows.Forms.RadioButton
        Me.rboItemOnly = New System.Windows.Forms.RadioButton
        Me.rboAll = New System.Windows.Forms.RadioButton
        Me.gboItemMode = New System.Windows.Forms.GroupBox
        Me.rboReplaceData = New System.Windows.Forms.RadioButton
        Me.rboAppend = New System.Windows.Forms.RadioButton
        Me.TableLayoutPanel1.SuspendLayout()
        Me.gboItemMode.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(165, 104)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(205, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "How do you want to copy quotation data?"
        '
        'rboHeaderOnly
        '
        Me.rboHeaderOnly.AutoSize = True
        Me.rboHeaderOnly.Checked = True
        Me.rboHeaderOnly.Location = New System.Drawing.Point(15, 25)
        Me.rboHeaderOnly.Name = "rboHeaderOnly"
        Me.rboHeaderOnly.Size = New System.Drawing.Size(111, 17)
        Me.rboHeaderOnly.TabIndex = 2
        Me.rboHeaderOnly.TabStop = True
        Me.rboHeaderOnly.Text = "Copy Header Only"
        Me.rboHeaderOnly.UseVisualStyleBackColor = True
        '
        'rboItemOnly
        '
        Me.rboItemOnly.AutoSize = True
        Me.rboItemOnly.Location = New System.Drawing.Point(15, 48)
        Me.rboItemOnly.Name = "rboItemOnly"
        Me.rboItemOnly.Size = New System.Drawing.Size(96, 17)
        Me.rboItemOnly.TabIndex = 3
        Me.rboItemOnly.Text = "Copy Item Only"
        Me.rboItemOnly.UseVisualStyleBackColor = True
        '
        'rboAll
        '
        Me.rboAll.AutoSize = True
        Me.rboAll.Location = New System.Drawing.Point(15, 71)
        Me.rboAll.Name = "rboAll"
        Me.rboAll.Size = New System.Drawing.Size(119, 17)
        Me.rboAll.TabIndex = 4
        Me.rboAll.Text = "Copy Item && Header"
        Me.rboAll.UseVisualStyleBackColor = True
        '
        'gboItemMode
        '
        Me.gboItemMode.Controls.Add(Me.rboReplaceData)
        Me.gboItemMode.Controls.Add(Me.rboAppend)
        Me.gboItemMode.Enabled = False
        Me.gboItemMode.Location = New System.Drawing.Point(140, 25)
        Me.gboItemMode.Name = "gboItemMode"
        Me.gboItemMode.Size = New System.Drawing.Size(171, 63)
        Me.gboItemMode.TabIndex = 5
        Me.gboItemMode.TabStop = False
        Me.gboItemMode.Text = "Item Copy Mode"
        '
        'rboReplaceData
        '
        Me.rboReplaceData.AutoSize = True
        Me.rboReplaceData.Location = New System.Drawing.Point(6, 42)
        Me.rboReplaceData.Name = "rboReplaceData"
        Me.rboReplaceData.Size = New System.Drawing.Size(91, 17)
        Me.rboReplaceData.TabIndex = 1
        Me.rboReplaceData.Text = "Replace Data"
        Me.rboReplaceData.UseVisualStyleBackColor = True
        '
        'rboAppend
        '
        Me.rboAppend.AutoSize = True
        Me.rboAppend.Checked = True
        Me.rboAppend.Location = New System.Drawing.Point(6, 19)
        Me.rboAppend.Name = "rboAppend"
        Me.rboAppend.Size = New System.Drawing.Size(88, 17)
        Me.rboAppend.TabIndex = 0
        Me.rboAppend.TabStop = True
        Me.rboAppend.Text = "Append Data"
        Me.rboAppend.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSO_ImporQuo
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(323, 145)
        Me.Controls.Add(Me.gboItemMode)
        Me.Controls.Add(Me.rboAll)
        Me.Controls.Add(Me.rboItemOnly)
        Me.Controls.Add(Me.rboHeaderOnly)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiSO_ImporQuo"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Data transfer"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.gboItemMode.ResumeLayout(False)
        Me.gboItemMode.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rboHeaderOnly As System.Windows.Forms.RadioButton
    Friend WithEvents rboItemOnly As System.Windows.Forms.RadioButton
    Friend WithEvents rboAll As System.Windows.Forms.RadioButton
    Friend WithEvents gboItemMode As System.Windows.Forms.GroupBox
    Friend WithEvents rboReplaceData As System.Windows.Forms.RadioButton
    Friend WithEvents rboAppend As System.Windows.Forms.RadioButton

End Class
