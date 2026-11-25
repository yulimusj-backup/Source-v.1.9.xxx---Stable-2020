<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_StockBundling_Quick1
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
        Me.chkShowCodeNo = New System.Windows.Forms.CheckBox
        Me.chkShowDesc = New System.Windows.Forms.CheckBox
        Me.chkShowProcessDesc = New System.Windows.Forms.CheckBox
        Me.chkShowProcessCode = New System.Windows.Forms.CheckBox
        Me.chkShowEndTime = New System.Windows.Forms.CheckBox
        Me.chkShowStart = New System.Windows.Forms.CheckBox
        Me.TableLayoutPanel1.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(182, 60)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 6
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
        'chkShowCodeNo
        '
        Me.chkShowCodeNo.AutoSize = True
        Me.chkShowCodeNo.Location = New System.Drawing.Point(9, 10)
        Me.chkShowCodeNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowCodeNo.Name = "chkShowCodeNo"
        Me.chkShowCodeNo.Size = New System.Drawing.Size(98, 17)
        Me.chkShowCodeNo.TabIndex = 0
        Me.chkShowCodeNo.Text = "Show Code No"
        Me.chkShowCodeNo.UseVisualStyleBackColor = True
        '
        'chkShowDesc
        '
        Me.chkShowDesc.AutoSize = True
        Me.chkShowDesc.Location = New System.Drawing.Point(9, 32)
        Me.chkShowDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowDesc.Name = "chkShowDesc"
        Me.chkShowDesc.Size = New System.Drawing.Size(109, 17)
        Me.chkShowDesc.TabIndex = 1
        Me.chkShowDesc.Text = "Show Description"
        Me.chkShowDesc.UseVisualStyleBackColor = True
        '
        'chkShowProcessDesc
        '
        Me.chkShowProcessDesc.AutoSize = True
        Me.chkShowProcessDesc.Location = New System.Drawing.Point(9, 76)
        Me.chkShowProcessDesc.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowProcessDesc.Name = "chkShowProcessDesc"
        Me.chkShowProcessDesc.Size = New System.Drawing.Size(150, 17)
        Me.chkShowProcessDesc.TabIndex = 3
        Me.chkShowProcessDesc.Text = "Show Process Description"
        Me.chkShowProcessDesc.UseVisualStyleBackColor = True
        '
        'chkShowProcessCode
        '
        Me.chkShowProcessCode.AutoSize = True
        Me.chkShowProcessCode.Location = New System.Drawing.Point(9, 54)
        Me.chkShowProcessCode.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowProcessCode.Name = "chkShowProcessCode"
        Me.chkShowProcessCode.Size = New System.Drawing.Size(122, 17)
        Me.chkShowProcessCode.TabIndex = 2
        Me.chkShowProcessCode.Text = "Show Process Code"
        Me.chkShowProcessCode.UseVisualStyleBackColor = True
        '
        'chkShowEndTime
        '
        Me.chkShowEndTime.AutoSize = True
        Me.chkShowEndTime.Location = New System.Drawing.Point(182, 32)
        Me.chkShowEndTime.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowEndTime.Name = "chkShowEndTime"
        Me.chkShowEndTime.Size = New System.Drawing.Size(101, 17)
        Me.chkShowEndTime.TabIndex = 5
        Me.chkShowEndTime.Text = "Show End Time"
        Me.chkShowEndTime.UseVisualStyleBackColor = True
        '
        'chkShowStart
        '
        Me.chkShowStart.AutoSize = True
        Me.chkShowStart.Location = New System.Drawing.Point(182, 10)
        Me.chkShowStart.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowStart.Name = "chkShowStart"
        Me.chkShowStart.Size = New System.Drawing.Size(104, 17)
        Me.chkShowStart.TabIndex = 4
        Me.chkShowStart.Text = "Show Start Time"
        Me.chkShowStart.UseVisualStyleBackColor = True
        '
        'FrmTrans_StockBundling_Quick1
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(338, 100)
        Me.Controls.Add(Me.chkShowEndTime)
        Me.Controls.Add(Me.chkShowStart)
        Me.Controls.Add(Me.chkShowProcessDesc)
        Me.Controls.Add(Me.chkShowProcessCode)
        Me.Controls.Add(Me.chkShowDesc)
        Me.Controls.Add(Me.chkShowCodeNo)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(344, 128)
        Me.Name = "FrmTrans_StockBundling_Quick1"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Columns Setting"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents chkShowCodeNo As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowDesc As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowProcessDesc As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowProcessCode As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowEndTime As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowStart As System.Windows.Forms.CheckBox

End Class
