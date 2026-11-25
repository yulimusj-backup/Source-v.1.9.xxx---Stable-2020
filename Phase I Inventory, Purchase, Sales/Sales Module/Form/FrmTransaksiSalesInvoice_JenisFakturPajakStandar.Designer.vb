<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_JenisFakturPajakStandar
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
        Me.rboHargaJual = New System.Windows.Forms.RadioButton
        Me.rboPenggantian = New System.Windows.Forms.RadioButton
        Me.rboTermin = New System.Windows.Forms.RadioButton
        Me.rboUangMuka = New System.Windows.Forms.RadioButton
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(134, 66)
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
        'rboHargaJual
        '
        Me.rboHargaJual.AutoSize = True
        Me.rboHargaJual.Checked = True
        Me.rboHargaJual.Location = New System.Drawing.Point(12, 12)
        Me.rboHargaJual.Name = "rboHargaJual"
        Me.rboHargaJual.Size = New System.Drawing.Size(76, 17)
        Me.rboHargaJual.TabIndex = 1
        Me.rboHargaJual.TabStop = True
        Me.rboHargaJual.Text = "Harga Jual"
        Me.rboHargaJual.UseVisualStyleBackColor = True
        '
        'rboPenggantian
        '
        Me.rboPenggantian.AutoSize = True
        Me.rboPenggantian.Location = New System.Drawing.Point(12, 35)
        Me.rboPenggantian.Name = "rboPenggantian"
        Me.rboPenggantian.Size = New System.Drawing.Size(85, 17)
        Me.rboPenggantian.TabIndex = 2
        Me.rboPenggantian.Text = "Penggantian"
        Me.rboPenggantian.UseVisualStyleBackColor = True
        '
        'rboTermin
        '
        Me.rboTermin.AutoSize = True
        Me.rboTermin.Location = New System.Drawing.Point(163, 35)
        Me.rboTermin.Name = "rboTermin"
        Me.rboTermin.Size = New System.Drawing.Size(57, 17)
        Me.rboTermin.TabIndex = 4
        Me.rboTermin.Text = "Termin"
        Me.rboTermin.UseVisualStyleBackColor = True
        '
        'rboUangMuka
        '
        Me.rboUangMuka.AutoSize = True
        Me.rboUangMuka.Enabled = False
        Me.rboUangMuka.Location = New System.Drawing.Point(163, 12)
        Me.rboUangMuka.Name = "rboUangMuka"
        Me.rboUangMuka.Size = New System.Drawing.Size(81, 17)
        Me.rboUangMuka.TabIndex = 3
        Me.rboUangMuka.Text = "Uang Muka"
        Me.rboUangMuka.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSalesInvoice_JenisFakturPajakStandar
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(292, 107)
        Me.Controls.Add(Me.rboTermin)
        Me.Controls.Add(Me.rboUangMuka)
        Me.Controls.Add(Me.rboPenggantian)
        Me.Controls.Add(Me.rboHargaJual)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiSalesInvoice_JenisFakturPajakStandar"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Please choose tax invoice type"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents rboHargaJual As System.Windows.Forms.RadioButton
    Friend WithEvents rboPenggantian As System.Windows.Forms.RadioButton
    Friend WithEvents rboTermin As System.Windows.Forms.RadioButton
    Friend WithEvents rboUangMuka As System.Windows.Forms.RadioButton

End Class
