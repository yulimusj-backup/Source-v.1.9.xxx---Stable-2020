<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmA_CekRounding
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
        Me.TxtDecimal = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboRoundingTipe = New System.Windows.Forms.ComboBox
        Me.TxtBaseAmount = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtDiscount = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtDiscountAmount = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtNetAmount = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboRoundingValue = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Cetak = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtDecimal
        '
        Me.TxtDecimal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDecimal.Location = New System.Drawing.Point(95, 63)
        Me.TxtDecimal.MaxLength = 150
        Me.TxtDecimal.Name = "TxtDecimal"
        Me.TxtDecimal.Size = New System.Drawing.Size(177, 20)
        Me.TxtDecimal.TabIndex = 2
        Me.TxtDecimal.Text = "0"
        Me.TxtDecimal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 63)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 13)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Decimal"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Rounding Tipe"
        '
        'cboRoundingTipe
        '
        Me.cboRoundingTipe.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboRoundingTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTipe.FormattingEnabled = True
        Me.cboRoundingTipe.Items.AddRange(New Object() {"Up", "Down", "Middle"})
        Me.cboRoundingTipe.Location = New System.Drawing.Point(95, 6)
        Me.cboRoundingTipe.Name = "cboRoundingTipe"
        Me.cboRoundingTipe.Size = New System.Drawing.Size(177, 21)
        Me.cboRoundingTipe.TabIndex = 0
        '
        'TxtBaseAmount
        '
        Me.TxtBaseAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBaseAmount.Location = New System.Drawing.Point(95, 89)
        Me.TxtBaseAmount.MaxLength = 12
        Me.TxtBaseAmount.Name = "TxtBaseAmount"
        Me.TxtBaseAmount.Size = New System.Drawing.Size(177, 20)
        Me.TxtBaseAmount.TabIndex = 3
        Me.TxtBaseAmount.Text = "0"
        Me.TxtBaseAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Base Amount"
        '
        'TxtDiscount
        '
        Me.TxtDiscount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDiscount.Location = New System.Drawing.Point(95, 115)
        Me.TxtDiscount.MaxLength = 150
        Me.TxtDiscount.Name = "TxtDiscount"
        Me.TxtDiscount.Size = New System.Drawing.Size(177, 20)
        Me.TxtDiscount.TabIndex = 4
        Me.TxtDiscount.Text = "0"
        Me.TxtDiscount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(28, 13)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Disc"
        '
        'TxtDiscountAmount
        '
        Me.TxtDiscountAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDiscountAmount.Location = New System.Drawing.Point(95, 141)
        Me.TxtDiscountAmount.MaxLength = 150
        Me.TxtDiscountAmount.Name = "TxtDiscountAmount"
        Me.TxtDiscountAmount.ReadOnly = True
        Me.TxtDiscountAmount.Size = New System.Drawing.Size(177, 20)
        Me.TxtDiscountAmount.TabIndex = 5
        Me.TxtDiscountAmount.TabStop = False
        Me.TxtDiscountAmount.Text = "0"
        Me.TxtDiscountAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 141)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Disc Amount"
        '
        'TxtNetAmount
        '
        Me.TxtNetAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNetAmount.Location = New System.Drawing.Point(95, 167)
        Me.TxtNetAmount.MaxLength = 150
        Me.TxtNetAmount.Name = "TxtNetAmount"
        Me.TxtNetAmount.ReadOnly = True
        Me.TxtNetAmount.Size = New System.Drawing.Size(177, 20)
        Me.TxtNetAmount.TabIndex = 6
        Me.TxtNetAmount.TabStop = False
        Me.TxtNetAmount.Text = "0"
        Me.TxtNetAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 167)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 47
        Me.Label5.Text = "Nett Amount"
        '
        'cboRoundingValue
        '
        Me.cboRoundingValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboRoundingValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingValue.FormattingEnabled = True
        Me.cboRoundingValue.Items.AddRange(New Object() {"1000", "500", "100", "50", "10", "5", "1", "0.5", "0.1", "0.05", "0.01", "0.005", "0.001"})
        Me.cboRoundingValue.Location = New System.Drawing.Point(95, 33)
        Me.cboRoundingValue.Name = "cboRoundingValue"
        Me.cboRoundingValue.Size = New System.Drawing.Size(177, 21)
        Me.cboRoundingValue.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 13)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Rounding Value"
        '
        'Cetak
        '
        Me.Cetak.Location = New System.Drawing.Point(197, 227)
        Me.Cetak.Name = "Cetak"
        Me.Cetak.Size = New System.Drawing.Size(75, 23)
        Me.Cetak.TabIndex = 50
        Me.Cetak.Text = "Cetak"
        Me.Cetak.UseVisualStyleBackColor = True
        '
        'FrmA_CekRounding
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.Cetak)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cboRoundingValue)
        Me.Controls.Add(Me.TxtNetAmount)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtDiscountAmount)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtDiscount)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtBaseAmount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboRoundingTipe)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtDecimal)
        Me.Controls.Add(Me.Label7)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 300)
        Me.Name = "FrmA_CekRounding"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pengecekan Format Discount"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtDecimal As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboRoundingTipe As System.Windows.Forms.ComboBox
    Friend WithEvents TxtBaseAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscount As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtDiscountAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtNetAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboRoundingValue As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Cetak As System.Windows.Forms.Button
End Class
