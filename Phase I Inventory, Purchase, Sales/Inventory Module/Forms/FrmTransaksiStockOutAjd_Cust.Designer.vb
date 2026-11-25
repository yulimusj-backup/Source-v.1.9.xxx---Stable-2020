<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockOutAjd_Cust
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
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtCity = New System.Windows.Forms.TextBox
        Me.txtCustomerName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.cboTipe = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(15, 95)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(24, 13)
        Me.Label21.TabIndex = 188
        Me.Label21.Text = "City"
        '
        'TxtCity
        '
        Me.TxtCity.Location = New System.Drawing.Point(116, 90)
        Me.TxtCity.MaxLength = 50
        Me.TxtCity.Name = "TxtCity"
        Me.TxtCity.Size = New System.Drawing.Size(318, 20)
        Me.TxtCity.TabIndex = 4
        '
        'txtCustomerName
        '
        Me.txtCustomerName.Location = New System.Drawing.Point(115, 64)
        Me.txtCustomerName.MaxLength = 150
        Me.txtCustomerName.Name = "txtCustomerName"
        Me.txtCustomerName.Size = New System.Drawing.Size(319, 20)
        Me.txtCustomerName.TabIndex = 3
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(84, 39)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCustomer.TabIndex = 6
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 185
        Me.Label1.Text = "Customer Code"
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(402, 39)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(32, 20)
        Me.btnFindCustomer.TabIndex = 2
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "...."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 187
        Me.Label2.Text = "Customer Name"
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(115, 39)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(281, 20)
        Me.txtCustomerCode.TabIndex = 1
        '
        'cboTipe
        '
        Me.cboTipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipe.FormattingEnabled = True
        Me.cboTipe.Items.AddRange(New Object() {"Replacement", "Loss", "Adjustment", "Sample"})
        Me.cboTipe.Location = New System.Drawing.Point(116, 12)
        Me.cboTipe.Name = "cboTipe"
        Me.cboTipe.Size = New System.Drawing.Size(318, 21)
        Me.cboTipe.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(28, 13)
        Me.Label3.TabIndex = 190
        Me.Label3.Text = "Tipe"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(349, 116)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(85, 20)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockOutAjd_Cust
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(446, 143)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboTipe)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.TxtCity)
        Me.Controls.Add(Me.txtCustomerName)
        Me.Controls.Add(Me.txtPrimaryCustomer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnFindCustomer)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtCustomerCode)
        Me.Name = "FrmTransaksiStockOutAjd_Cust"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents cboTipe As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
