<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPersonalSetting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPersonalSetting))
        Me.GroupBox15 = New System.Windows.Forms.GroupBox
        Me.cboFindBy_Employee = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboFindBy_Supplier = New System.Windows.Forms.ComboBox
        Me.Label52 = New System.Windows.Forms.Label
        Me.cboFindBy_Product = New System.Windows.Forms.ComboBox
        Me.Label51 = New System.Windows.Forms.Label
        Me.cboFindBy_Customer = New System.Windows.Forms.ComboBox
        Me.Label50 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox15.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox15
        '
        Me.GroupBox15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox15.Controls.Add(Me.cboFindBy_Employee)
        Me.GroupBox15.Controls.Add(Me.Label1)
        Me.GroupBox15.Controls.Add(Me.cboFindBy_Supplier)
        Me.GroupBox15.Controls.Add(Me.Label52)
        Me.GroupBox15.Controls.Add(Me.cboFindBy_Product)
        Me.GroupBox15.Controls.Add(Me.Label51)
        Me.GroupBox15.Controls.Add(Me.cboFindBy_Customer)
        Me.GroupBox15.Controls.Add(Me.Label50)
        Me.GroupBox15.Location = New System.Drawing.Point(9, 10)
        Me.GroupBox15.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox15.Name = "GroupBox15"
        Me.GroupBox15.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox15.Size = New System.Drawing.Size(297, 124)
        Me.GroupBox15.TabIndex = 0
        Me.GroupBox15.TabStop = False
        Me.GroupBox15.Text = "Default Find By"
        '
        'cboFindBy_Employee
        '
        Me.cboFindBy_Employee.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFindBy_Employee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFindBy_Employee.FormattingEnabled = True
        Me.cboFindBy_Employee.Items.AddRange(New Object() {"By Global Setting", "By Employee Code", "By Employee Name"})
        Me.cboFindBy_Employee.Location = New System.Drawing.Point(109, 91)
        Me.cboFindBy_Employee.Margin = New System.Windows.Forms.Padding(2)
        Me.cboFindBy_Employee.Name = "cboFindBy_Employee"
        Me.cboFindBy_Employee.Size = New System.Drawing.Size(185, 21)
        Me.cboFindBy_Employee.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Employee"
        '
        'cboFindBy_Supplier
        '
        Me.cboFindBy_Supplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFindBy_Supplier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFindBy_Supplier.FormattingEnabled = True
        Me.cboFindBy_Supplier.Items.AddRange(New Object() {"By Global Setting", "By Supplier Code", "By Supplier Name"})
        Me.cboFindBy_Supplier.Location = New System.Drawing.Point(109, 66)
        Me.cboFindBy_Supplier.Margin = New System.Windows.Forms.Padding(2)
        Me.cboFindBy_Supplier.Name = "cboFindBy_Supplier"
        Me.cboFindBy_Supplier.Size = New System.Drawing.Size(185, 21)
        Me.cboFindBy_Supplier.TabIndex = 2
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Location = New System.Drawing.Point(5, 68)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(45, 13)
        Me.Label52.TabIndex = 5
        Me.Label52.Text = "Supplier"
        '
        'cboFindBy_Product
        '
        Me.cboFindBy_Product.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFindBy_Product.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFindBy_Product.FormattingEnabled = True
        Me.cboFindBy_Product.Items.AddRange(New Object() {"By Global Setting", "By Product Code", "By Product Name"})
        Me.cboFindBy_Product.Location = New System.Drawing.Point(109, 41)
        Me.cboFindBy_Product.Margin = New System.Windows.Forms.Padding(2)
        Me.cboFindBy_Product.Name = "cboFindBy_Product"
        Me.cboFindBy_Product.Size = New System.Drawing.Size(185, 21)
        Me.cboFindBy_Product.TabIndex = 1
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Location = New System.Drawing.Point(5, 44)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(44, 13)
        Me.Label51.TabIndex = 3
        Me.Label51.Text = "Product"
        '
        'cboFindBy_Customer
        '
        Me.cboFindBy_Customer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFindBy_Customer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFindBy_Customer.FormattingEnabled = True
        Me.cboFindBy_Customer.Items.AddRange(New Object() {"By Global Setting", "By Customer Code", "By Customer Name"})
        Me.cboFindBy_Customer.Location = New System.Drawing.Point(109, 17)
        Me.cboFindBy_Customer.Margin = New System.Windows.Forms.Padding(2)
        Me.cboFindBy_Customer.Name = "cboFindBy_Customer"
        Me.cboFindBy_Customer.Size = New System.Drawing.Size(185, 21)
        Me.cboFindBy_Customer.TabIndex = 0
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(5, 20)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(51, 13)
        Me.Label50.TabIndex = 1
        Me.Label50.Text = "Customer"
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(149, 143)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(230, 144)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmPersonalSetting
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(315, 177)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox15)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(325, 189)
        Me.Name = "FrmPersonalSetting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Personal Setting"
        Me.GroupBox15.ResumeLayout(False)
        Me.GroupBox15.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox15 As System.Windows.Forms.GroupBox
    Friend WithEvents cboFindBy_Supplier As System.Windows.Forms.ComboBox
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents cboFindBy_Product As System.Windows.Forms.ComboBox
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents cboFindBy_Customer As System.Windows.Forms.ComboBox
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cboFindBy_Employee As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
