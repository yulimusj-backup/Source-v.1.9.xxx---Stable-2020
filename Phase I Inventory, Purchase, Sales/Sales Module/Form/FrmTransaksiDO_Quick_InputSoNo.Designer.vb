<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO_Quick_InputSoNo
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerSoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCustomerCode, Me.ColCustomerName, Me.ColCustomerSoNo, Me.ColPrimaryCustomer})
        Me.DgvItem.Location = New System.Drawing.Point(9, 10)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(494, 261)
        Me.DgvItem.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.FillWeight = 50.0!
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ColNo.Width = 50
        '
        'ColCustomerCode
        '
        Me.ColCustomerCode.HeaderText = "Customer Code"
        Me.ColCustomerCode.Name = "ColCustomerCode"
        Me.ColCustomerCode.ReadOnly = True
        '
        'ColCustomerName
        '
        Me.ColCustomerName.HeaderText = "Customer Name"
        Me.ColCustomerName.Name = "ColCustomerName"
        Me.ColCustomerName.ReadOnly = True
        '
        'ColCustomerSoNo
        '
        Me.ColCustomerSoNo.HeaderText = "Customer Do No"
        Me.ColCustomerSoNo.Name = "ColCustomerSoNo"
        '
        'ColPrimaryCustomer
        '
        Me.ColPrimaryCustomer.HeaderText = "Primary Customer"
        Me.ColPrimaryCustomer.Name = "ColPrimaryCustomer"
        Me.ColPrimaryCustomer.Visible = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(437, 276)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 24)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransaksiDO_Quick_InputSoNo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 309)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DgvItem)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiDO_Quick_InputSoNo"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input Customer SO No"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerSoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
