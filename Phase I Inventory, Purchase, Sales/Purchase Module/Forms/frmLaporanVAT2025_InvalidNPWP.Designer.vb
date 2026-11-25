<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaporanVAT2025_InvalidNPWP
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.dgvInvalidNPWP = New System.Windows.Forms.DataGridView
        Me.ColInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEdit = New System.Windows.Forms.DataGridViewButtonColumn
        Me.colPrimaryCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryInvoice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvInvalidNPWP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnOK)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.dgvInvalidNPWP)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(646, 313)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(479, 280)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        Me.btnOK.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(560, 280)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dgvInvalidNPWP
        '
        Me.dgvInvalidNPWP.AllowUserToAddRows = False
        Me.dgvInvalidNPWP.AllowUserToDeleteRows = False
        Me.dgvInvalidNPWP.AllowUserToResizeRows = False
        Me.dgvInvalidNPWP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvInvalidNPWP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvInvalidNPWP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColInvoiceNo, Me.ColInvoiceDate, Me.ColCustomerName, Me.ColDocType, Me.ColDocNo, Me.ColEdit, Me.colPrimaryCustomer, Me.colPrimaryInvoice})
        Me.dgvInvalidNPWP.Location = New System.Drawing.Point(9, 32)
        Me.dgvInvalidNPWP.Name = "dgvInvalidNPWP"
        Me.dgvInvalidNPWP.Size = New System.Drawing.Size(631, 243)
        Me.dgvInvalidNPWP.TabIndex = 0
        '
        'ColInvoiceNo
        '
        Me.ColInvoiceNo.HeaderText = "Invoice No"
        Me.ColInvoiceNo.Name = "ColInvoiceNo"
        '
        'ColInvoiceDate
        '
        Me.ColInvoiceDate.HeaderText = "Date"
        Me.ColInvoiceDate.Name = "ColInvoiceDate"
        '
        'ColCustomerName
        '
        Me.ColCustomerName.HeaderText = "Customer"
        Me.ColCustomerName.Name = "ColCustomerName"
        '
        'ColDocType
        '
        Me.ColDocType.HeaderText = "Document Type"
        Me.ColDocType.Items.AddRange(New Object() {"NPWP", "NIK", "Passport", "Other ID"})
        Me.ColDocType.Name = "ColDocType"
        '
        'ColDocNo
        '
        Me.ColDocNo.HeaderText = "Document No"
        Me.ColDocNo.Name = "ColDocNo"
        Me.ColDocNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColEdit
        '
        Me.ColEdit.HeaderText = "Edit Data"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.Width = 75
        '
        'colPrimaryCustomer
        '
        Me.colPrimaryCustomer.HeaderText = "Primary Customer"
        Me.colPrimaryCustomer.Name = "colPrimaryCustomer"
        Me.colPrimaryCustomer.Visible = False
        '
        'colPrimaryInvoice
        '
        Me.colPrimaryInvoice.HeaderText = "Primary Invoice"
        Me.colPrimaryInvoice.Name = "colPrimaryInvoice"
        Me.colPrimaryInvoice.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(468, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Based On Government Regulation No.: PMK 112/PMK.03/2022. NPWP No. Contains of 16 " & _
            "Digit."
        '
        'frmLaporanVAT2025_InvalidNPWP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(650, 321)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(666, 360)
        Me.Name = "frmLaporanVAT2025_InvalidNPWP"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data For Invalid NPWP"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvInvalidNPWP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvInvalidNPWP As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEdit As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents colPrimaryCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryInvoice As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
