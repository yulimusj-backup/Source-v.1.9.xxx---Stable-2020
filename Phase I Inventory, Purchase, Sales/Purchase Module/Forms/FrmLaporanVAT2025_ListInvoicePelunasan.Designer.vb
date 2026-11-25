<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanVAT2025_ListInvoicePelunasan
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.col_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_CustomerCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_CustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_NoTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_TanggalTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_NoDP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_TaxInvoiceDP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_TanggalDP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_SubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_VAT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_Total = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_PrimaryTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col_PrimaryDP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_PrimaryCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(552, 362)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(546, 40)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pada period ini terdapat transaksi invoice pelunasan yang tidak dapat dibuatkan i" & _
            "mport coretax nya, berikut adalah daftar invoice pelunasan : "
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col_No, Me.col_CustomerCode, Me.col_CustomerName, Me.col_NoTransaksi, Me.col_TanggalTransaksi, Me.col_NoDP, Me.col_TaxInvoiceDP, Me.col_TanggalDP, Me.col_SubTotal, Me.col_VAT, Me.col_Total, Me.col_PrimaryTransaksi, Me.col_PrimaryDP, Me.Col_PrimaryCustomer})
        Me.dgvItem.Location = New System.Drawing.Point(6, 59)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(541, 296)
        Me.dgvItem.TabIndex = 0
        '
        'col_No
        '
        Me.col_No.HeaderText = "No"
        Me.col_No.Name = "col_No"
        '
        'col_CustomerCode
        '
        Me.col_CustomerCode.HeaderText = "Customer Code"
        Me.col_CustomerCode.Name = "col_CustomerCode"
        '
        'col_CustomerName
        '
        Me.col_CustomerName.HeaderText = "Customer Name"
        Me.col_CustomerName.Name = "col_CustomerName"
        '
        'col_NoTransaksi
        '
        Me.col_NoTransaksi.HeaderText = "Invoice No"
        Me.col_NoTransaksi.Name = "col_NoTransaksi"
        '
        'col_TanggalTransaksi
        '
        Me.col_TanggalTransaksi.HeaderText = "Invoice Date"
        Me.col_TanggalTransaksi.Name = "col_TanggalTransaksi"
        '
        'col_NoDP
        '
        Me.col_NoDP.HeaderText = "Down Payment No"
        Me.col_NoDP.Name = "col_NoDP"
        '
        'col_TaxInvoiceDP
        '
        Me.col_TaxInvoiceDP.HeaderText = "Down Payment Tax No"
        Me.col_TaxInvoiceDP.Name = "col_TaxInvoiceDP"
        '
        'col_TanggalDP
        '
        Me.col_TanggalDP.HeaderText = "Down Payment Date"
        Me.col_TanggalDP.Name = "col_TanggalDP"
        '
        'col_SubTotal
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.col_SubTotal.DefaultCellStyle = DataGridViewCellStyle4
        Me.col_SubTotal.HeaderText = "Sub Total"
        Me.col_SubTotal.Name = "col_SubTotal"
        '
        'col_VAT
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.col_VAT.DefaultCellStyle = DataGridViewCellStyle5
        Me.col_VAT.HeaderText = "VAT"
        Me.col_VAT.Name = "col_VAT"
        '
        'col_Total
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.col_Total.DefaultCellStyle = DataGridViewCellStyle6
        Me.col_Total.HeaderText = "Total"
        Me.col_Total.Name = "col_Total"
        '
        'col_PrimaryTransaksi
        '
        Me.col_PrimaryTransaksi.HeaderText = "Primary Transaksi"
        Me.col_PrimaryTransaksi.Name = "col_PrimaryTransaksi"
        Me.col_PrimaryTransaksi.Visible = False
        '
        'col_PrimaryDP
        '
        Me.col_PrimaryDP.HeaderText = "Primary DP"
        Me.col_PrimaryDP.Name = "col_PrimaryDP"
        Me.col_PrimaryDP.Visible = False
        '
        'Col_PrimaryCustomer
        '
        Me.Col_PrimaryCustomer.HeaderText = "Primary Customer"
        Me.Col_PrimaryCustomer.Name = "Col_PrimaryCustomer"
        Me.Col_PrimaryCustomer.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(489, 380)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'FrmLaporanVAT2025_ListInvoicePelunasan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(576, 421)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(584, 448)
        Me.Name = "FrmLaporanVAT2025_ListInvoicePelunasan"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "List Invoice dengan DP"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents col_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_CustomerCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_CustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_NoTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_TanggalTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_NoDP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_TaxInvoiceDP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_TanggalDP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_SubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_VAT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_Total As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_PrimaryTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col_PrimaryDP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_PrimaryCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
