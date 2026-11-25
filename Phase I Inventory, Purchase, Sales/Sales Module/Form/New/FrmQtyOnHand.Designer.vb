<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmQtyOnHand
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.dgvHeader = New System.Windows.Forms.DataGridView
        Me.ColH_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColH_InvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColH_InvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColH_TaxInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColH_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNoSMESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvItem)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvHeader)
        Me.SplitContainer1.Size = New System.Drawing.Size(741, 406)
        Me.SplitContainer1.SplitterDistance = 192
        Me.SplitContainer1.TabIndex = 0
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColCodeNoSMESS, Me.ColDesc, Me.ColQtyOn, Me.ColQty, Me.ColUnit, Me.ColPrimary})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(741, 192)
        Me.dgvItem.TabIndex = 1
        '
        'dgvHeader
        '
        Me.dgvHeader.AllowUserToAddRows = False
        Me.dgvHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColH_Customer, Me.ColH_InvoiceNo, Me.ColH_InvoiceDate, Me.ColH_TaxInvoiceNo, Me.ColH_Primary})
        Me.dgvHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHeader.Location = New System.Drawing.Point(0, 0)
        Me.dgvHeader.Name = "dgvHeader"
        Me.dgvHeader.ReadOnly = True
        Me.dgvHeader.RowTemplate.Height = 24
        Me.dgvHeader.Size = New System.Drawing.Size(741, 210)
        Me.dgvHeader.TabIndex = 0
        '
        'ColH_Customer
        '
        Me.ColH_Customer.HeaderText = "Customer"
        Me.ColH_Customer.Name = "ColH_Customer"
        Me.ColH_Customer.ReadOnly = True
        Me.ColH_Customer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColH_InvoiceNo
        '
        Me.ColH_InvoiceNo.HeaderText = "Invoice No"
        Me.ColH_InvoiceNo.Name = "ColH_InvoiceNo"
        Me.ColH_InvoiceNo.ReadOnly = True
        Me.ColH_InvoiceNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColH_InvoiceDate
        '
        Me.ColH_InvoiceDate.HeaderText = "Date"
        Me.ColH_InvoiceDate.Name = "ColH_InvoiceDate"
        Me.ColH_InvoiceDate.ReadOnly = True
        Me.ColH_InvoiceDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColH_TaxInvoiceNo
        '
        Me.ColH_TaxInvoiceNo.HeaderText = "Tax Invoice No"
        Me.ColH_TaxInvoiceNo.Name = "ColH_TaxInvoiceNo"
        Me.ColH_TaxInvoiceNo.ReadOnly = True
        Me.ColH_TaxInvoiceNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColH_Primary
        '
        Me.ColH_Primary.HeaderText = "Primary"
        Me.ColH_Primary.Name = "ColH_Primary"
        Me.ColH_Primary.ReadOnly = True
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code No Custom"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.ReadOnly = True
        Me.ColCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColCodeNoSMESS
        '
        Me.ColCodeNoSMESS.HeaderText = "Code No SMESS"
        Me.ColCodeNoSMESS.Name = "ColCodeNoSMESS"
        Me.ColCodeNoSMESS.ReadOnly = True
        Me.ColCodeNoSMESS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Visible = False
        '
        'ColQtyOn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyOn.HeaderText = "Qty On Hand"
        Me.ColQtyOn.Name = "ColQtyOn"
        Me.ColQtyOn.ReadOnly = True
        Me.ColQtyOn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        '
        'FrmQtyOnHand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(741, 406)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmQtyOnHand"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Information"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents dgvHeader As System.Windows.Forms.DataGridView
    Friend WithEvents ColH_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColH_InvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColH_InvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColH_TaxInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColH_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNoSMESS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
