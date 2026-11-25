<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_Explain2
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
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoTrans = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_Sales = New System.Windows.Forms.TabPage
        Me.TabPage_SalesReturn = New System.Windows.Forms.TabPage
        Me.dgvDataSR = New System.Windows.Forms.DataGridView
        Me.ColSR_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_Expiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSR_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_Sales.SuspendLayout()
        Me.TabPage_SalesReturn.SuspendLayout()
        CType(Me.dgvDataSR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToResizeColumns = False
        Me.dgvData.AllowUserToResizeRows = False
        Me.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDate, Me.ColNoTrans, Me.ColCustomer, Me.ColBatchNo, Me.ColExpiry, Me.ColQty, Me.ColUnit})
        Me.dgvData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvData.Location = New System.Drawing.Point(3, 3)
        Me.dgvData.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.Size = New System.Drawing.Size(913, 456)
        Me.dgvData.TabIndex = 4
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 63
        '
        'ColNoTrans
        '
        Me.ColNoTrans.HeaderText = "Transaction No"
        Me.ColNoTrans.Name = "ColNoTrans"
        Me.ColNoTrans.Width = 130
        '
        'ColCustomer
        '
        Me.ColCustomer.HeaderText = "Customer"
        Me.ColCustomer.Name = "ColCustomer"
        Me.ColCustomer.Width = 93
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.Width = 91
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.Width = 71
        '
        'ColQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 55
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 58
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(575, 78)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_Sales)
        Me.TabControl1.Controls.Add(Me.TabPage_SalesReturn)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(927, 491)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage_Sales
        '
        Me.TabPage_Sales.Controls.Add(Me.dgvData)
        Me.TabPage_Sales.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_Sales.Name = "TabPage_Sales"
        Me.TabPage_Sales.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Sales.Size = New System.Drawing.Size(919, 462)
        Me.TabPage_Sales.TabIndex = 0
        Me.TabPage_Sales.Text = "Sales"
        Me.TabPage_Sales.UseVisualStyleBackColor = True
        '
        'TabPage_SalesReturn
        '
        Me.TabPage_SalesReturn.Controls.Add(Me.dgvDataSR)
        Me.TabPage_SalesReturn.Location = New System.Drawing.Point(4, 25)
        Me.TabPage_SalesReturn.Name = "TabPage_SalesReturn"
        Me.TabPage_SalesReturn.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_SalesReturn.Size = New System.Drawing.Size(919, 462)
        Me.TabPage_SalesReturn.TabIndex = 1
        Me.TabPage_SalesReturn.Text = "Sales Return"
        Me.TabPage_SalesReturn.UseVisualStyleBackColor = True
        '
        'dgvDataSR
        '
        Me.dgvDataSR.AllowUserToAddRows = False
        Me.dgvDataSR.AllowUserToDeleteRows = False
        Me.dgvDataSR.AllowUserToResizeColumns = False
        Me.dgvDataSR.AllowUserToResizeRows = False
        Me.dgvDataSR.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDataSR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDataSR.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSR_No, Me.ColSR_Date, Me.ColSR_DocumentNo, Me.ColSR_Customer, Me.ColSR_BatchNo, Me.ColSR_Expiry, Me.ColSR_Qty, Me.ColSR_Unit})
        Me.dgvDataSR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDataSR.Location = New System.Drawing.Point(3, 3)
        Me.dgvDataSR.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvDataSR.Name = "dgvDataSR"
        Me.dgvDataSR.RowTemplate.Height = 24
        Me.dgvDataSR.Size = New System.Drawing.Size(913, 456)
        Me.dgvDataSR.TabIndex = 5
        '
        'ColSR_No
        '
        Me.ColSR_No.HeaderText = "No"
        Me.ColSR_No.Name = "ColSR_No"
        Me.ColSR_No.Width = 51
        '
        'ColSR_Date
        '
        Me.ColSR_Date.HeaderText = "Date"
        Me.ColSR_Date.Name = "ColSR_Date"
        Me.ColSR_Date.Width = 63
        '
        'ColSR_DocumentNo
        '
        Me.ColSR_DocumentNo.HeaderText = "Transaction No"
        Me.ColSR_DocumentNo.Name = "ColSR_DocumentNo"
        Me.ColSR_DocumentNo.Width = 130
        '
        'ColSR_Customer
        '
        Me.ColSR_Customer.HeaderText = "Customer"
        Me.ColSR_Customer.Name = "ColSR_Customer"
        Me.ColSR_Customer.Width = 93
        '
        'ColSR_BatchNo
        '
        Me.ColSR_BatchNo.HeaderText = "Batch No"
        Me.ColSR_BatchNo.Name = "ColSR_BatchNo"
        Me.ColSR_BatchNo.Width = 91
        '
        'ColSR_Expiry
        '
        Me.ColSR_Expiry.HeaderText = "Expiry"
        Me.ColSR_Expiry.Name = "ColSR_Expiry"
        Me.ColSR_Expiry.Width = 71
        '
        'ColSR_Qty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSR_Qty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSR_Qty.HeaderText = "Qty"
        Me.ColSR_Qty.Name = "ColSR_Qty"
        Me.ColSR_Qty.Width = 55
        '
        'ColSR_Unit
        '
        Me.ColSR_Unit.HeaderText = "Unit"
        Me.ColSR_Unit.Name = "ColSR_Unit"
        Me.ColSR_Unit.Width = 58
        '
        'FrmTrans_MaterialPlanningDashboard_Explain2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(927, 491)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(935, 528)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_Explain2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Explaination"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_Sales.ResumeLayout(False)
        Me.TabPage_SalesReturn.ResumeLayout(False)
        CType(Me.dgvDataSR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoTrans As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_Sales As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_SalesReturn As System.Windows.Forms.TabPage
    Friend WithEvents dgvDataSR As System.Windows.Forms.DataGridView
    Friend WithEvents ColSR_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_Expiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSR_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
