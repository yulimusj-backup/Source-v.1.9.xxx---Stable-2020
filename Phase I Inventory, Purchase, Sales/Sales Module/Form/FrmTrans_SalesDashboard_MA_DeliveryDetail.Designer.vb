<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SalesDashboard_MA_DeliveryDetail
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvData = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriorityLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColQtyIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPriorityLevel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEdit = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPromiseToDeliver = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPromiseToDeliver = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsLate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColBookingNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBookQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Button1 = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvData
        '
        Me.dgvData.AllowUserToAddRows = False
        Me.dgvData.AllowUserToDeleteRows = False
        Me.dgvData.AllowUserToResizeRows = False
        Me.dgvData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDate, Me.ColDesc, Me.ColPriorityLevel, Me.ColQtyIn, Me.ColQtyOut, Me.ColBalance, Me.ColStatus, Me.ColPrimaryPriorityLevel, Me.ColEdit, Me.ColPrimary, Me.ColTipeItem, Me.ColDocNo, Me.ColPromiseToDeliver, Me.ColQtyPromiseToDeliver, Me.ColIsLate, Me.ColBookingNo, Me.ColBookQty})
        Me.dgvData.Location = New System.Drawing.Point(10, 10)
        Me.dgvData.Name = "dgvData"
        Me.dgvData.RowTemplate.Height = 24
        Me.dgvData.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvData.Size = New System.Drawing.Size(754, 355)
        Me.dgvData.TabIndex = 16
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 36
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColPriorityLevel
        '
        Me.ColPriorityLevel.HeaderText = "Priority Level"
        Me.ColPriorityLevel.Name = "ColPriorityLevel"
        Me.ColPriorityLevel.Width = 73
        '
        'ColQtyIn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyIn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyIn.HeaderText = "Qty In"
        Me.ColQtyIn.Name = "ColQtyIn"
        Me.ColQtyIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyIn.Width = 41
        '
        'ColQtyOut
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOut.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyOut.HeaderText = "Qty Out"
        Me.ColQtyOut.Name = "ColQtyOut"
        Me.ColQtyOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyOut.Width = 49
        '
        'ColBalance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBalance.Width = 52
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.Width = 62
        '
        'ColPrimaryPriorityLevel
        '
        Me.ColPrimaryPriorityLevel.HeaderText = "Primary Priority Level"
        Me.ColPrimaryPriorityLevel.Name = "ColPrimaryPriorityLevel"
        Me.ColPrimaryPriorityLevel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPriorityLevel.Visible = False
        Me.ColPrimaryPriorityLevel.Width = 148
        '
        'ColEdit
        '
        Me.ColEdit.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColEdit.HeaderText = "Edit"
        Me.ColEdit.Name = "ColEdit"
        Me.ColEdit.Visible = False
        Me.ColEdit.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColEdit.Width = 38
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 81
        '
        'ColTipeItem
        '
        Me.ColTipeItem.HeaderText = "Tipe Item"
        Me.ColTipeItem.Name = "ColTipeItem"
        Me.ColTipeItem.Visible = False
        Me.ColTipeItem.Width = 91
        '
        'ColDocNo
        '
        Me.ColDocNo.HeaderText = "Document No"
        Me.ColDocNo.Name = "ColDocNo"
        Me.ColDocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDocNo.Width = 79
        '
        'ColPromiseToDeliver
        '
        Me.ColPromiseToDeliver.HeaderText = "Promise To Deliver"
        Me.ColPromiseToDeliver.Name = "ColPromiseToDeliver"
        Me.ColPromiseToDeliver.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPromiseToDeliver.Visible = False
        Me.ColPromiseToDeliver.Width = 134
        '
        'ColQtyPromiseToDeliver
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyPromiseToDeliver.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyPromiseToDeliver.HeaderText = "Qty Promise"
        Me.ColQtyPromiseToDeliver.Name = "ColQtyPromiseToDeliver"
        Me.ColQtyPromiseToDeliver.Visible = False
        Me.ColQtyPromiseToDeliver.Width = 110
        '
        'ColIsLate
        '
        Me.ColIsLate.HeaderText = "Is Late"
        Me.ColIsLate.Name = "ColIsLate"
        Me.ColIsLate.Visible = False
        Me.ColIsLate.Width = 56
        '
        'ColBookingNo
        '
        Me.ColBookingNo.HeaderText = "Inventory Booking"
        Me.ColBookingNo.Name = "ColBookingNo"
        Me.ColBookingNo.Width = 118
        '
        'ColBookQty
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBookQty.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColBookQty.HeaderText = "Book Qty"
        Me.ColBookQty.Name = "ColBookQty"
        Me.ColBookQty.Width = 76
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(522, 141)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 19)
        Me.Button1.TabIndex = 17
        Me.Button1.TabStop = False
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(57, 370)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(83, 17)
        Me.CheckBox1.TabIndex = 18
        Me.CheckBox1.Text = "Show Detail"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.PictureBox1.Location = New System.Drawing.Point(140, 371)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(75, 18)
        Me.PictureBox1.TabIndex = 19
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(219, 372)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(228, 13)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Shows delivery, return, and cancellation history"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(532, 372)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Shows delivery estimation"
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.PictureBox2.BackColor = System.Drawing.Color.White
        Me.PictureBox2.Location = New System.Drawing.Point(452, 371)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(75, 18)
        Me.PictureBox2.TabIndex = 21
        Me.PictureBox2.TabStop = False
        '
        'FrmTrans_SalesDashboard_MA_DeliveryDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(773, 398)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.dgvData)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(693, 436)
        Me.Name = "FrmTrans_SalesDashboard_MA_DeliveryDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Delivery"
        CType(Me.dgvData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvData As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriorityLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColQtyIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPriorityLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEdit As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPromiseToDeliver As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPromiseToDeliver As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsLate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColBookingNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBookQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
End Class
