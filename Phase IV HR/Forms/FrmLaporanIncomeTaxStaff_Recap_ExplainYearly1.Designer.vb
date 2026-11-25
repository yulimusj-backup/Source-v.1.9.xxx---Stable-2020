<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItemDetail = New System.Windows.Forms.DataGridView
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColItem_Period = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ColDetail_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_IsRutin = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDetail_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItemDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItemDetail
        '
        Me.dgvItemDetail.AllowUserToAddRows = False
        Me.dgvItemDetail.AllowUserToDeleteRows = False
        Me.dgvItemDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemDetail.ColumnHeadersHeight = 25
        Me.dgvItemDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetail_No, Me.ColDetail_Desc, Me.ColDetail_IsRutin, Me.ColDetail_Amount})
        Me.dgvItemDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemDetail.Location = New System.Drawing.Point(0, 0)
        Me.dgvItemDetail.Name = "dgvItemDetail"
        Me.dgvItemDetail.ReadOnly = True
        Me.dgvItemDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemDetail.RowTemplate.Height = 24
        Me.dgvItemDetail.Size = New System.Drawing.Size(612, 248)
        Me.dgvItemDetail.TabIndex = 1
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_Period, Me.ColItem_Amount})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(612, 378)
        Me.dgvItem.TabIndex = 2
        '
        'ColItem_Period
        '
        Me.ColItem_Period.HeaderText = "Period"
        Me.ColItem_Period.Name = "ColItem_Period"
        Me.ColItem_Period.ReadOnly = True
        Me.ColItem_Period.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Period.Width = 55
        '
        'ColItem_Amount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Amount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColItem_Amount.HeaderText = "Amount"
        Me.ColItem_Amount.Name = "ColItem_Amount"
        Me.ColItem_Amount.ReadOnly = True
        Me.ColItem_Amount.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColItem_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColItem_Amount.Width = 62
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 58)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvItem)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvItemDetail)
        Me.SplitContainer1.Size = New System.Drawing.Size(612, 630)
        Me.SplitContainer1.SplitterDistance = 378
        Me.SplitContainer1.TabIndex = 3
        Me.SplitContainer1.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(425, 86)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Label2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(266, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Label3"
        '
        'ColDetail_No
        '
        Me.ColDetail_No.HeaderText = "No"
        Me.ColDetail_No.Name = "ColDetail_No"
        Me.ColDetail_No.ReadOnly = True
        Me.ColDetail_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_No.Width = 32
        '
        'ColDetail_Desc
        '
        Me.ColDetail_Desc.HeaderText = "Description"
        Me.ColDetail_Desc.Name = "ColDetail_Desc"
        Me.ColDetail_Desc.ReadOnly = True
        Me.ColDetail_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Desc.Width = 85
        '
        'ColDetail_IsRutin
        '
        Me.ColDetail_IsRutin.HeaderText = "Is Routine"
        Me.ColDetail_IsRutin.Name = "ColDetail_IsRutin"
        Me.ColDetail_IsRutin.ReadOnly = True
        Me.ColDetail_IsRutin.Width = 77
        '
        'ColDetail_Amount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.NullValue = "0"
        Me.ColDetail_Amount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDetail_Amount.HeaderText = "Amount"
        Me.ColDetail_Amount.Name = "ColDetail_Amount"
        Me.ColDetail_Amount.ReadOnly = True
        Me.ColDetail_Amount.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDetail_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Amount.Width = 62
        '
        'FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(632, 700)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(650, 650)
        Me.Name = "FrmLaporanIncomeTaxStaff_Recap_ExplainYearly1"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Annual Amount"
        CType(Me.dgvItemDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItemDetail As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColItem_Period As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ColDetail_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_IsRutin As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDetail_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
