<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2_EndAmount2
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColAcc_AccountCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_AmountClosing = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAcc_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvDetailItem = New System.Windows.Forms.DataGridView
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnRepair = New System.Windows.Forms.Button
        Me.ColItem_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_RefAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_ClosingAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_ClosingAmountDetail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItem_StatusRepair = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDetailItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(968, 580)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.SplitContainer1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(960, 551)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "By Account"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvItem)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgvDetailItem)
        Me.SplitContainer1.Size = New System.Drawing.Size(954, 545)
        Me.SplitContainer1.SplitterDistance = 212
        Me.SplitContainer1.TabIndex = 5
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAcc_AccountCode, Me.ColAcc_AccDesc, Me.ColAcc_Amount, Me.ColAcc_AmountClosing, Me.ColAcc_PrimaryAkun})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(954, 212)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 5
        '
        'ColAcc_AccountCode
        '
        Me.ColAcc_AccountCode.HeaderText = "Account"
        Me.ColAcc_AccountCode.MaxInputLength = 50
        Me.ColAcc_AccountCode.Name = "ColAcc_AccountCode"
        Me.ColAcc_AccountCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_AccountCode.Width = 65
        '
        'ColAcc_AccDesc
        '
        Me.ColAcc_AccDesc.HeaderText = "Description"
        Me.ColAcc_AccDesc.MaxInputLength = 150
        Me.ColAcc_AccDesc.Name = "ColAcc_AccDesc"
        Me.ColAcc_AccDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_AccDesc.Width = 85
        '
        'ColAcc_Amount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAcc_Amount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColAcc_Amount.HeaderText = "Amount Ref"
        Me.ColAcc_Amount.Name = "ColAcc_Amount"
        Me.ColAcc_Amount.Width = 107
        '
        'ColAcc_AmountClosing
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAcc_AmountClosing.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColAcc_AmountClosing.HeaderText = "Amount Closing"
        Me.ColAcc_AmountClosing.Name = "ColAcc_AmountClosing"
        Me.ColAcc_AmountClosing.Width = 131
        '
        'ColAcc_PrimaryAkun
        '
        Me.ColAcc_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColAcc_PrimaryAkun.Name = "ColAcc_PrimaryAkun"
        Me.ColAcc_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAcc_PrimaryAkun.Visible = False
        Me.ColAcc_PrimaryAkun.Width = 94
        '
        'dgvDetailItem
        '
        Me.dgvDetailItem.AllowUserToAddRows = False
        Me.dgvDetailItem.AllowUserToResizeColumns = False
        Me.dgvDetailItem.AllowUserToResizeRows = False
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvDetailItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvDetailItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetailItem.ColumnHeadersHeight = 25
        Me.dgvDetailItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetailItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItem_CodeNo, Me.ColItem_Desc, Me.ColItem_Qty, Me.ColItem_Unit, Me.ColItem_RefAmount, Me.ColItem_ClosingAmount, Me.ColItem_ClosingAmountDetail, Me.ColItem_StatusRepair})
        Me.dgvDetailItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDetailItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvDetailItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvDetailItem.Name = "dgvDetailItem"
        Me.dgvDetailItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetailItem.RowTemplate.Height = 24
        Me.dgvDetailItem.Size = New System.Drawing.Size(954, 329)
        Me.dgvDetailItem.StandardTab = True
        Me.dgvDetailItem.TabIndex = 6
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(960, 551)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "By Product"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(12, 598)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 25)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnRepair
        '
        Me.btnRepair.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRepair.Location = New System.Drawing.Point(93, 598)
        Me.btnRepair.Name = "btnRepair"
        Me.btnRepair.Size = New System.Drawing.Size(75, 25)
        Me.btnRepair.TabIndex = 2
        Me.btnRepair.Text = "Repair"
        Me.btnRepair.UseVisualStyleBackColor = True
        '
        'ColItem_CodeNo
        '
        Me.ColItem_CodeNo.HeaderText = "Code No"
        Me.ColItem_CodeNo.Name = "ColItem_CodeNo"
        Me.ColItem_CodeNo.Width = 88
        '
        'ColItem_Desc
        '
        Me.ColItem_Desc.HeaderText = "Description"
        Me.ColItem_Desc.Name = "ColItem_Desc"
        Me.ColItem_Desc.Width = 104
        '
        'ColItem_Qty
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_Qty.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColItem_Qty.HeaderText = "Qty"
        Me.ColItem_Qty.Name = "ColItem_Qty"
        Me.ColItem_Qty.Width = 55
        '
        'ColItem_Unit
        '
        Me.ColItem_Unit.HeaderText = "Unit"
        Me.ColItem_Unit.Name = "ColItem_Unit"
        Me.ColItem_Unit.Width = 58
        '
        'ColItem_RefAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_RefAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColItem_RefAmount.HeaderText = "Ref. Amount"
        Me.ColItem_RefAmount.Name = "ColItem_RefAmount"
        Me.ColItem_RefAmount.Width = 111
        '
        'ColItem_ClosingAmount
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItem_ClosingAmount.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColItem_ClosingAmount.HeaderText = "Closing Amount"
        Me.ColItem_ClosingAmount.Name = "ColItem_ClosingAmount"
        Me.ColItem_ClosingAmount.Width = 131
        '
        'ColItem_ClosingAmountDetail
        '
        Me.ColItem_ClosingAmountDetail.HeaderText = "Closing Amount Detail"
        Me.ColItem_ClosingAmountDetail.Name = "ColItem_ClosingAmountDetail"
        Me.ColItem_ClosingAmountDetail.Width = 171
        '
        'ColItem_StatusRepair
        '
        Me.ColItem_StatusRepair.HeaderText = "Repair"
        Me.ColItem_StatusRepair.Name = "ColItem_StatusRepair"
        Me.ColItem_StatusRepair.Width = 56
        '
        'Form2_EndAmount2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.btnRepair)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.TabControl1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "Form2_EndAmount2"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form2_EndAmount2"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDetailItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColAcc_AccountCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_AmountClosing As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAcc_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents dgvDetailItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnRepair As System.Windows.Forms.Button
    Friend WithEvents ColItem_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_RefAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_ClosingAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_ClosingAmountDetail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItem_StatusRepair As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
