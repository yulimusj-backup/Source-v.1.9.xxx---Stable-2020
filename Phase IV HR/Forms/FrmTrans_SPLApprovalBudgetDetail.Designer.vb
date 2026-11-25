<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPLApprovalBudgetDetail
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColSpl = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProposed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUsed = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtBudget = New System.Windows.Forms.TextBox
        Me.TxtUsed = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.BtnClose = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSpl, Me.ColDate, Me.ColProposed, Me.ColUsed, Me.ColBalance, Me.ColStatus, Me.ColPrimary})
        Me.DgvItem.Location = New System.Drawing.Point(12, 36)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.ReadOnly = True
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.Size = New System.Drawing.Size(505, 214)
        Me.DgvItem.TabIndex = 1
        '
        'ColSpl
        '
        Me.ColSpl.HeaderText = "Overtime Order"
        Me.ColSpl.Name = "ColSpl"
        Me.ColSpl.ReadOnly = True
        Me.ColSpl.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColProposed
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProposed.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColProposed.HeaderText = "Proposed"
        Me.ColProposed.Name = "ColProposed"
        Me.ColProposed.ReadOnly = True
        Me.ColProposed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColUsed
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUsed.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColUsed.HeaderText = "Used"
        Me.ColUsed.Name = "ColUsed"
        Me.ColUsed.ReadOnly = True
        Me.ColUsed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColBalance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        Me.ColPrimary.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Budget"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Used"
        '
        'TxtBudget
        '
        Me.TxtBudget.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBudget.Location = New System.Drawing.Point(64, 3)
        Me.TxtBudget.Name = "TxtBudget"
        Me.TxtBudget.ReadOnly = True
        Me.TxtBudget.Size = New System.Drawing.Size(190, 20)
        Me.TxtBudget.TabIndex = 4
        Me.TxtBudget.Text = "00.00"
        Me.TxtBudget.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtUsed
        '
        Me.TxtUsed.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtUsed.Location = New System.Drawing.Point(58, 2)
        Me.TxtUsed.Name = "TxtUsed"
        Me.TxtUsed.ReadOnly = True
        Me.TxtUsed.Size = New System.Drawing.Size(183, 20)
        Me.TxtUsed.TabIndex = 5
        Me.TxtUsed.Text = "00.00"
        Me.TxtUsed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 6)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtBudget)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtUsed)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Size = New System.Drawing.Size(505, 24)
        Me.SplitContainer1.SplitterDistance = 257
        Me.SplitContainer1.TabIndex = 6
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(442, 227)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 7
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_SPLApprovalBudgetDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(529, 262)
        Me.Controls.Add(Me.DgvItem)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(545, 300)
        Me.Name = "FrmTrans_SPLApprovalBudgetDetail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Approval Detail"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtBudget As System.Windows.Forms.TextBox
    Friend WithEvents TxtUsed As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents ColSpl As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProposed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUsed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
